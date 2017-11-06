VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Post"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Post"

'VGC 10/23/2017: CHANGES!
' ** TimerInterval: 100

' ** For GTR, either shows arrow, sets focus and stops, or shows arrows and continues.

' ** Shortcut Alt keys responsive from this form:
' **   Accrued Interest: A {cmdAccrue}
' **   Close:            C {cmdClose}
' **   Load Timer:       D {chkLoadTimer}
' **   Clear Trans:      E {cmdClearPostingJournal}
' **   Fee Calc:         F {cmdFeeCalculations}
' **   Print Checks:     H {cmdChecksPrint}
' **   Check Reconcile:  K {cmdCheckReconcile}
' **   Preview Checks:   L {cmdChecksPreview}
' **   Check Memos:      M {cmdCheckMemos}
' **   Print Journal:    N {cmdPostingJournalPrint}
' **   POSPay File:      O {cmdPOSPay}
' **   Post Trans:       P {cmdPost}
' **   Clear Checks:     R {cmdChecksClear}
' **   Sweep Cash:       S {cmdSweep}
' **   Trans Maint:      T {cmdTransaction}
' **   Preview Journal:  W {cmdPostingJournalPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Force Date:       T {blnForce: cmdTransaction}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

' ** Pre-Form Processing:
' **   cmdTransaction:
' **     Presence of recs in Journal, tblJournal_Column
' **     Posting Date
' **   cmdPostingJournalPreview:
' **     Presence of recs in Journal
' **   cmdPostingJournalPrint:
' **     Presence of recs in Journal
' **   cmdPost:
' **     Presence of recs in Journal
' **   cmdClearPostingJournal:
' **     Whole procedure
' **   cmdChecksPrint:
' **     {Nothing}
' **   cmdChecksPreview:
' **     Whole procedure
' **   cmdChecksClear:
' **     Whole procedure
' **   cmdCheckReconcile:
' **
' **   cmdSweep:
' **     {Nothing}
' **   cmdAccrue:
' **     Empties temp table
' **   cmdFeeCalculations:
' **     Posting Date

Private lngOpts As Long, arr_varOpt() As Variant
Private Const OPT_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const OPT_NAM As Integer = 0
Private Const OPT_VAL As Integer = 1

Private msgResponse As VbMsgBoxResult
Private intStyle As Integer
Private strTitle As String
Private strMsg As String
Private strDocName As String

Private lngTpp As Long, blnChecksCalled As Boolean
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnGoingToReport2 As Boolean, strMenuArrow As String
Private blnJustClose As Boolean, blnUpdateTransCnt As Boolean, blnFormOpening As Boolean, blnForce As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim grp As DAO.Group, usr As DAO.User
        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           .Painting = True
150           .Repaint

160           strCallingForm = Trim(.OpenArgs)

170           gblnSignal = False

180           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

190           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

              ' ** Loop through all the security groups, checking if our current user is in the admins group.
              ' ** grp.Name = "ViewOnly" already prevented from seeing this menu.
200           gblnAdmin = False
210           For Each grp In DBEngine.Workspaces(0).Groups
220             If grp.Name = "Admins" Then
230               For Each usr In grp.Users
240                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
250                   gblnAdmin = True
260                   Exit For
270                 End If
280               Next
290             End If
300           Next

310           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
320           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

330           blnJustClose = False: blnChecksCalled = False: strMenuArrow = vbNullString
340           blnGTR_Emblem = False: lngGTR_Stat = 0&
350           blnFormOpening = False: blnUpdateTransCnt = False: blnForce = False

360           If gblnGoToReport = True Then
                ' ** Let's see if we can get the arrow to show up more quickly.
370             If IsNull(garr_varGoToReport(GTR_FRM2)) = True And garr_varGoToReport(GTR_RNAM) <> "rptPostingJournal_Column" Then
                  ' ** No more forms.
380               lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
390               lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
400               Select Case lngBtn_Left
                  Case .cmdTransaction.Left
410                 .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
420                 .GoToReport_arw_menu_l_img.Visible = True
430               Case .cmdChecksPreview.Left
440                 .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
450                 .GoToReport_arw_menu_r_img.Visible = True
460               End Select
470               DoEvents
480             Else
                  ' ** Moving on.
490               .GoToReport_Emblem_01_img.Visible = True
500               DoEvents
510               lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
520               lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
530               Select Case lngBtn_Left
                  Case .cmdTransaction.Left
540                 .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
550                 .GoToReport_arw_menu_l_img.Visible = True
560               Case .cmdChecksPreview.Left
570                 .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
580                 .GoToReport_arw_menu_r_img.Visible = True
590               End Select
600               DoEvents
610             End If
620           End If

630           UpdateTransCnt  ' ** Function: Below.

640           gblnSingleUser = True
650           gstrFormQuerySpec = THIS_NAME

              ' ** This is also included in JrnlDateCheck(), next line,
              ' ** but it only hits under specific circumstances.
660           ChkCalStage  ' ** Module Procedure:  modXAdminFuncs.

              ' ** Heavy-duty journal date check!
670           JrnlDateCheck  ' ** Module Function:  modXAdminFuncs.

680           Select Case .chkLoadTimer_vis
              Case True
690             .chkLoadTimer.Visible = True
700             .chkLoadTimer_lbl1.Visible = True
710           Case False
720             .chkLoadTimer.Visible = False
730             .chkLoadTimer_lbl1.Visible = False
740           End Select
750           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
760           DoEvents

770           DoCmd.Hourglass False

780         Else
790           Cancel = -1
800         End If
810       Else
820         Cancel = -1
830       End If
840     End With

850     If Cancel = -1 Then
860       Beep
870       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
880     End If

EXITP:
890     Set usr = Nothing
900     Set grp = Nothing
910     Exit Sub

ERRH:
920     DoCmd.Hourglass False
930     Select Case ERR.Number
        Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer, Cancel As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Load Timer:       D {chkLoadTimer}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Force Date:       T {blnForce: cmdTransaction}

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
1200            chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
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
          Case vbKeyT
1300        intRetVal = 0
1310        blnForce = True
1320        cmdTransaction_Click  ' ** Procedure: Below.
1330      End Select
1340    End If

        ' ** Ctrl-Shift keys.
1350    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1360      Select Case intRetVal
          Case vbKeyF
1370        With Me
1380          intRetVal = 0
1390          .FocusHolder.SetFocus
1400        End With
1410      Case vbKeyT
1420        With Me
1430          intRetVal = 0
1440          If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
1450            FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
1460          ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
1470            Select Case .chkLoadTimer.Visible
                Case True
1480              .FocusHolder.SetFocus
1490              DoEvents
1500              .chkLoadTimer.Visible = False
1510              .chkLoadTimer_lbl1.Visible = False
1520              gblnTimer = False
1530            Case False
1540              .chkLoadTimer.Visible = True
1550              .chkLoadTimer_lbl1.Visible = True
1560              chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
1570            End Select
1580          End If
1590          .chkLoadTimer_vis = .chkLoadTimer.Visible
1600        End With
1610      Case vbKeyX
1620        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1630          intRetVal = 0
1640          blnJustClose = True
1650          SetOption_Dev  ' ** Module Function: modUtilities.
1660          gblnSignal = True  ' ** Borrowing this variable.
1670          DoCmd.Close acForm, THIS_NAME
1680          DoCmd.Hourglass False
1690        End If
1700      End Select
1710    End If

EXITP:
1720    KeyCode = intRetVal
1730    Exit Sub

ERRH:
1740    intRetVal = 0
1750    Select Case ERR.Number
        Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Private Sub Form_Timer()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

1810    With Me
1820      .TimerInterval = 0&
1830      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1840        lngGTR_Stat = lngGTR_Stat + 1&
1850        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1860        lngTmp01 = lngTmp01 + 1&
1870        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1880        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1890        If lngTmp01 > 1& Then
1900          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1910        End If
1920        If lngTmp01 = 24& Then
1930          lngTmp01 = 0&
1940          lngTmp02 = lngTmp02 + 1&
1950          If lngTmp02 < 5& Then
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
2130        DoEvents
2140      End If
2150      If blnGoingToReport = False And blnUpdateTransCnt = False Then
2160        Select Case gstrReturningForm
            Case "frmPostingDate", "frmJournal", "frmJournal_Columns"
2170          .cmdTransaction.SetFocus
2180          If IsLoaded("frmMenu_Background", acForm) = True Then  ' ** Module Function: modFileUtilities.
2190            DoCmd.Close acForm, "frmMenu_Background"
2200          End If
2210          Pref_RemNot "frmJournal", "Account", "LastAcctNo"  ' ** Module Procedure: modPreferenceFuncs.
2220          Pref_RemNot "frmJournal_Columns", "Account", "LastAcctNo"  ' ** Module Procedure: modPreferenceFuncs.
2230        Case "frmCheckPOSPay"
2240          .cmdPOSPay.SetFocus
2250        Case "frmCheckMemo"
2260          .cmdCheckMemos.SetFocus
2270        Case "frmRpt_Checks"
2280          .cmdChecksPrint.SetFocus
2290        Case "frmCheckReconcile"
2300          .cmdCheckReconcile.SetFocus
2310        Case "frmSweeper"
2320          .cmdSweep.SetFocus
2330        Case "frmAccruedIncome"
2340          .cmdAccrue.SetFocus
2350        Case "frmFeeCalculations"
2360          .cmdFeeCalculations.SetFocus
2370        Case Else
2380          .cmdTransaction.SetFocus
2390        End Select
2400        gstrReturningForm = vbNullString
2410        .Repaint
2420        DoEvents
2430        UpdateTransCnt  ' ** Function: Below.
2440        If gblnGoToReport = False Then
2450          blnGTR_Emblem = False
2460          GTREmblem_Off  ' ** Procedure: Below.
2470        End If
2480        DoCmd.Hourglass False
2490      End If
2500      If gblnGoToReport = True Then
2510        DoCmd.Hourglass True  ' ** Make sure it's still running.
2520        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
2530        Else
2540          If blnGTR_Emblem = False And garr_varGoToReport(GTR_INV1) = True Then
2550            blnGTR_Emblem = True
2560            lngTmp01 = 0&: lngTmp02 = 0&
2570          End If
2580          DoEvents
2590          Select Case blnGoingToReport
              Case True
2600            Select Case garr_varGoToReport(GTR_INV1)
                Case True
2610              Select Case garr_varGoToReport(GTR_CTL1)
                  Case "cmdTransaction"
2620                .TimerInterval = 50&
2630                cmdTransaction_Click  ' ** Procedure: Below.
2640              Case "cmdChecksPrint"
2650                .TimerInterval = 50&
2660                cmdChecksPrint_Click  ' ** Procedure: Below.
2670              Case "cmdCheckReconcile"
2680                .TimerInterval = 50&
2690                cmdCheckReconcile_Click  ' ** Procedure: Below.
2700              Case "cmdFeeCalculations"
2710                .TimerInterval = 50&
2720                cmdFeeCalculations_Click  ' ** Procedure: Below.
2730              Case Else
                    ' ** Shouldn't happen.
2740                blnGoingToReport = False
2750                DoCmd.Hourglass False
2760                gblnGoToReport = False
2770                blnGTR_Emblem = False
2780                GTREmblem_Off  ' ** Procedure: Below.
2790                Beep
2800              End Select
2810            Case False
                  ' ** Just sit there, and we're done.
2820              Select Case garr_varGoToReport(GTR_CTL1)
                  Case "cmdPostingJournalPrint", "cmdChecksPreview"
                    ' ** Just sit there, and we're done.
2830                blnGoingToReport = False
2840                gblnGoToReport = False
2850                .GoToReport_arw_menu_l_img.Visible = False
2860                .GoToReport_arw_menu_r_img.Visible = False
2870                DoCmd.Hourglass False  ' ** Should already be off.
2880                DoEvents
2890              Case Else
2900                blnGoingToReport = False
2910                gblnGoToReport = False
2920                blnGTR_Emblem = False
2930                GTREmblem_Off  ' ** Procedure: Below.
2940                Beep
2950                DoCmd.Hourglass False
2960              End Select
2970            End Select
2980          Case False
2990            If garr_varGoToReport(GTR_FRM1) = THIS_NAME Then
                  ' ** This is first menu in path.
3000              blnGoingToReport = True: blnGoingToReport2 = False
3010              If IsNull(garr_varGoToReport(GTR_FRM2)) = True And garr_varGoToReport(GTR_INV1) = False And _
                      garr_varGoToReport(GTR_RNAM) <> "rptPostingJournal_Column" Then
                    ' ** No more forms.
3020                .Controls(garr_varGoToReport(GTR_CTL1)).SetFocus
3030                lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
3040                lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
3050                Select Case lngBtn_Left
                    Case .cmdTransaction.Left
3060                  .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
3070                  .GoToReport_arw_menu_l_img.Visible = True
3080                Case .cmdChecksPreview.Left
3090                  .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
3100                  .GoToReport_arw_menu_r_img.Visible = True
3110                End Select
3120                Beep
3130                DoCmd.Hourglass False
3140                DoEvents
3150                .TimerInterval = (GTR_WAIT * 2&)
3160              Else
                    ' ** Moving on.
3170                .Controls(garr_varGoToReport(GTR_CTL1)).SetFocus
3180                lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
3190                lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
3200                Select Case lngBtn_Left
                    Case .cmdTransaction.Left
3210                  .GoToReport_arw_menu_l_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
3220                  .GoToReport_arw_menu_l_img.Visible = True
3230                Case .cmdChecksPreview.Left
3240                  .GoToReport_arw_menu_r_img.Top = lngBtn_Top + (3& * lngTpp)
3250                  .GoToReport_arw_menu_r_img.Visible = True
3260                End Select
3270                Select Case garr_varGoToReport(GTR_INV1)
                    Case True
3280                  DoEvents
3290                  .TimerInterval = 50&
3300                Case False
                      ' ** Just sit there, and we're done.
3310                  blnGoingToReport = False
3320                  DoCmd.Hourglass False
3330                  gblnGoToReport = False
3340                  blnGTR_Emblem = False
3350                  GTREmblem_Off  ' ** Procedure: Below.
3360                  Beep
3370                End Select
3380              End If
3390            Else
                  ' ** Shouldn't be here, then!
3400              blnGoingToReport = False
3410              DoCmd.Hourglass False
3420              gblnGoToReport = False
3430              blnGTR_Emblem = False
3440              GTREmblem_Off  ' ** Procedure: Below.
3450              Beep
3460            End If
3470          End Select
3480        End If
3490      ElseIf blnGoingToReport2 = True Then
3500        blnGoingToReport2 = False
3510        .GoToReport_arw_menu_l_img.Visible = False
3520        .GoToReport_arw_menu_r_img.Visible = False
3530      End If
3540      If blnUpdateTransCnt = True Then
3550        blnUpdateTransCnt = False
3560        UpdateTransCnt  ' ** Function: Below.
3570      End If
3580      If gblnTimer = True Then
3590        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
3600          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
3610          gblnTimer = False
3620          DoCmd.Hourglass False
3630        End If
3640      End If
3650    End With

EXITP:
3660    Exit Sub

ERRH:
3670    DoCmd.Hourglass False
3680    blnGoingToReport = False
3690    gblnGoToReport = False
3700    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Ignore.
3710    Case Else
3720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3730    End Select
3740    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3810    With Me

3820      If gblnSignal = False Then  ' ** Borrowing this variable.
3830        DoCmd.Hourglass True
3840        DoEvents
3850      End If

          ' ** Make sure all reports are closed.
3860      Do While Reports.Count > 0
3870        DoCmd.Close acReport, Reports(0).Name
3880      Loop

3890      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3900      blnGoingToReport = False

3910      If blnChecksCalled = True And blnJustClose = False Then
3920        Set dbs = CurrentDb
            ' ** Empty tblJournal_Check_Staging1.
3930        Set qdf = dbs.QueryDefs("qryPrintChecks_16a")
3940        qdf.Execute
3950        Set qdf = Nothing
            ' ** Empty tblJournal_Check_Staging2.
3960        Set qdf = dbs.QueryDefs("qryPrintChecks_16b")
3970        qdf.Execute
3980        Set qdf = Nothing
3990        dbs.Close
4000        Set dbs = Nothing
4010      End If

4020      gblnAdmin = False
4030      gstrFormQuerySpec = vbNullString

4040      If blnJustClose = False Then

4050        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
4060        DoEvents

4070        If IsLoaded("frmMenu_Background", acForm) = True Then  ' ** Module Function: modFileUtilities.
4080          DoCmd.Close acForm, "frmMenu_Background"
4090        End If
4100        gblnSetFocus = True
4110        gstrReturningForm = THIS_NAME
4120        Select Case strCallingForm
            Case "frmMenu_Main"
4130          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Title"
4140        End Select

4150      End If  ' ** blnJustClose.

4160    End With

EXITP:
4170    Set qdf = Nothing
4180    Set dbs = Nothing
4190    Exit Sub

ERRH:
4200    DoCmd.Hourglass False
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4310    With Me
4320      If .cmdClose.Caption <> "&Close" Then
4330        .cmdClose.Caption = "&Close"
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
        Case 2501  ' ** The Close action was canceled.
          ' ** Ignore.
5040    Case Else
5050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5060    End Select
5070    Resume EXITP

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
5190          .cmdTransaction.SetFocus
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
5290          Else
5300            .cmdFeeCalculations.SetFocus
5310          End If
5320        End With
5330      End Select
5340    End If

EXITP:
5350    KeyCode = intRetVal
5360    Exit Sub

ERRH:
5370    intRetVal = 0
5380    Select Case ERR.Number
        Case Else
5390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5400    End Select
5410    Resume EXITP

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

Private Sub cmdTransaction_Click()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransaction_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnHasJrnl As Boolean, blnHasJrnlCol As Boolean
        Dim blnHasJrnl2 As Boolean, blnHasJrnlCol2 As Boolean
        Dim strInputOption As String, lngSubForm As Long
        Dim lngLocs As Long, arr_varLoc As Variant
        Dim lngCnt3 As Long, lngCnt4 As Long, lngCnt5 As Long
        Dim varTmp00 As Variant, strTmp01 As String, blnTmp02 As Boolean, blnTmp03 As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varLoc().
        'Const J_JID    As Integer = 0
        Const J_JCID   As Integer = 1
        'Const J_LID    As Integer = 2
        Const J_LIDNEW As Integer = 3
        'Const J_NAM    As Integer = 4
        Const J_NAMNEW As Integer = 5
        'Const J_NMD    As Integer = 6
        Const J_NMDNEW As Integer = 7

5610    With Me
5620      If blnFormOpening = False Then

5630        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
5640        DoEvents

5650        If .chkLoadTimer.Visible = True Then
5660          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
5670          If gblnTimer = True Then
                'gblnSetFocus = True
5680            strTmp01 = "frmJournal^frmJournal_Columns"
5690            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strTmp01 & "~True"
5700          End If
5710        End If
5720        DoEvents

5730        DoCmd.Hourglass True
5740        DoEvents

5750        .TimerInterval = 0&

5760        blnRetVal = True

5770        blnHasJrnl = True: blnHasJrnlCol = True: blnHasJrnl2 = False: blnHasJrnlCol2 = False
5780        strInputOption = vbNullString
5790        lngCnt3 = 0&: lngCnt4 = 0&: lngCnt5 = 0&

5800        gstrDividendAccountNumber = vbNullString
5810        gstrInterestAccountNumber = vbNullString
5820        gstrPurchaseAccountNumber = vbNullString
5830        gstrSaleAccountNumber = vbNullString
5840        gstrMiscAccountNumber = vbNullString
5850        gstrAccountNo = vbNullString
5860        gstrAccountName = vbNullString

5870        UpdateImported  ' ** Procedure: Below.
5880        If gblnGoToReport = True Then
5890          Journal_OptLoad  ' ** Procedure: Below.
5900        End If

5910        DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!

5920        Set dbs = CurrentDb
5930        With dbs

5940          If gblnGoToReport = False Then
                ' *****************************************
                ' ** These check for various Null fields.
                ' *****************************************
                ' ** PostingDate table, Username = Null.
5950            varTmp00 = DCount("*", "qryPostingDate_03")
5960            If IsNull(varTmp00) = False Then lngCnt3 = varTmp00
                ' ** Journal, grouped by journal_USER, for journal_USER = Null, with cnt.
5970            varTmp00 = DLookup("[cnt]", "qryPostingDate_10")
5980            If IsNull(varTmp00) = False Then lngCnt4 = varTmp00
                ' ** Journal, grouped by transdate, for transdate = Null, with cnt.
5990            varTmp00 = DLookup("[cnt]", "qryPostingDate_11")
6000            If IsNull(varTmp00) = False Then lngCnt5 = varTmp00
                ' *****************************************
6010          End If

6020          DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!

              ' *****************************************
              ' ** Only if Null's present.
              ' *****************************************
              ' ** Is it possible this could cause a problem on multi-user systems?
6030          If lngCnt3 > 0& Or lngCnt4 > 0& Or lngCnt5 > 0& Then
6040            If lngCnt3 > 0& Then
                  ' ** Null Username in PostingDate.
6050              CheckDates  ' ** Module Function: modDateCheckFuncs.
6060            End If
6070            If lngCnt4 > 0& Then
6080              Set dbs = CurrentDb
6090              With dbs
                    ' ** Delete Journal, for journal_USER = Null.
6100                Set qdf = .QueryDefs("qryPostingDate_12")
6110                qdf.Execute
6120                Set qdf = Nothing
6130                .Close
6140              End With
6150              Set dbs = Nothing
6160            End If
6170            If lngCnt5 > 0& Then
                  ' ** Null transdate in Journal.
6180              Set dbs = CurrentDb
6190              With dbs
                    ' ** Delete Journal, for transdate = Null.
6200                Set qdf = .QueryDefs("qryPostingDate_13")
6210                qdf.Execute
6220                Set qdf = Nothing
6230                .Close
6240              End With
6250              Set dbs = Nothing
6260            End If
6270          End If
              ' *****************************************

6280          DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!

              ' ** Journal, by Currentuser().
6290          Set qdf = .QueryDefs("qryJournal_06")
6300          Set rst = qdf.OpenRecordset
6310          With rst
6320            If rst.BOF = True And rst.EOF = True Then
6330              blnHasJrnl = False
6340            End If
6350            .Close
6360          End With
6370          Set rst = Nothing
6380          Set qdf = Nothing

6390          DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!

              ' ** tblJournal_Column, by CurrentUser().
6400          Set qdf = .QueryDefs("qryJournal_Columns_05")
6410          Set rst = qdf.OpenRecordset
6420          With rst
6430            If rst.BOF = True And rst.EOF = True Then
6440              blnHasJrnlCol = False
6450            End If
6460            .Close
6470          End With
6480          Set rst = Nothing
6490          Set qdf = Nothing

6500          DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!

6510          blnHasJrnl2 = blnHasJrnl: blnHasJrnlCol2 = blnHasJrnlCol

6520          If gblnGoToReport = False Then
6530            If blnHasJrnlCol = True Then
                  ' ** tblJournal_Column, linked to Location, just discrepancies.
6540              Set qdf = .QueryDefs("qryJournal_Columns_05a")
6550              Set rst = qdf.OpenRecordset
6560              With rst
6570                If .BOF = True And .EOF = True Then
                      ' ** No discrepancies.
6580                  lngLocs = 0&
6590                Else
6600                  .MoveLast
6610                  lngLocs = .RecordCount
6620                  .MoveFirst
6630                  arr_varLoc = .GetRows(lngLocs)
                      ' *********************************************************
                      ' ** Array: arr_varLoc()
                      ' **
                      ' **   Field  Element  Name                    Constant
                      ' **   =====  =======  ======================  ==========
                      ' **     1       0     Journal_ID              J_JID
                      ' **     2       1     JrnlCol_ID              J_JCID
                      ' **     3       2     Location_ID             J_LID
                      ' **     4       3     Location_ID_new         J_LIDNEW
                      ' **     5       4     Loc_Name                J_NAM
                      ' **     6       5     Loc_Name_new            J_NAMNEW
                      ' **     7       6     Loc_Name_display        J_NMD
                      ' **     8       7     Loc_Name_display_new    J_NMDNEW
                      ' **
                      ' *********************************************************
6640                End If
6650                .Close
6660              End With
6670              Set rst = Nothing
6680              Set qdf = Nothing
6690              DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!
6700              If lngLocs > 0& Then
6710                For lngX = 0& To (lngLocs - 1&)
6720                  If IsNull(arr_varLoc(J_LIDNEW, lngX)) = False Then
                        ' ** Update tblJournal_Column, by specified [jcolid], [locid].
6730                    Set qdf = .QueryDefs("qryJournal_Columns_05b")
6740                    With qdf.Parameters
6750                      ![jcolid] = arr_varLoc(J_JCID, lngX)
6760                      ![locid] = arr_varLoc(J_LIDNEW, lngX)
6770                    End With
6780                    qdf.Execute
6790                    Set qdf = Nothing
6800                  End If
6810                  DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!
6820                  If IsNull(arr_varLoc(J_NAMNEW, lngX)) = False Then
                        ' ** Update tblJournal_Column, by specified [jcolid], [locnam].
6830                    Set qdf = .QueryDefs("qryJournal_Columns_05c")
6840                    With qdf.Parameters
6850                      ![jcolid] = arr_varLoc(J_JCID, lngX)
6860                      ![locnam] = arr_varLoc(J_NAMNEW, lngX)
6870                    End With
6880                    qdf.Execute
6890                    Set qdf = Nothing
6900                  End If
6910                  DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!
6920                  If IsNull(arr_varLoc(J_NMDNEW, lngX)) = False Then
                        ' ** Update tblJournal_Column, by specified [jcolid], [locnamdis].
6930                    Set qdf = .QueryDefs("qryJournal_Columns_05d")
6940                    With qdf.Parameters
6950                      ![jcolid] = arr_varLoc(J_JCID, lngX)
6960                      ![locnamdis] = arr_varLoc(J_NMDNEW, lngX)
6970                    End With
6980                    qdf.Execute
6990                    Set qdf = Nothing
7000                  End If
7010                Next
7020              End If  ' ** lngLocs.
7030            End If  ' ** blnHasJrnlCol.
7040          End If  ' ** gblnGoToReport.

7050          DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!

              ' ** Because of JrnlPrefCheck(), this user DEFINITELY has a preference!

7060          If gblnGoToReport = False Then

                ' ** tblPreference_User, linked to qryPostingDate_04 (tblPreference_Control,
                ' ** just frmPostingDate.opgInput, with ctl_id_child, ctlspec_optionvalue),
                ' ** by prefuser_integer, for CurrentUser().
7070            Set qdf = .QueryDefs("qryPostingDate_05")  '##dbs_id
7080            Set rst = qdf.OpenRecordset
7090            With rst
7100              If .BOF = True And .EOF = True Then
                    ' ** Preference not set.
                    ' ** Any time no preference is set, always open frmPostingDate
                    ' ** regardless of whether there are Journal entries or not.
7110                blnHasJrnl = False
7120                blnHasJrnlCol = False
7130              Else
7140                .MoveFirst
7150                strInputOption = ![ctl_name_child]
7160              End If
7170              .Close
7180            End With
7190            Set rst = Nothing
7200            Set qdf = Nothing

7210            If (blnHasJrnl = False And blnHasJrnl2 = True) Or (blnHasJrnlCol = False And blnHasJrnlCol2 = True) Then
                  ' ** What got lost?
                  'Debug.Print "'PREF LOST OR QRY NOT WORKING!"
7220              JrnlPrefCheck  ' ** Module Function: modXAdminFuncs.
7230              DoEvents
7240              varTmp00 = DLookup("[ctl_name_child]", "qryPostingDate_05")
7250              Select Case IsNull(varTmp00)
                  Case True
                    ' ** I think it's really there!
7260                strInputOption = "opgInput_optClassic"
7270              Case False
7280                strInputOption = varTmp00
7290              End Select
7300              blnHasJrnl = blnHasJrnl2
7310              blnHasJrnlCol = blnHasJrnlCol2
7320            End If

7330          End If  ' ** gblnGoToReport.

7340          .Close
7350        End With  ' ** dbs.

7360        DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!

7370        If gblnGoToReport = True Then
7380          Select Case IsNull(garr_varGoToReport(GTR_FRM2))
              Case True
7390            If garr_varGoToReport(GTR_RNAM) = "rptPostingJournal_Column" Then
7400              strInputOption = "opgInput_optColumnar"
7410            End If
7420          Case False
7430            Select Case garr_varGoToReport(GTR_FRM2)
                Case "frmJournal"
7440              strInputOption = "opgInput_optClassic"
7450            Case "frmJournal_Columns"
7460              strInputOption = "opgInput_optColumnar"
7470            End Select
7480          End Select
7490        End If  ' ** gblnGoToReport.

7500        Select Case blnForce
            Case True
7510          blnHasJrnl = False: blnHasJrnlCol = False
7520          strDocName = "frmPostingDate"
7530        Case False
7540          Select Case strInputOption
              Case "opgInput_optClassic"
7550            Select Case blnHasJrnl
                Case True
7560              strDocName = "frmJournal"
7570            Case False
7580              strDocName = "frmPostingDate"
7590            End Select
7600          Case "opgInput_optColumnar"
7610            If blnHasJrnl = True Or blnHasJrnlCol = True Then
7620              strDocName = "frmJournal_Columns"
7630            Else
7640              strDocName = "frmPostingDate"
7650            End If
7660          Case Else
7670            Select Case blnHasJrnl
                Case True
7680              strDocName = "frmJournal"
7690            Case False
7700              strDocName = "frmPostingDate"
7710            End Select
7720          End Select
7730        End Select  ' ** blnForce.

7740        If strDocName = "frmJournal_Columns" Then
7750          blnRetVal = JrnlCol_Set  ' ** Module Function: modUtilities.
7760        End If

7770        If blnRetVal = True Then

7780          DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!

              ' ** Determine whether this is a single-user environment or not.
7790          gblnSingleUser = IsSingleUser  ' ** Module Function: modSecurityFunctions.

7800          Select Case gblnGoToReport
              Case True
7810            lngSubForm = 0&
7820            Select Case garr_varGoToReport(GTR_RNAM)
                Case "rptPostingJournal_Column"
                  ' ** frmJournal not an option.
7830            Case "rptListOfTaxLots"
7840              lngSubForm = JRN_SAL
7850            Case "rptMap_Dividend"
7860              lngSubForm = JRN_DIV
7870            Case "rptMap_Reinvest_Div"
7880              lngSubForm = JRN_DIV
7890            Case "rptMap_Interest"
7900              lngSubForm = JRN_INT
7910            Case "rptMap_Reinvest_Int"
7920              lngSubForm = JRN_INT
7930            Case "rptMap_Received"
7940              lngSubForm = JRN_MIS
7950            Case "rptMap_Reinvest_Rec"
7960              lngSubForm = JRN_MIS
7970            Case "rptMap_Split"
7980              lngSubForm = JRN_PUR
7990            Case "rptMap_Misc_LTCL"
8000              lngSubForm = JRN_MIS
8010            Case "rptMap_Misc_STCG"
8020              lngSubForm = JRN_MIS
8030            Case "rptMap_Misc_STCL"
8040              lngSubForm = JRN_MIS
8050            Case Else
                  ' ** Shouldn't be any others.
8060              strDocName = vbNullString
8070            End Select
8080            Select Case strDocName
                Case "frmPostingDate"
8090              Set dbs = CurrentDb
8100              With dbs
8110                Set grstPostingDate = .OpenRecordset("PostingDate", dbOpenDynaset, dbConsistent)
8120                With grstPostingDate
8130                  .FindFirst "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
8140                  If .NoMatch = False Then
8150                    glngPostingDateID = ![PostingDate_ID]
8160                    .Edit
8170                    ![Posting_Date] = Date    ' ** Default to today's date.
8180                    .Update
8190                    blnTmp03 = True
8200                  Else
                        ' ** There should always be one by this time.
8210                    blnTmp02 = True
8220                    .AddNew
8230                    ![Posting_Date] = Date
8240                    ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
8250                    .Update
8260                    .Bookmark = .LastModified
8270                    glngPostingDateID = ![PostingDate_ID]
8280                  End If
8290                  .Close
8300                End With  ' ** grstPostingDate.
8310                DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!
8320                Set grstPostingDate = Nothing
                    ' ** Update tblCalendar_Staging, for unique_id, by specified [unqid].
8330                Set qdf = .QueryDefs("qryCalendar_02")
8340                With qdf.Parameters
8350                  ![unqid] = glngPostingDateID
8360                End With
8370                qdf.Execute
8380                Set qdf = Nothing
8390                .Close
8400              End With  ' ** dbs.
8410              Set dbs = Nothing
8420              If garr_varGoToReport(GTR_RNAM) = "rptPostingJournal_Column" Then
8430                strDocName = "frmJournal_Columns"
8440                lngSubForm = JRN_DIV  ' ** Not applicable.
8450              Else
8460                Select Case strInputOption
                    Case "opgInput_optClassic"
8470                  strDocName = "frmJournal"
8480                Case "opgInput_optColumnar"
8490                  strDocName = "frmJournal_Columns"
8500                  lngSubForm = JRN_DIV  ' ** Not applicable.
8510                Case vbNullString
8520                  strDocName = "frmJournal"
8530                End Select
8540              End If
8550            Case "frmJournal"
                  ' ** Just proceed normally.
8560            Case "frmJournal_Columns"
                  ' ** Just proceed normally.
8570              lngSubForm = JRN_DIV  ' ** Not applicable.
8580            End Select
8590            DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!
8600            If strDocName <> vbNullString Then
8610              gblnSetFocus = True
8620              blnJustClose = True
8630              glngJournalForm = lngSubForm
8640              DoCmd.Close acForm, THIS_NAME  ' ** Close.
8650              gstrReturningForm = vbNullString
8660              DoCmd.OpenForm strDocName, , , , , , THIS_NAME
8670              DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!
8680              Forms(strDocName).TimerInterval = 50&
8690              blnGoingToReport = False
8700              blnGTR_Emblem = False
8710            End If
8720          Case False

8730            blnTmp02 = False: blnTmp03 = False
8740            Set dbs = CurrentDb
8750            With dbs
8760              Set grstPostingDate = .OpenRecordset("PostingDate", dbOpenDynaset, dbConsistent)
8770              With grstPostingDate
8780                .FindFirst "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
8790                If .NoMatch = False Then
8800                  glngPostingDateID = ![PostingDate_ID]
8810                  If IsNull(![Posting_Date]) = True Then
8820                    varTmp00 = DCount("*", "journal", "[journal_USER] = '" & CurrentUser & "'")  ' ** Internal Access Function: Trust Accountant login.
8830                    Select Case IsNull(varTmp00)
                        Case True
8840                      .Edit
8850                      ![Posting_Date] = Date
8860                      .Update
8870                      blnTmp03 = True
8880                    Case False
8890                      If varTmp00 > 0 Then
8900                        varTmp00 = DMax("[transdate]", "journal", "[journal_USER] = '" & CurrentUser & "'")  ' ** Internal Access Function: Trust Accountant login.
8910                        Select Case IsNull(varTmp00)
                            Case True
8920                          .Edit
8930                          ![Posting_Date] = Date
8940                          .Update
8950                          blnTmp03 = True
8960                        Case False
8970                          Select Case IsDate(varTmp00)
                              Case True
8980                            .Edit
8990                            ![Posting_Date] = CDate(varTmp00)
9000                            .Update
9010                            blnTmp03 = True
9020                          Case False
9030                            .Edit
9040                            ![Posting_Date] = Date
9050                            .Update
9060                            blnTmp03 = True
9070                          End Select
9080                        End Select
9090                      Else
9100                        .Edit
9110                        ![Posting_Date] = Date
9120                        .Update
9130                        blnTmp03 = True
9140                      End If
9150                    End Select
9160                  End If
9170                Else
                      ' ** There should always be one by this time.
9180                  blnTmp02 = True
9190                  .AddNew
9200                  ![Posting_Date] = Date
9210                  ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
9220                  .Update
9230                  .Bookmark = .LastModified
9240                  glngPostingDateID = ![PostingDate_ID]
9250                End If
9260                .Close
9270              End With  ' ** grstPostingDate.
9280              DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!
9290              Set grstPostingDate = Nothing
                  ' ** Update tblCalendar_Staging, for unique_id, by specified [unqid].
9300              Set qdf = .QueryDefs("qryCalendar_02")
9310              With qdf.Parameters
9320                ![unqid] = glngPostingDateID
9330              End With
9340              qdf.Execute
9350              Set qdf = Nothing
9360              DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!
9370              .Close
9380            End With

9390            If blnHasJrnl = True Or blnHasJrnlCol = True Then
9400              If strDocName = "frmPostingDate" Then
9410                JrnlDateCheck  ' ** Module Function: modXAdminFuncs.
9420                Select Case strInputOption
                    Case "opgInput_optClassic"
9430                  strDocName = "frmJournal"
9440                Case "opgInput_optColumnar"
9450                  strDocName = "frmJournal_Columns"
9460                End Select
9470              End If
9480            Else
                  'Debug.Print "'blnHasJrnl: " & blnHasJrnl & "  blnHasJrnlCol: " & blnHasJrnlCol
9490            End If

9500            blnJustClose = True
9510            glngJournalForm = JRN_DIV
9520            gblnSetFocus = True
9530            DoCmd.Close acForm, THIS_NAME  ' ** Close.
9540            gstrReturningForm = vbNullString
9550            DoCmd.OpenForm strDocName, , , , , , THIS_NAME
9560            DoEvents  ' ** Trying to keep it from slipping into 'Not Responding'!

9570          End Select  ' ** blnGoingToReport.

9580        End If  ' ** blnRetVal.

9590        If blnRetVal = False And gblnTimer = True Then
9600          gblnTimer = False
9610          DoCmd.Close acForm, "frmLoadTimer"
9620          .chkLoadTimer = False
9630          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
9640        End If

9650      End If  ' ** blnFormOpening.

9660      DoCmd.Hourglass False

9670    End With

        'PLAIN:
        'HERE! 1  cmdTransaction_Click()
        'CTRL MOUSE:
        'HERE! 3  cmdTransaction_MouseDown()
        'HERE! 1  cmdTransaction_Click()
        'CTRL ENTER:
        'HERE! 2  cmdTransaction_KeyDown()
        'HERE! 1  cmdTransaction_Click()

EXITP:
9680    Set grstPostingDate = Nothing
9690    Set rst = Nothing
9700    Set qdf = Nothing
9710    Set dbs = Nothing
9720    Exit Sub

ERRH:
9730    blnFormOpening = False
9740    DoCmd.Hourglass False
9750    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Ignore.
9760    Case Else
9770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9780    End Select
9790    Resume EXITP

End Sub

Private Sub cmdTransaction_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransaction_KeyDown"

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
9890          .cmdPostingJournalPreview.SetFocus
9900        End With
9910      Case vbKeyLeft
9920        With Me
9930          intRetVal = 0
9940          Select Case strMenuArrow
              Case "cmdChecksPrint_KeyDown"
9950            .cmdChecksPrint.SetFocus
9960          Case Else
9970            .cmdChecksPreview.SetFocus
9980          End Select
9990        End With
10000     Case vbKeyRight
10010       With Me
10020         intRetVal = 0
10030         Select Case strMenuArrow
              Case "cmdChecksPrint_KeyDown"
10040           .cmdChecksPrint.SetFocus
10050         Case Else
10060           .cmdChecksPreview.SetFocus
10070         End Select
10080       End With
10090     Case vbKeyUp
10100       With Me
10110         intRetVal = 0
10120         .cmdClearPostingJournal.SetFocus
10130       End With
10140     Case vbKeyDown
10150       With Me
10160         intRetVal = 0
10170         .cmdPostingJournalPreview.SetFocus
10180       End With
10190     End Select
10200   End If

        ' ** Shift keys.
10210   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10220     Select Case intRetVal
          Case vbKeyTab
10230       With Me
10240         intRetVal = 0
10250         .cmdClose.SetFocus
10260       End With
10270     End Select
10280   End If

        ' ** Ctrl keys.
10290   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10300     Select Case intRetVal
          Case vbKeyReturn
10310       intRetVal = 0
10320       blnForce = True
10330       cmdTransaction_Click  ' ** Procedure: Above.
10340     End Select
10350   End If

EXITP:
10360   KeyCode = intRetVal
10370   Exit Sub

ERRH:
10380   intRetVal = 0
10390   Select Case ERR.Number
        Case Else
10400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10410   End Select
10420   Resume EXITP

End Sub

Private Sub cmdTransaction_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTransaction_MouseDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer

        ' ** Use bit masks to determine which key was pressed.
10510   intShiftDown = (Shift And acShiftMask) > 0
10520   intAltDown = (Shift And acAltMask) > 0
10530   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Ctrl keys.
10540   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10550     blnForce = True
10560   End If

EXITP:
10570   Exit Sub

ERRH:
10580   Select Case ERR.Number
        Case Else
10590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10600   End Select
10610   Resume EXITP

End Sub

Private Sub cmdPostingJournalPreview_Click()
' ** Check for gblnAdmin vs. others.
' ** See also TransCount() in modPostUtilities.

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPostingJournalPreview_Click"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim blnContinue As Boolean

10710   With Me
10720     If blnFormOpening = False Then

10730       DoCmd.Hourglass True
10740       DoEvents

10750       .TimerInterval = 0&

10760       blnContinue = True
10770       blnFormOpening = True

10780       gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10790       UpdateImported  ' ** Procedure: Below.

10800       Set dbs = CurrentDb
10810       With dbs
10820         Select Case gblnAdmin
              Case True
                ' **Journal, with add'l fields, all journal_USER's.
10830           Set qdf1 = .QueryDefs("qryPost_Journal_02a")  '#jnox
10840         Case False
                ' ** qryPost_Journal_02b (Journal, with add'l fields, by specified journal_USER), by specified [usr].
10850           Set qdf1 = .QueryDefs("qryPost_Journal_02c")  '#jnox
10860           With qdf1.Parameters
10870             ![usr] = gstrJournalUser
10880           End With
10890         End Select
10900         Set rst = qdf1.OpenRecordset
10910         With rst
10920           If .BOF = True And .EOF = True Then
10930             blnContinue = False
10940             DoCmd.Hourglass False
10950             Select Case gblnAdmin
                  Case True
10960               MsgBox "There are no entries in the posting journal.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
10970             Case False
10980               MsgBox "There are no entries in the posting journal" & vbCrLf & _
                      "for User: " & gstrJournalUser & ".", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
10990             End Select
11000           End If
11010           .Close
11020         End With  ' ** rst.
11030         Set rst = Nothing
11040         Set qdf1 = Nothing
11050         If blnContinue = True Then
11060           Select Case gblnAdmin
                Case True
                  ' ** qryPost_Journal_02a (Journal, with add'l fields; all journal_USER's), grouped by curr_id.
11070             Set qdf1 = .QueryDefs("qryPost_Journal_02e")
11080             Set qdf2 = .QueryDefs("qryPost_Journal_02g")
11090           Case False
                  ' ** qryPost_Journal_02c (qryPost_Journal_02b (Journal, with add'l fields;
                  ' ** by specified journal_USER), by specified [usr]), grouped by curr_id.
11100             Set qdf1 = .QueryDefs("qryPost_Journal_02f")
11110             Set qdf2 = .QueryDefs("qryPost_Journal_02h")
11120             With qdf1.Parameters
11130               ![usr] = gstrJournalUser
11140             End With
11150             With qdf2.Parameters
11160               ![usr] = gstrJournalUser
11170             End With
11180           End Select
11190           Set rst = qdf1.OpenRecordset
11200           With rst
11210             .MoveLast
11220             lngRecs = .RecordCount
11230             .MoveFirst
11240             If lngRecs = 1 Then
11250               If ![curr_code] = "USD" Then
11260                 strDocName = "rptPostingJournal_Classic"
11270               Else
11280                 strDocName = "rptPostingJournal_Classic_ForEx"
11290               End If
11300             Else
11310               strDocName = "rptPostingJournal_Classic_ForEx"
11320             End If
11330             .Close
11340           End With  ' ** rst.
11350           Set rst = Nothing
11360           Set qdf1 = Nothing
11370           DoEvents
11380           Set rst = qdf2.OpenRecordset
11390           rst.MoveFirst
11400           lngRecs = rst![TransCnt]
11410           rst.Close
11420           Set rst = Nothing
11430           Set qdf2 = Nothing
11440           DoEvents
11450         End If  ' ** blnContinue.
11460         .Close
11470       End With  ' ** dbs.

11480       If blnContinue = True Then

              ' ** Determine whether 'Posted By' should be included in transaction description.
11490         gblnSingleUser = IsSingleUser  ' ** Module Function: modSecurityFunctions.

11500         gstrReportCallingForm = THIS_NAME  ' ** When not vbNullString, sets this form .Visible = False.
11510         DoCmd.Hourglass False
11520         Select Case gblnAdmin
              Case True
11530           DoCmd.OpenReport strDocName, acViewPreview, , , , CStr(lngRecs)
11540         Case False
11550           DoCmd.OpenReport strDocName, acViewPreview, , "[journal_USER] = '" & gstrJournalUser & "'", , CStr(lngRecs)
11560         End Select
11570         DoCmd.Maximize
11580         DoCmd.RunCommand acCmdFitToWindow

11590       End If  ' ** blnContinue.

11600       blnFormOpening = False

11610       DoCmd.Hourglass False

11620     End If
11630   End With

EXITP:
11640   Set rst = Nothing
11650   Set qdf1 = Nothing
11660   Set qdf2 = Nothing
11670   Set dbs = Nothing
11680   Exit Sub

ERRH:
11690   blnFormOpening = False
11700   DoCmd.Hourglass False
11710   DoCmd.Restore
11720   If Me.Visible = False Then
11730     Me.Visible = True
11740   End If
11750   Select Case ERR.Number
        Case Else
11760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11770   End Select
11780   Resume EXITP

End Sub

Private Sub cmdPostingJournalPreview_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPostingJournalPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11810   intRetVal = KeyCode

11820   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
11830   intShiftDown = (Shift And acShiftMask) > 0
11840   intAltDown = (Shift And acAltMask) > 0
11850   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11860   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11870     Select Case intRetVal
          Case vbKeyLeft
11880       With Me
11890         intRetVal = 0
11900         .cmdChecksClear.SetFocus
11910       End With
11920     Case vbKeyRight
11930       With Me
11940         intRetVal = 0
11950         .cmdChecksClear.SetFocus
11960       End With
11970     Case vbKeyUp
11980       With Me
11990         intRetVal = 0
12000         .cmdTransaction.SetFocus
12010       End With
12020     Case vbKeyDown
12030       With Me
12040         intRetVal = 0
12050         .cmdPostingJournalPrint.SetFocus
12060       End With
12070     End Select
12080   End If

EXITP:
12090   KeyCode = intRetVal
12100   Exit Sub

ERRH:
12110   intRetVal = 0
12120   Select Case ERR.Number
        Case Else
12130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12140   End Select
12150   Resume EXITP

End Sub

Private Sub cmdPostingJournalPrint_Click()
' ** Check for gblnAdmin vs. others.
' ** See also TransCount() in modPostUtilities.

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPostingJournalPrint_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim blnContinue As Boolean

12210   With Me
12220     If blnFormOpening = False Then

12230       DoCmd.Hourglass True
12240       DoEvents

12250       .TimerInterval = 0&

12260       blnContinue = True
12270       blnFormOpening = True

12280       gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
12290       UpdateImported  ' ** Procedure: Below.

12300       Set dbs = CurrentDb
12310       With dbs
12320         Select Case gblnAdmin
              Case True
                ' **Journal, with add'l fields, all journal_USER's.
12330           Set qdf = .QueryDefs("qryPost_Journal_02a")  '#jnox
12340         Case False
                ' ** qryPost_Journal_02b (Journal, with add'l fields, by specified journal_USER), by specified [usr].
12350           Set qdf = .QueryDefs("qryPost_Journal_02c")  '#jnox
12360           With qdf.Parameters
12370             ![usr] = gstrJournalUser
12380           End With
12390         End Select
12400         Set rst = qdf.OpenRecordset
12410         With rst
12420           If .BOF = True And .EOF = True Then
12430             blnContinue = False
12440             DoCmd.Hourglass False
12450             Select Case gblnAdmin
                  Case True
12460               MsgBox "There are no entries in the posting journal.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
12470             Case False
12480               MsgBox "There are no entries in the posting journal" & vbCrLf & _
                      "for User: " & gstrJournalUser & ".", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
12490             End Select
12500           End If
12510           .Close
12520         End With  ' ** rst.
12530         Set rst = Nothing
12540         Set qdf = Nothing
12550         If blnContinue = True Then
12560           Select Case gblnAdmin
                Case True
                  ' ** qryPost_Journal_02a (Journal, with add'l fields; all journal_USER's), grouped by curr_id.
12570             Set qdf = .QueryDefs("qryPost_Journal_02e")
12580           Case False
                  ' ** qryPost_Journal_02c (qryPost_Journal_02b (Journal, with add'l fields;
                  ' ** by specified journal_USER), by specified [usr]), grouped by curr_id.
12590             Set qdf = .QueryDefs("qryPost_Journal_02f")
12600             With qdf.Parameters
12610               ![usr] = gstrJournalUser
12620             End With
12630           End Select
12640           Set rst = qdf.OpenRecordset
12650           With rst
12660             .MoveLast
12670             lngRecs = .RecordCount
12680             .MoveFirst
12690             If lngRecs = 1 Then
12700               If ![curr_code] = "USD" Then
12710                 strDocName = "rptPostingJournal_Classic"
12720               Else
12730                 strDocName = "rptPostingJournal_Classic_ForEx"
12740               End If
12750             Else
12760               strDocName = "rptPostingJournal_Classic_ForEx"
12770             End If
12780             .Close
12790           End With  ' ** rst.
12800           Set rst = Nothing
12810           Set qdf = Nothing
12820         End If  ' ** blnContinue.
12830         .Close
12840       End With  ' ** dbs.

12850       If blnContinue = True Then

              ' ** Determine whether 'Posted By' should be included in transaction description.
12860         gblnSingleUser = IsSingleUser  ' ** Module Function: modSecurityFunctions.

12870         gstrReportCallingForm = vbNullString  ' ** When vbNullString, leaves this form .Visible = True.
12880         DoCmd.Hourglass False
12890         If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
12900           DoCmd.OpenReport strDocName, acViewPreview
12910           DoCmd.Maximize
12920           DoCmd.RunCommand acCmdFitToWindow
12930         Else
12940           Select Case gblnAdmin
                Case True
                  '##GTR_Ref: rptPostingJournal_Classic
                  '##GTR_Ref: rptPostingJournal_Classic_ForEx
12950             DoCmd.OpenReport strDocName, acViewNormal
12960           Case False
12970             DoCmd.OpenReport strDocName, acViewNormal, , "[journal_USER] = '" & gstrJournalUser & "'"
12980           End Select
12990         End If

13000       End If

13010       blnFormOpening = False

13020       DoCmd.Hourglass False

13030     End If
13040   End With

EXITP:
13050   Set rst = Nothing
13060   Set qdf = Nothing
13070   Set dbs = Nothing
13080   Exit Sub

ERRH:
13090   blnFormOpening = False
13100   DoCmd.Hourglass False
13110   Select Case ERR.Number
        Case 2202  ' ** You must set a default printer before you design, print, or preview.
13120     MsgBox "There does not appear to be a default printer defined." & vbCrLf & _
            "Please define a default printer before trying to print this report.", _
            vbCritical + vbOKOnly, "Default Printer Required"
13130   Case Else
13140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13150   End Select
13160   Resume EXITP

End Sub

Private Sub cmdPostingJournalPrint_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPostingJournalPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13210   intRetVal = KeyCode

13220   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
13230   intShiftDown = (Shift And acShiftMask) > 0
13240   intAltDown = (Shift And acAltMask) > 0
13250   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13260   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13270     Select Case intRetVal
          Case vbKeyLeft
13280       With Me
13290         intRetVal = 0
13300         .cmdCheckMemos.SetFocus
13310       End With
13320     Case vbKeyRight
13330       With Me
13340         intRetVal = 0
13350         .cmdCheckMemos.SetFocus
13360       End With
13370     Case vbKeyUp
13380       With Me
13390         intRetVal = 0
13400         .cmdPostingJournalPreview.SetFocus
13410       End With
13420     Case vbKeyDown
13430       With Me
13440         intRetVal = 0
13450         .cmdPost.SetFocus
13460       End With
13470     End Select
13480   End If

EXITP:
13490   KeyCode = intRetVal
13500   Exit Sub

ERRH:
13510   intRetVal = 0
13520   Select Case ERR.Number
        Case Else
13530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13540   End Select
13550   Resume EXITP

End Sub

Private Sub cmdPostingJournalPrint_LostFocus()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPostingJournalPrint_LostFocus"

13610   With Me
13620     If (.GoToReport_arw_menu_l_img.Visible = True) Then
13630       If (.GoToReport_arw_menu_l_img.Top = .cmdPostingJournalPrint.Top + (3& * lngTpp)) Then
13640         .GoToReport_arw_menu_l_img.Visible = False
13650       End If
13660     End If
13670   End With

EXITP:
13680   Exit Sub

ERRH:
13690   Select Case ERR.Number
        Case Else
13700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13710   End Select
13720   Resume EXITP

End Sub

Private Sub cmdPost_Click()
' ** Check for gblnAdmin vs. others.

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPost_Click"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef
        Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim grp As DAO.Group, usr As DAO.User
        Dim datStatementDate As Date
        Dim strUser As String, strPostType As String, lngRecsToPost As Long, strAdminResponse As String, strUserPosted As String
        Dim lngRecs1 As Long, lngRecs2 As Long
        Dim blnContinue As Boolean, blnBadCombo As Boolean, blnHasLiability As Boolean
        Dim lngX As Long
        Dim intRetVal As Integer

        'Const POST_NON As String = "0"
        Const POST_ALL As String = "1"
        'Const POST_SYS As String = "2"
        Const POST_USR As String = "3"

13810   With Me
13820     If blnFormOpening = False Then

13830       DoCmd.Hourglass True
13840       DoEvents

13850       .TimerInterval = 0&

13860       blnContinue = True
13870       blnBadCombo = False  ' ** Unless proven otherwise.
13880       blnHasLiability = False
13890       blnFormOpening = True

13900       UpdateImported  ' ** Procedure: Below.

13910       Set dbs = CurrentDb
13920       With dbs

13930         Set rst1 = dbs.OpenRecordset("Statement Date", dbOpenDynaset, dbReadOnly)
13940         With rst1
13950           .MoveFirst  ' ** Must be just one record in this table.
13960           datStatementDate = ![Statement_Date]
13970           .Close
13980         End With

              ' ** Journal, grouped, with cnt, by specified [sdat].
13990         Set qdf1 = .QueryDefs("qryPost_Journal_05")
14000         With qdf1.Parameters
14010           ![sdat] = datStatementDate
14020         End With
14030         Set rst1 = qdf1.OpenRecordset
14040         With rst1
14050           .MoveFirst  ' ** Must be exactly one record with exactly one field returned.
14060           If ![cnt] > 0 Then
14070             DoCmd.Hourglass False
14080             Beep
14090             If MsgBox("You are trying to post transaction(s) dated prior to the date you last ran statements:" & vbCrLf & _
                      "  " & Format(datStatementDate, "mm/dd/yyyy") & _
                      vbCrLf & vbCrLf & "This posting may change the beginning balance for your next statement." & _
                      vbCrLf & "You may want to run statements for the changed accounts again. " & _
                      "See 'Statements' in the manual for more information." & _
                      vbCrLf & vbCrLf & "Do you really want to post these transactions?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Warning: Possible Data Conflict") = vbNo Then
14100               blnContinue = False
14110             Else
14120               DoCmd.Hourglass True
14130               DoEvents
14140             End If
14150           End If
14160           .Close
14170         End With

14180         .Close
14190       End With

14200       If blnContinue = True Then

              ' ** This section is duplicated later on in PostTransactions(), but we
              ' ** need to know now in which subset of records to make the cross-check.

              ' ** Security Check.
14210         For Each grp In DBEngine.Workspaces(0).Groups
14220           If grp.Name = SGRP_ADMINS Then
                  ' ** Any User found in Admins (¹) Group permitted access.
14230             For Each usr In grp.Users
14240               If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
14250                 gblnAdmin = True
14260                 Exit For
14270               End If
14280             Next
14290           End If
14300         Next

14310         strAdminResponse = vbNullString: strUserPosted = vbNullString
14320         lngRecsToPost = DCount("ID", "journal")
14330         lngRecs1 = DCount("journal_USER", "journal", "journal_USER = CurrentUser")  ' ** Internal Access Function: Trust Accountant login.
14340         If lngRecs1 = lngRecsToPost Then
                ' ** If the only entries in the Journal are the current user's,
                ' ** don't bother checking for Admin status.
14350           strUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
14360           strPostType = POST_USR
14370         Else
                ' ** If multiple users are in the Journal, check for Admin status.
14380           Select Case gblnAdmin
                Case True
                  ' ** Admin will need to make the choice here, rather than later in PostTransactions().
14390             strUser = SelectUser  ' ** Public Function: modPostUtilities.
14400             strAdminResponse = strUser
14410             If strUser = vbNullString Then
14420               blnContinue = False
14430             Else
14440               Select Case strUser
                    Case "All"
14450                 strUser = vbNullString
14460                 strPostType = POST_ALL
14470               Case Else
14480                 strUserPosted = strUser
14490                 strPostType = POST_USR
14500               End Select
14510             End If
14520           Case False
14530             strUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
14540             strPostType = POST_USR
14550           End Select
14560         End If

14570       End If  ' ** blnContinue.

14580       If blnContinue = True Then

14590         Set dbs = CurrentDb
14600         With dbs

14610           Select Case strPostType
                Case POST_ALL
                  ' ** Journal, grouped by accountno, assetno, just 'Sold', 'Withdrawn', AssetType <> '60', '80', '81', with NumRecs, all Users.
                  ' ** Money Market Funds, Interest Bearing - Own, Interest Bearing - Other.  Now with 'Liability', 'Liability (-)'!
14620             Set qdf1 = .QueryDefs("qryPost_01a")
14630           Case POST_USR
                  ' ** Journal, grouped by accountno, assetno, just 'Sold', 'Withdrawn', AssetType <> '60', '80', '81', with NumRecs; by specified [usr].
                  ' ** Money Market Funds, Interest Bearing - Own, Interest Bearing - Other.  Now with 'Liability', 'Liability (-)'!
14640             Set qdf1 = .QueryDefs("qryPost_01b")
14650             With qdf1.Parameters
14660               ![usr] = strUser
14670             End With
14680           End Select
14690           Set rst1 = qdf1.OpenRecordset()
14700           With rst1
14710             If .BOF = True And .EOF = True Then
14720               lngRecs1 = 0&
14730             Else
14740               .MoveLast
14750               lngRecs1 = .RecordCount
14760               .MoveFirst
14770             End If
14780           End With

14790           Select Case strPostType
                Case POST_ALL
                  ' ** Journal, grouped by accountno, assetno, just 'Purchase', 'Deposit', with NumRecs, all Users.  Now with 'Liability', 'Liability (+)'!
14800             Set qdf2 = .QueryDefs("qryPost_02a")
14810           Case POST_USR
                  ' ** Journal, grouped by accountno, assetno, just 'Purchase', 'Deposit', with NumRecs, by specified [usr].  Now with 'Liability', 'Liability (+)'!
14820             Set qdf2 = .QueryDefs("qryPost_02b")
14830             With qdf2.Parameters
14840               ![usr] = strUser
14850             End With
14860           End Select
14870           Set rst2 = qdf2.OpenRecordset()
14880           With rst2
14890             If .BOF = True And .EOF = True Then
14900               lngRecs2 = 0&
14910             Else
14920               .MoveLast
14930               lngRecs2 = .RecordCount
14940               .MoveFirst
14950             End If
14960           End With

14970           If (lngRecs1 > 0) And (lngRecs2 > 0) Then

14980             For lngX = 1& To lngRecs1
14990               If rst1![assettype] <> "60" And rst1![assettype] <> "80" And rst1![assettype] <> "81" Then
15000                 rst2.FindFirst "[accountno] = '" & rst1![accountno] & "' AND [assetno] = " & CStr(rst1![assetno])
15010                 Select Case rst2.NoMatch
                      Case True
15020                   rst2.MoveFirst
15030                   If lngX < lngRecs1 Then rst1.MoveNext
15040                 Case False
15050                   blnBadCombo = True
15060                   If rst1![hasliability] = True Or rst2![hasliability] = True Then
15070                     blnHasLiability = True
15080                   End If
15090                   Exit For
15100                 End Select
15110               Else
15120                 If lngX < lngRecs1 Then rst1.MoveNext
15130               End If
15140             Next

15150             If blnBadCombo = True Then
15160               DoCmd.Hourglass False
                    ' ** There is at least one account with multiple transactions for an asset of the type prohibited.
15170               strMsg = "You cannot have deposit" & IIf(blnHasLiability = False, " or purchase", ", purchase or liability (+)") & _
                      " transactions" & vbCrLf & _
                      "for the same asset in a single account posted at" & vbCrLf & _
                      "the same time as sale" & IIf(blnHasLiability = False, " or withdrawal", ", withdrawal or liability (-)") & _
                      " transactions." & vbCrLf & vbCrLf & _
                      "Cancel one or more of these transactions" & vbCrLf & "and post separately."
15180               intStyle = vbCritical + vbOKOnly
15190               strTitle = Left(("Multiple Account-Asset Transactions Found" & Space(55)), 55)
15200               blnContinue = False
15210               msgResponse = MsgBox(strMsg, intStyle, strTitle)
15220             End If

15230           End If

15240           rst1.Close
15250           rst2.Close

15260           .Close
15270         End With  ' ** dbs.

15280       End If  ' ** blnContinue.

            ' ** OK to try to continue.
15290       If blnContinue = True Then

15300         If strAdminResponse = vbNullString Then
15310           DoCmd.Hourglass False
                ' ** Admin should only get a total of 2 messages.
15320           strMsg = "Are you sure you want to post journal entries at this time?"
15330           intStyle = vbQuestion + vbYesNo + vbDefaultButton1
15340           strTitle = "Post Journal Entries"
15350           msgResponse = MsgBox(strMsg, intStyle, strTitle)
15360         Else
15370           msgResponse = vbYes
15380         End If

15390         If msgResponse = vbYes Then
15400           DoCmd.Hourglass True
15410           DoEvents
15420           If strAdminResponse = vbNullString Then
15430             DoCmd.Hourglass False
                  ' ** Admin should only get a total of 2 messages.
15440             strMsg = "Have you checked these transactions in the posting journal?"
15450             intStyle = vbQuestion + vbYesNo + vbDefaultButton1
15460             strTitle = "Post Journal Entries"
15470             msgResponse = MsgBox(strMsg, intStyle, strTitle)
15480           Else
15490             msgResponse = vbYes
15500           End If
15510           If msgResponse <> vbYes Then
15520             blnContinue = False
15530           End If  ' ** msgResponse.
15540         Else
15550           blnContinue = False
15560         End If  ' ** msgResponse.

15570       End If  ' ** blnContinue.

15580       If blnContinue = True Then
15590         DoCmd.Hourglass True
15600         DoEvents
15610         Set dbs = CurrentDb
15620         With dbs
15630           Select Case strPostType
                Case POST_ALL
                  ' ** Journal, just PrintCheck = True, all Users.
15640             Set qdf1 = .QueryDefs("qryPost_Journal_06a")
15650             strMsg = "the Print Check items!"
15660           Case POST_USR
                  ' ** Journal, just PrintCheck = True, by specified [usr].
15670             Set qdf1 = .QueryDefs("qryPost_Journal_06b")
15680             With qdf1.Parameters
15690               ![usr] = strUser
15700             End With
15710             strMsg = "the Print Check items" & vbCrLf & "for User: " & strUser & "!"
15720           End Select

15730           Set rst1 = qdf1.OpenRecordset
15740           With rst1
15750             If .BOF = True And .EOF = True Then
                    ' ** All's well.
15760             Else
15770               DoCmd.Hourglass False
15780               msgResponse = MsgBox("Warning! Posting transactions will clear " & strMsg & vbCrLf & vbCrLf & _
                      "Do you want to continue?", vbQuestion + vbYesNo, "Clear Print Check")
15790               If msgResponse <> vbYes Then
15800                 blnContinue = False
15810               Else
15820                 DoCmd.Hourglass True
15830                 DoEvents
15840               End If
15850             End If
15860             .Close
15870           End With

15880           .Close
15890         End With  ' ** dbs.
15900       End If  ' ** blnContinue.

15910       If blnContinue = True Then
              ' ** If a new account had postings prior to its predate, this will move the Balance table back.
15920         CheckDates  ' ** Module Function: modDateCheckFuncs.
15930         Pref_RemNot "frmJournal", "Account", "LastAcctNo"  ' ** Module Procedure: modPreferenceFuncs.
15940         Pref_RemNot "frmJournal_Columns", "Account", "LastAcctNo"  ' ** Module Procedure: modPreferenceFuncs.
15950         intRetVal = PostTransactions(strAdminResponse)  ' ** Module Procedure: modPostUtilities.
              ' ** Return codes:
              ' **   0  POST_INIT     Initialized.
              ' **   1  POST_NOTRANS  No transactions.
              ' **   2  POST_CANCEL   User Canceled.
              ' **   3  POST_NULL     Null asset pricing.
              ' **   4  POST_ROLLBK   Problem, caused rollback.
              ' **   5  POST_ERROR    Other Error, from error handler.
              ' **   6  POST_DONE     Post successful.
              ' ** Recalculate glngUserCntLedger.
15960         IsSingleUser True  ' ** Module Function: modSecurityFunctions.
15970         blnUpdateTransCnt = True
15980         Me.TimerInterval = 500&
15990       End If  ' ** blnContinue.

16000       blnFormOpening = False

16010       DoCmd.Hourglass False

16020     End If
16030   End With

EXITP:
16040   Set usr = Nothing
16050   Set grp = Nothing
16060   Set rst1 = Nothing
16070   Set rst2 = Nothing
16080   Set qdf1 = Nothing
16090   Set qdf2 = Nothing
16100   Set dbs = Nothing
16110   Exit Sub

ERRH:
16120   blnFormOpening = False
16130   DoCmd.Hourglass False
16140   Select Case ERR.Number
        Case Else
16150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16160   End Select
16170   Resume EXITP

End Sub

Private Sub cmdPost_KeyDown(KeyCode As Integer, Shift As Integer)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPost_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16210   intRetVal = KeyCode

16220   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
16230   intShiftDown = (Shift And acShiftMask) > 0
16240   intAltDown = (Shift And acAltMask) > 0
16250   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16260   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16270     Select Case intRetVal
          Case vbKeyLeft
16280       With Me
16290         intRetVal = 0
16300         .cmdCheckReconcile.SetFocus
16310       End With
16320     Case vbKeyRight
16330       With Me
16340         intRetVal = 0
16350         .cmdCheckReconcile.SetFocus
16360       End With
16370     Case vbKeyUp
16380       With Me
16390         intRetVal = 0
16400         .cmdPostingJournalPrint.SetFocus
16410       End With
16420     Case vbKeyDown
16430       With Me
16440         intRetVal = 0
16450         .cmdClearPostingJournal.SetFocus
16460       End With
16470     End Select
16480   End If

EXITP:
16490   KeyCode = intRetVal
16500   Exit Sub

ERRH:
16510   intRetVal = 0
16520   Select Case ERR.Number
        Case Else
16530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16540   End Select
16550   Resume EXITP

End Sub

Private Sub cmdClearPostingJournal_Click()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearPostingJournal_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCnt1 As Long, lngCnt2 As Long, lngCnt3 As Long, lngCnt4 As Long, lngCnt5 As Long, lngNullUsers As Long
        Dim lngUsrs1 As Long, arr_varUsr1 As Variant
        Dim lngUsrs2 As Long, arr_varUsr2 As Variant
        Dim intMode As Integer
        Dim strUserCleared1 As String, strUserCleared2 As String
        Dim msgResponse2 As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long

        ' ** Array: arr_varUsr1().
        Const U_USR As Integer = 0
        Const U_CNT As Integer = 1

16610   With Me
16620     If blnFormOpening = False Then

16630       DoCmd.Hourglass True
16640       DoEvents

16650       .TimerInterval = 0&

16660       blnFormOpening = True

16670       UpdateTransCnt  ' ** Function: Below.
16680       gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
16690       lngNullUsers = 0&: strUserCleared1 = vbNullString: strUserCleared2 = vbNullString
16700       lngCnt1 = 0&: lngCnt2 = 0&: lngCnt3 = 0&: lngCnt4 = 0&: lngCnt5 = 0&

16710       Set dbs = CurrentDb
16720       With dbs

              ' ** PostingDate table, Username = Null.
16730         varTmp00 = DCount("*", "qryPostingDate_03")
16740         If IsNull(varTmp00) = False Then lngCnt3 = varTmp00
              ' ** Journal, grouped by journal_USER, for journal_USER = Null, with cnt.
16750         varTmp00 = DLookup("[cnt]", "qryPostingDate_10")
16760         If IsNull(varTmp00) = False Then lngCnt4 = varTmp00
              ' ** Journal, grouped by transdate, for transdate = Null, with cnt.
16770         varTmp00 = DLookup("[cnt]", "qryPostingDate_11")
16780         If IsNull(varTmp00) = False Then lngCnt5 = varTmp00

16790         Select Case gblnAdmin
              Case True  ' ** Admins group user.
                ' ** Journal, grouped, with cnt, just journal_USER = Null.
16800           varTmp00 = DLookup("[cnt]", "qryJournal_Clear_04")
16810           If IsNull(varTmp00) = False Then
16820             If varTmp00 > 0 Then
                    ' ** If Null users exist, assign them to the administrator, so that they can see them and delete them.
16830               lngNullUsers = CLng(varTmp00)
                    ' ** Update Journal, for journal_USER = Null, by specified [usr].
16840               Set qdf = .QueryDefs("qryJournal_Clear_05")
16850               With qdf.Parameters
16860                 ![usr] = gstrJournalUser
16870               End With
16880               qdf.Execute
16890               Set qdf = Nothing
16900             End If
16910           End If
                ' ** Journal, grouped by journal_USER, with cnt, all Users.
16920           Set qdf = .QueryDefs("qryJournal_Clear_01c")
16930           Set rst = qdf.OpenRecordset
16940           With rst
16950             If .BOF = True And .EOF = True Then
16960               lngCnt1 = 0&
16970             Else
16980               .MoveLast
16990               lngUsrs1 = .RecordCount
17000               .MoveFirst
17010               arr_varUsr1 = .GetRows(lngUsrs1)
                    ' *************************************************
                    ' ** Array: arr_varUsr1()
                    ' **
                    ' **   Field  Element  Name            Constant
                    ' **   =====  =======  ==============  ==========
                    ' **     1       0     journal_USER    U_USR
                    ' **     2       1     cnt             U_CNT
                    ' **
                    ' *************************************************
17020               lngCnt1 = 0&
17030               For lngX = 0& To (lngUsrs1 - 1&)
17040                 lngCnt1 = lngCnt1 + arr_varUsr1(U_CNT, lngX)
17050               Next
17060             End If
17070             .Close
17080           End With
17090           Set rst = Nothing
17100           Set qdf = Nothing
                ' ** tblJournal_Column, grouped by journal_USER, just uncommitted, with cnt, all Users.
17110           Set qdf = .QueryDefs("qryJournal_Clear_01i")
17120           Set rst = qdf.OpenRecordset
17130           With rst
17140             If .BOF = True And .EOF = True Then
17150               lngCnt2 = 0&
17160             Else
17170               .MoveLast
17180               lngUsrs2 = .RecordCount
17190               .MoveFirst
17200               arr_varUsr2 = .GetRows(lngUsrs2)
17210               lngCnt2 = 0&
17220               For lngX = 0& To (lngUsrs2 - 1&)
17230                 lngCnt2 = lngCnt2 + arr_varUsr2(U_CNT, lngX)
17240               Next
17250             End If
17260             .Close
17270           End With
17280           Set rst = Nothing
17290           Set qdf = Nothing
17300         Case False  ' ** Non-Admins group user.
                ' ** Journal, grouped, with cnt, by specified [usr].
17310           Set qdf = .QueryDefs("qryJournal_Clear_01b")
17320           With qdf.Parameters
17330             ![usr] = gstrJournalUser
17340           End With
17350           Set rst = qdf.OpenRecordset
17360           With rst
17370             If .BOF = True And .EOF = True Then
17380               lngCnt1 = 0&
17390             Else
17400               .MoveFirst
17410               lngCnt1 = ![cnt]
17420             End If
17430             .Close
17440           End With
17450           Set rst = Nothing
17460           Set qdf = Nothing
                ' **tblJournal_Column, grouped, just uncommitted, with cnt, by specified [usr].
17470           Set qdf = .QueryDefs("qryJournal_Clear_01h")
17480           With qdf.Parameters
17490             ![usr] = gstrJournalUser
17500           End With
17510           Set rst = qdf.OpenRecordset
17520           With rst
17530             If .BOF = True And .EOF = True Then
17540               lngCnt2 = 0&
17550             Else
17560               .MoveFirst
17570               lngCnt2 = ![cnt]
17580             End If
17590             .Close
17600           End With
17610           Set rst = Nothing
17620           Set qdf = Nothing
17630         End Select

17640         .Close
17650       End With  ' ** dbs.

17660       If lngCnt1 > 0& Or lngCnt2 > 0& Then

17670         gstrCrtRpt_NetLoss = vbNullString  ' ** Borrowing this variable from the Court Reports.
17680         Select Case gblnAdmin
              Case True  ' ** Admins group user.
17690           If lngUsrs1 > 0& And lngUsrs2 = 0& Then
17700             If lngUsrs1 = 1& Then
17710               intMode = 1  ' ** J only (1 user, Journal).
17720             ElseIf lngUsrs1 > 1& Then
17730               intMode = 2  ' ** mJ only (multi-user, Journal).
17740             End If
17750           ElseIf lngUsrs1 = 0& And lngUsrs2 > 0& Then
17760             If lngUsrs2 = 1& Then
17770               intMode = 3  ' ** U only (1 user, Uncommitted)
17780             ElseIf lngUsrs2 > 1& Then
17790               intMode = 4  ' ** mU only (multi-user, Uncommitted)
17800             End If
17810           ElseIf lngUsrs1 > 0& And lngUsrs2 > 0& Then
17820             If lngUsrs1 = 1& And lngUsrs2 = 1& Then
17830               intMode = 5  ' ** J & U
17840             ElseIf lngUsrs1 > 1& And lngUsrs2 = 1& Then
17850               intMode = 6  ' ** mJ & U
17860             ElseIf lngUsrs1 = 1& And lngUsrs2 > 1& Then
17870               intMode = 7  ' ** J & mU
17880             ElseIf lngUsrs1 > 1& And lngUsrs2 > 1& Then
17890               intMode = 8  ' ** mJ & mU
17900             End If
17910           End If
17920         Case False  ' ** Non-Admins group user.
17930           If lngCnt1 > 0& And lngCnt2 = 0& Then
17940             intMode = 9  ' ** J only
17950           ElseIf lngCnt1 = 0& And lngCnt2 > 0& Then
17960             intMode = 10  ' ** U only
17970           ElseIf lngCnt1 > 0 And lngCnt2 > 0& Then
17980             intMode = 11  ' ** J & U
17990           End If
18000         End Select

              ' ** gblnAdmin:
              ' **   True
              ' **     intMode = 1  : (lngCnt1 > 0& And lngUsrs1 = 1&) And (lngCnt2 = 0& And lngUsrs2 = 0&) J only   1 form
              ' **     intMode = 2  : (lngCnt1 > 0& And lngUsrs1 > 1&) And (lngCnt2 = 0& And lngUsrs2 = 0&) mJ only  1 form
              ' **     intMode = 3  : (lngCnt1 = 0& And lngUsrs1 = 0&) And (lngCnt2 > 0& And lngUsrs2 = 1&) U only   1 form
              ' **     intMode = 4  : (lngCnt1 = 0& And lngUsrs1 = 0&) And (lngCnt2 > 0& And lngUsrs2 > 1&) mU only  1 form
              ' **     intMode = 5  : (lngCnt1 > 0& And lngUsrs1 = 1&) And (lngCnt2 > 0& And lngUsrs2 = 1&) J & U    1 form
              ' **     intMode = 6  : (lngCnt1 > 0& And lngUsrs1 > 1&) And (lngCnt2 > 0& And lngUsrs2 = 1&) mJ & U   2 forms
              ' **     intMode = 7  : (lngCnt1 > 0& And lngUsrs1 = 1&) And (lngCnt2 > 0& And lngUsrs2 > 1&) J & mU   2 forms
              ' **     intMode = 8  : (lngCnt1 > 0& And lngUsrs1 > 1&) And (lngCnt2 > 0& And lngUsrs2 > 1&) mJ & mU  2 forms
              ' **   False
              ' **     intMode = 9  : (lngCnt1 > 0& And lngCnt2 = 0&) J only
              ' **     intMode = 10 : (lngCnt1 = 0& And lngCnt2 > 0&) U only
              ' **     intMode = 11 : (lngCnt1 > 0& And lngCnt2 > 0&) J & U

18010         Select Case intMode
              Case 1, 3, 5, 7, 9, 10, 11  ' ** frmMenu_Post_Clear opened first or only.
18020           strTmp01 = gstrJournalUser
18030           If lngUsrs1 = 1& Then
18040             If arr_varUsr1(U_USR, 0) <> gstrJournalUser Then
                    ' ** One User, but the current User isn't it!
18050               strTmp01 = arr_varUsr1(U_USR, 0)
18060             End If
18070           End If
18080           DoCmd.Hourglass False
18090           gdblCrtRpt_CostTot = 0#  ' ** Borrowing this variable from the Court Reports.
18100           strDocName = "frmMenu_Post_Clear"
18110           DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME & "~" & CStr(intMode) & "~" & _
                  IIf(gblnAdmin = True, "True", "False") & "~" & strTmp01 & "~" & _
                  CStr(lngCnt1) & "~" & CStr(lngCnt2) & "~" & CStr(lngUsrs1) & "~" & CStr(lngUsrs2)
18120           DoEvents
18130           Select Case gdblCrtRpt_CostTot
                Case 0#
18140             msgResponse = vbNo
18150             msgResponse2 = vbNo
18160           Case 1#
18170             msgResponse = vbYes
18180             strUserCleared1 = strTmp01
18190             msgResponse2 = vbNo
18200           Case 2#
18210             msgResponse = vbNo
18220             msgResponse2 = vbYes
18230             strUserCleared2 = strTmp01
18240           Case 3#
18250             msgResponse = vbYes
18260             strUserCleared1 = strTmp01
18270             msgResponse2 = vbYes
18280             strUserCleared2 = strTmp01
18290           End Select
18300         Case 2, 6, 8  ' ** frmMenu_Post_Clear_Multi opened first or only.
18310           DoCmd.Hourglass False
18320           gstrJournalUser = vbNullString  ' ** Don't clear this on the form!
18330           strDocName = "frmMenu_Post_Clear_Multi"
18340           DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME & "~" & "ClearJournal"
18350           If gstrJournalUser = vbNullString Then
18360             msgResponse = vbNo
18370           Else
18380             msgResponse = vbYes
18390             strUserCleared1 = gstrJournalUser
18400           End If
18410           msgResponse2 = vbNo
18420         End Select

18430         Select Case intMode
              Case 6  ' ** frmMenu_Post_Clear opened second.
18440           DoCmd.Hourglass False
18450           gdblCrtRpt_CostTot = 0#  ' ** Borrowing this variable from the Court Reports.
18460           strDocName = "frmMenu_Post_Clear"
18470           DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME & "~" & CStr(intMode) & "~" & _
                  IIf(gblnAdmin = True, "True", "False") & "~" & gstrJournalUser & "~" & _
                  CStr(lngCnt1) & "~" & CStr(lngCnt2) & "~" & CStr(lngUsrs1) & "~" & CStr(lngUsrs2)
18480           Select Case gdblCrtRpt_CostTot
                Case 0#
18490             msgResponse2 = vbNo
18500           Case 2#
18510             msgResponse2 = vbYes
18520             strUserCleared2 = gstrJournalUser
18530           End Select
18540         Case 4, 7, 8  ' ** frmMenu_Post_Clear_Multi opened second or only.
18550           DoCmd.Hourglass False
18560           gstrCrtRpt_NetLoss = vbNullString  ' ** Don't clear this on the form!
18570           strDocName = "frmMenu_Post_Clear_Multi"
18580           DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME & "~" & "ClearUncommitted"
18590           If gstrCrtRpt_NetLoss = vbNullString Then
18600             msgResponse2 = vbNo
18610           Else
18620             msgResponse2 = vbYes
18630             strUserCleared2 = gstrCrtRpt_NetLoss
18640           End If
18650         End Select

18660         If msgResponse = vbYes Then
18670           DoCmd.Hourglass True
18680           DoEvents
18690           strTmp01 = Empty
18700           strTmp01 = gstrJournalUser
18710           If lngUsrs1 = 1& Then
18720             If arr_varUsr1(U_USR, 0) <> gstrJournalUser Then
                    ' ** One User, but the current User isn't it!
18730               strTmp01 = arr_varUsr1(U_USR, 0)
18740             End If
18750           End If
18760           Set dbs = CurrentDb
18770           With dbs
18780             Select Case strTmp01
                  Case "All"
                    ' ** Empty Journal.
18790               Set qdf = .QueryDefs("qryJournal_Clear_02")
18800             Case Else
                    ' ** Empty Journal, by specified [usr].
18810               Set qdf = .QueryDefs("qryJournal_Clear_03")
18820               With qdf.Parameters
18830                 ![usr] = strTmp01
18840               End With
18850             End Select
18860             qdf.Execute
18870             Set qdf = Nothing
18880             .Close
18890           End With
18900           Set dbs = Nothing

18910           UpdateTransCnt  ' ** Function: Below.
18920           DoEvents

18930           If msgResponse2 = vbNo Then
18940             DoCmd.Hourglass False
18950             Select Case gblnAdmin
                  Case True
18960               Select Case strTmp01
                    Case "All"
18970                 MsgBox "Posting Journal Cleared!", vbInformation + vbOKOnly, ("Journal Cleared" & Space(40))
18980               Case Else
18990                 If strUserCleared1 = vbNullString Then strUserCleared1 = gstrJournalUser
19000                 MsgBox "Posting Journal Cleared for User: " & strUserCleared1 & ".", _
                        vbInformation + vbOKOnly, ("Journal Cleared" & Space(40))
19010               End Select
19020             Case False
19030               If strUserCleared1 = vbNullString Then strUserCleared1 = gstrJournalUser
19040               MsgBox "Posting Journal Cleared for User: " & strUserCleared1 & ".", _
                      vbInformation + vbOKOnly, ("Journal Cleared" & Space(40))
19050             End Select
19060           End If
19070         Else
19080           If msgResponse2 <> vbYes Then
19090             DoCmd.Hourglass False
19100           End If
19110         End If

19120         If msgResponse2 = vbYes Then
19130           If gstrCrtRpt_NetLoss <> vbNullString Then gstrJournalUser = gstrCrtRpt_NetLoss
19140           DoCmd.Hourglass True
19150           DoEvents
19160           Set dbs = CurrentDb
19170           With dbs
19180             Select Case gstrJournalUser
                  Case "All"
                    ' ** Empty tblJournal_Column.
19190               Set qdf = .QueryDefs("qryJournal_Clear_06")
19200             Case Else
                    ' ** Empty tblJournal_Column, by specified [usr].
19210               Set qdf = .QueryDefs("qryJournal_Clear_07")
19220               With qdf.Parameters
19230                 ![usr] = gstrJournalUser
19240               End With
19250             End Select
19260             qdf.Execute
19270             Set qdf = Nothing
19280             .Close
19290           End With
19300           Set dbs = Nothing

19310           UpdateTransCnt  ' ** Function: Below.
19320           DoEvents
19330           DoCmd.Hourglass False

19340           Select Case gblnAdmin
                Case True
19350             Select Case msgResponse
                  Case vbYes
19360               MsgBox "The Posting Journal and Uncommitted Entries Cleared!", _
                      vbInformation + vbOKOnly, ("Journal And Uncommitted Cleared" & Space(40))
19370             Case vbNo
19380               MsgBox "Uncommitted Journal Entries Cleared!", vbInformation + vbOKOnly, ("Uncommitted Cleared" & Space(40))
19390             End Select
19400           Case False
19410             Select Case msgResponse
                  Case vbYes
19420               If strUserCleared2 = vbNullString Then strUserCleared2 = gstrJournalUser
19430               MsgBox "The Posting Journal and Uncommitted Entries Cleared for User: " & strUserCleared2 & ".", _
                      vbInformation + vbOKOnly, ("Journal And Uncommitted Cleared" & Space(40))
19440             Case vbNo
19450               If strUserCleared2 = vbNullString Then strUserCleared2 = gstrJournalUser
19460               MsgBox "Uncommitted Journal Entries Cleared for User: " & strUserCleared2 & ".", _
                      vbInformation + vbOKOnly, ("Uncommitted Cleared" & Space(40))
19470             End Select
19480           End Select
19490         Else
19500           DoCmd.Hourglass False
19510         End If

19520         DoCmd.Hourglass True
19530         DoEvents

19540         If lngCnt3 > 0& Or lngCnt4 > 0& Or lngCnt5 > 0& Then
19550           If lngCnt3 > 0& Then
                  ' ** Null Username in PostingDate.
19560             CheckDates  ' ** Module Function: modDateCheckFuncs.
19570           End If
19580           If lngCnt4 > 0& Then
19590             Set dbs = CurrentDb
19600             With dbs
                    ' ** Delete Journal, for journal_USER = Null.
19610               Set qdf = .QueryDefs("qryPostingDate_12")
19620               qdf.Execute
19630               Set qdf = Nothing
19640               .Close
19650             End With
19660             Set dbs = Nothing
19670           End If
19680           If lngCnt5 > 0& Then
                  ' ** Null transdate in Journal.
19690             Set dbs = CurrentDb
19700             With dbs
                    ' ** Delete Journal, for transdate = Null.
19710               Set qdf = .QueryDefs("qryPostingDate_13")
19720               qdf.Execute
19730               Set qdf = Nothing
19740               .Close
19750             End With
19760             Set dbs = Nothing
19770           End If
19780         End If

19790         DoCmd.Hourglass False

19800       Else

19810         If lngCnt3 > 0& Or lngCnt4 > 0& Or lngCnt5 > 0& Then
19820           If lngCnt3 > 0& Then
                  ' ** Null Username in PostingDate.
19830             CheckDates  ' ** Module Function: modDateCheckFuncs.
19840           End If
19850           If lngCnt4 > 0& Then
19860             Set dbs = CurrentDb
19870             With dbs
                    ' ** Delete Journal, for journal_USER = Null.
19880               Set qdf = .QueryDefs("qryPostingDate_12")
19890               qdf.Execute
19900               Set qdf = Nothing
19910               .Close
19920             End With
19930             Set dbs = Nothing
19940           End If
19950           If lngCnt5 > 0& Then
                  ' ** Null transdate in Journal.
19960             Set dbs = CurrentDb
19970             With dbs
                    ' ** Delete Journal, for transdate = Null.
19980               Set qdf = .QueryDefs("qryPostingDate_13")
19990               qdf.Execute
20000               Set qdf = Nothing
20010               .Close
20020             End With
20030             Set dbs = Nothing
20040           End If
20050           Pref_RemNot "frmJournal", "Account", "LastAcctNo"  ' ** Module Procedure: modPreferenceFuncs.
20060           Pref_RemNot "frmJournal_Columns", "Account", "LastAcctNo"  ' ** Module Procedure: modPreferenceFuncs.
20070           MsgBox "Posting Journal Cleared!", vbInformation + vbOKOnly, ("Journal Cleared" & Space(40))
20080         Else
20090           DoCmd.Hourglass False
20100           MsgBox "There are no pending transactions in the Posting Journal," & vbCrLf & _
                  "committed or uncommitted.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
20110         End If

20120       End If

20130       blnFormOpening = False

20140     End If
20150   End With

EXITP:
20160   Set rst = Nothing
20170   Set qdf = Nothing
20180   Set dbs = Nothing
20190   Exit Sub

ERRH:
20200   blnFormOpening = False
20210   DoCmd.Hourglass False
20220   Select Case ERR.Number
        Case Else
20230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20240   End Select
20250   Resume EXITP

End Sub

Private Sub cmdClearPostingJournal_KeyDown(KeyCode As Integer, Shift As Integer)

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearPostingJournal_KeyDown"

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
          Case vbKeyLeft
20370       With Me
20380         intRetVal = 0
20390         Select Case strMenuArrow
              Case "cmdAccrue_KeyDown"
20400           .cmdAccrue.SetFocus
20410         Case "cmdFeeCalculations_KeyDown"
20420           .cmdFeeCalculations.SetFocus
20430         Case Else
20440           .cmdSweep.SetFocus
20450         End Select
20460       End With
20470     Case vbKeyRight
20480       With Me
20490         intRetVal = 0
20500         Select Case strMenuArrow
              Case "cmdAccrue_KeyDown"
20510           .cmdAccrue.SetFocus
20520         Case "cmdFeeCalculations_KeyDown"
20530           .cmdFeeCalculations.SetFocus
20540         Case Else
20550           .cmdSweep.SetFocus
20560         End Select
20570       End With
20580     Case vbKeyUp
20590       With Me
20600         intRetVal = 0
20610         .cmdPost.SetFocus
20620       End With
20630     Case vbKeyDown
20640       With Me
20650         intRetVal = 0
20660         .cmdTransaction.SetFocus
20670       End With
20680     End Select
20690   End If

EXITP:
20700   KeyCode = intRetVal
20710   Exit Sub

ERRH:
20720   intRetVal = 0
20730   Select Case ERR.Number
        Case Else
20740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20750   End Select
20760   Resume EXITP

End Sub

Private Sub cmdChecksPreview_Click()
' ** Check for gblnAdmin vs. others.

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChecksPreview_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngChks As Long, arr_varChk() As Variant
        Dim lngRecs As Long, lngUsers As Long
        Dim blnWithSubs As Boolean, blnWithCriteria As Boolean, blnWithVoids As Boolean
        Dim blnContinue As Boolean
        Dim lngX As Long, lngE As Long

        ' ** Array: arr_varChk().
        Const C_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const C_JID As Integer = 0
        Const C_ORD As Integer = 1

20810   With Me
20820     If blnFormOpening = False Then

20830       DoCmd.Hourglass True
20840       DoEvents

20850       .TimerInterval = 0&

20860       blnContinue = True
20870       blnFormOpening = True

20880       gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
20890       UpdateImported  ' ** Procedure: Below.
20900       DoEvents

20910       Set dbs = CurrentDb
20920       With dbs

              ' ** Append Journal, not in tblCheckPrint, just PrintCheck = True, to tblCheckPrint.
20930         Set qdf = .QueryDefs("qryPrintChecks_05_32_02")
20940         qdf.Execute
20950         Set qdf = Nothing
20960         DoEvents

              ' ** Delete tblCheckPrint, not in Journal.
20970         Set qdf = .QueryDefs("qryPrintChecks_05_32_03")
20980         qdf.Execute
20990         Set qdf = Nothing
21000         DoEvents

              ' ** Journal, with add'l fields, all journal_USER's.
21010         Set qdf = .QueryDefs("qryPrintChecks_20a")
21020         Set rst = qdf.OpenRecordset
21030         With rst
21040           If .BOF = True And .EOF = True Then
21050             DoCmd.Hourglass False
21060             blnContinue = False
21070             MsgBox "There are no entries in the posting journal.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
21080           End If
21090           .Close
21100         End With
21110         Set rst = Nothing
21120         Set qdf = Nothing
21130         DoEvents

21140         If blnContinue = True Then
                ' ** qryPrintChecks_20a (Journal, with add'l fields, all journal_USER's),
                ' ** linked to tblJournal_Memo, just 'Paid', for PrintCheck = True.
21150           Set qdf = .QueryDefs("qryPrintChecks_20b")
21160           Set rst = qdf.OpenRecordset
21170           If rst.BOF = True And rst.EOF = True Then
21180             DoCmd.Hourglass False
21190             blnContinue = False
21200             MsgBox "There are no Paid Check entries in the posting journal.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
21210             rst.Close
21220           Else
21230             rst.Close
21240           End If
21250           Set rst = Nothing
21260           Set qdf = Nothing
21270         End If
21280         DoEvents

21290         If blnContinue = True Then
21300           lngUsers = 1&
21310           Select Case gblnAdmin
                Case True
                  ' ** qryPrintChecks_20c (qryPrintChecks_20a (Journal, with add'l fields, all journal_USER's),
                  ' ** linked to tblJournal_Memo, including journal_USER, just 'Paid', for PrintCheck = True),
                  ' ** grouped by journal_USER, with cnt.
21320             Set qdf = .QueryDefs("qryPrintChecks_20d")
21330             Set rst = qdf.OpenRecordset
21340             With rst
21350               If .BOF = True And .EOF = True Then
21360                 blnContinue = False
21370                 lngUsers = 0&
21380               Else
21390                 .MoveLast
21400                 lngUsers = .RecordCount
21410               End If
21420               .Close
21430             End With
21440             Set rst = Nothing
21450             Set qdf = Nothing
21460             If blnContinue = True Then
                    ' ** qryPrintChecks_20a (Journal, with add'l fields, all journal_USER's), linked
                    ' ** to tblJournal_Memo, including journal_USER, just 'Paid', for PrintCheck = True.
21470               Set qdf = .QueryDefs("qryPrintChecks_20c")
21480             End If
21490           Case False
                  ' ** qryPrintChecks_20c (qryPrintChecks_20a (Journal, with add'l fields, all journal_USER's),
                  ' ** linked to tblJournal_Memo, grouped and summed, including journal_USER, just 'Paid',
                  ' ** for PrintCheck = True), by specified [usr].
21500             Set qdf = .QueryDefs("qryPrintChecks_20e")
21510             With qdf.Parameters
21520               ![usr] = gstrJournalUser
21530             End With
21540           End Select
21550         End If  ' ** blnContinue.
21560         DoEvents

21570         If blnContinue = True Then
21580           Set rst = qdf.OpenRecordset
21590           With rst
21600             If .BOF = True And .EOF = True Then
21610               DoCmd.Hourglass False
21620               blnContinue = False
21630               Select Case gblnAdmin
                    Case True
                      ' ** Report also has 'No Data' MsgBox that cancels opening.
21640                 MsgBox "There are no Paid Check entries in the posting journal.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
21650               Case False
                      ' ** Report also has 'No Data' MsgBox that cancels opening.
21660                 MsgBox "There are no Paid Check entries in the posting journal" & vbCrLf & _
                        "for User: " & gstrJournalUser & ".", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
21670               End Select
21680             Else
21690               blnWithSubs = False: blnWithCriteria = False: blnWithVoids = False

21700               Select Case gblnAdmin
                    Case True
21710                 If lngUsers > 1& Then
                        ' ** Reports also have 'No Data' MsgBox that cancels opening.
21720                   DoCmd.Hourglass False
21730                   If MsgBox("Do you wish to see the list broken down by Journal User?", _
                            vbQuestion + vbYesNo + vbDefaultButton2, "Divide List By User") = vbYes Then
21740                     blnWithSubs = True
21750                     strDocName = "rptCheckList_Users"  ' ** With journal_USER subdivisions; for all users.
21760                   Else
21770                     strDocName = "rptCheckList_All"  ' ** Without journal_USER subdivisions; for all or single user.
21780                   End If
21790                 Else
21800                   strDocName = "rptCheckList_All"  ' ** Without journal_USER subdivisions; for all or single user.
21810                 End If
21820               Case False
21830                 blnWithCriteria = True
21840                 strDocName = "rptCheckList_All"  ' ** Without journal_USER subdivisions; for all or single user.
21850               End Select
21860             End If
21870             .Close
21880           End With  ' ** rst.
21890           Set rst = Nothing
21900           Set qdf = Nothing
21910         End If  ' ** blnContinue.
21920         DoEvents

              ' ** Check decimal places.
21930         If blnContinue = True Then
21940           DoCmd.Hourglass True
21950           DoEvents
                ' ** Journal, just discrepancies in Paid's, with Ax, icash_new, pcash_new.
21960           Set qdf = .QueryDefs("qryPrintChecks_25")
21970           Set rst = qdf.OpenRecordset
21980           If rst.BOF = True And rst.EOF = True Then
                  ' ** All's well.
21990             rst.Close
22000           Else
22010             rst.Close
                  ' ** Update qryPrintChecks_26 (Journal, with DLookups() to qryPrintChecks_25
                  ' ** (Journal, just discrepancies in Paid's, with Ax, icash_new, pcash_new)).
22020             Set qdf = dbs.QueryDefs("qryPrintChecks_27")
22030             qdf.Execute
22040           End If
22050           Set rst = Nothing
22060           Set qdf = Nothing
22070         End If
22080         DoEvents

              ' ** The report itself calls GetDollarString() from modUtilities.

22090         If blnContinue = True Then

22100           DoCmd.Hourglass True
22110           DoEvents

22120           Select Case blnWithSubs
                Case True  ' ** With user sorting.
                  ' ** qryPrintChecks_20a (Journal, with add'l fields, all journal_USER's), linked
                  ' ** to tblJournal_Memo, including journal_USER, just 'Paid', for PrintCheck = True.
                  'journal_USER, accountno, transdate, sortOrder, ID
22130             Set qdf = .QueryDefs("qryPrintChecks_20c")
22140           Case False  ' ** Without user sorting, either 1 user, or all users.
22150             Select Case blnWithCriteria
                  Case True
                    ' ** qryPrintChecks_20c (qryPrintChecks_20a (Journal, with add'l fields, all journal_USER's),
                    ' ** linked to tblJournal_Memo, grouped and summed, including journal_USER, just 'Paid',
                    ' ** for PrintCheck = True), by specified [usr].
                    'journal_USER, accountno, transdate, sortOrder, ID
22160               Set qdf = .QueryDefs("qryPrintChecks_20e")
22170               With qdf.Parameters
22180                 ![usr] = gstrJournalUser
22190               End With
22200             Case False
                    ' ** qryPrintChecks_20a (Journal, with add'l fields, all journal_USER's),
                    ' ** linked to tblJournal_Memo, just 'Paid', for PrintCheck = True.
                    'accountno, transdate, sortOrder, ID
22210               Set qdf = .QueryDefs("qryPrintChecks_20b")
22220             End Select
22230           End Select
22240           Set rst = qdf.OpenRecordset
22250           With rst
22260             lngChks = 0&
22270             ReDim arr_varChk(C_ELEMS, 0)
22280             .MoveLast
22290             lngRecs = .RecordCount
22300             .MoveFirst
                  'BECAUSE ID'S WILL BE DUPLICATED WHEN VOIDS ARE INCLUDED, THIS WILL HAVE TO DONE DIFFERENTLY!
22310             For lngX = 1& To lngRecs
22320               lngChks = lngChks + 1&
22330               lngE = lngChks - 1&
22340               ReDim Preserve arr_varChk(C_ELEMS, lngE)
22350               arr_varChk(C_JID, lngE) = ![ID]
22360               arr_varChk(C_ORD, lngE) = lngX
22370               If lngX < lngRecs Then .MoveNext
22380             Next
22390             .Close
22400           End With
22410           Set rst = Nothing
22420           Set qdf = Nothing
22430           DoEvents

22440         End If  ' ** blnContinue.

22450         If blnContinue = True Then

                ' ** Empty tblMark.
22460           Set qdf = .QueryDefs("qryPrintChecks_14a")
22470           qdf.Execute
22480           DoEvents

                ' ** Number the checks sequentially for the report.
22490           Set rst = .OpenRecordset("tblMark", dbOpenDynaset, dbAppendOnly)
22500           With rst
22510             For lngX = 0& To (lngChks - 1&)
22520               .AddNew
22530               ![unique_id] = arr_varChk(C_JID, lngX)
22540               ![mark] = CBool(True)
22550               ![Value] = CStr(arr_varChk(C_ORD, lngX))
22560               .Update
22570             Next
22580             .Close
22590           End With
22600           Set rst = Nothing
22610           Set qdf = Nothing
22620           DoEvents

                ' *********************
                ' ** Check for voids.
                ' *********************

22630           Select Case blnWithSubs
                Case True
                  ' ** tblCheckVoid, linked to Journal.
22640             Set qdf = .QueryDefs("qryPrintChecks_20_01")
22650           Case False
22660             Select Case blnWithCriteria
                  Case True
                    ' ** tblCheckVoid, linked to Journal, by specified [usr].
22670               Set qdf = .QueryDefs("qryPrintChecks_20_01_01")
22680               With qdf.Parameters
22690                 ![usr] = gstrJournalUser
22700               End With
22710             Case False
                    ' ** tblCheckVoid, linked to Journal.
22720               Set qdf = .QueryDefs("qryPrintChecks_20_01")
22730             End Select
22740           End Select
22750           Set rst = qdf.OpenRecordset
22760           With rst
22770             If .BOF = True And .EOF = True Then
                    ' ** No voids.
22780             Else
22790               .MoveLast
22800               If .RecordCount > 0 Then
22810                 blnWithVoids = True
22820               End If
22830             End If
22840             .Close
22850           End With
22860           Set rst = Nothing
22870           Set qdf = Nothing
22880           DoEvents

22890           Select Case blnWithVoids
                Case True

                  ' ** Empty tblCheckPrint_Report.
22900             Set qdf = .QueryDefs("qryPrintChecks_20_06")
22910             qdf.Execute
22920             Set qdf = Nothing
22930             DoEvents

22940             Select Case blnWithSubs
                  Case True
                    ' ** Append qryPrintChecks_20_04 (qryPrintChecks_20_02 (Union of qryPrintChecks_20
                    ' ** (qryPrintChecks_20c (qryPrintChecks_20a (Journal, linked to tblCheckPrint,
                    ' ** with add'l fields, all journal_USER's), linked to tblJournal_Memo, including
                    ' ** journal_USER, just 'Paid', for PrintCheck = True), linked to tblMark, with
                    ' ** Memo_Print, Rpt_Ord), qryPrintChecks_20_01 (tblCheckVoid, linked to Journal)),
                    ' ** sorted by journal_USER, accountno, transdate, sortOrder, ID) to tblCheckPrint_Report.
22950               Set qdf = .QueryDefs("qryPrintChecks_20_08")
22960             Case False
22970               Select Case blnWithCriteria
                    Case True
                      ' ** Append qryPrintChecks_20_05 (qryPrintChecks_20_02 (Union of qryPrintChecks_20
                      ' ** (qryPrintChecks_20c (qryPrintChecks_20a (Journal, linked to tblCheckPrint,
                      ' ** with add'l fields, all journal_USER's), linked to tblJournal_Memo, including
                      ' ** journal_USER, just 'Paid', for PrintCheck = True), linked to tblMark, with
                      ' ** Memo_Print, Rpt_Ord), qryPrintChecks_20_01 (tblCheckVoid, linked to Journal)),
                      ' ** sorted by journal_USER, accountno, transdate, sortOrder, ID, by specified [usr]) to tblCheckPrint_Report.
22980                 Set qdf = .QueryDefs("qryPrintChecks_20_09")
22990                 With qdf.Parameters
23000                   ![usr] = gstrJournalUser
23010                 End With
23020               Case False
                      ' ** Append qryPrintChecks_20_03 (qryPrintChecks_20_02 (Union of qryPrintChecks_20
                      ' ** (qryPrintChecks_20c (qryPrintChecks_20a (Journal, linked to tblCheckPrint,
                      ' ** with add'l fields, all journal_USER's), linked to tblJournal_Memo, including
                      ' ** journal_USER, just 'Paid', for PrintCheck = True), linked to tblMark, with
                      ' ** Memo_Print, Rpt_Ord), qryPrintChecks_20_01 (tblCheckVoid, linked to Journal)),
                      ' ** sorted by accountno, transdate, sortOrder, ID) to tblCheckPrint_Report.
23030                 Set qdf = .QueryDefs("qryPrintChecks_20_07")
23040               End Select
23050             End Select
23060             qdf.Execute
23070             Set qdf = Nothing
23080             DoEvents

                  ' ** Report uses: qryPrintChecks_20_11.
23090           Case False
                  ' ** Report uses: qryPrintChecks_20.
23100           End Select

23110         End If  ' ** blnContinue.

23120         .Close
23130       End With  ' ** dbs.
23140       Set dbs = Nothing

23150       If blnContinue = True Then
23160         .WithVoids = blnWithVoids
23170         gstrReportCallingForm = THIS_NAME  ' ** When not vbNullString, sets this form .Visible = False.
23180         DoCmd.Hourglass False
23190         Select Case blnWithCriteria
              Case True
23200           DoCmd.OpenReport strDocName, acViewPreview, , "[journal_USER] = '" & gstrJournalUser & "'"
23210         Case False
                '##GTR_Ref: rptCheckList_All
                '##GTR_Ref: rptCheckList_Users
23220           DoCmd.OpenReport strDocName, acViewPreview
23230         End Select
23240         DoCmd.Maximize
23250         DoCmd.RunCommand acCmdFitToWindow
23260       End If

23270       blnFormOpening = False

23280       DoCmd.Hourglass False

23290     End If
23300   End With

EXITP:
23310   Set rst = Nothing
23320   Set qdf = Nothing
23330   Set dbs = Nothing
23340   Exit Sub

ERRH:
23350   blnFormOpening = False
23360   DoCmd.Hourglass False
23370   DoCmd.Restore
23380   If Me.Visible = False Then
23390     Me.Visible = True
23400   End If
23410   Select Case ERR.Number
        Case Else
23420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23430   End Select
23440   Resume EXITP

End Sub

Private Sub cmdChecksPreview_KeyDown(KeyCode As Integer, Shift As Integer)

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChecksPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23510   intRetVal = KeyCode

23520   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
23530   intShiftDown = (Shift And acShiftMask) > 0
23540   intAltDown = (Shift And acAltMask) > 0
23550   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23560   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23570     Select Case intRetVal
          Case vbKeyLeft
23580       With Me
23590         intRetVal = 0
23600         strMenuArrow = THIS_PROC
23610         .cmdTransaction.SetFocus
23620       End With
23630     Case vbKeyRight
23640       With Me
23650         intRetVal = 0
23660         strMenuArrow = THIS_PROC
23670         .cmdTransaction.SetFocus
23680       End With
23690     Case vbKeyUp
23700       With Me
23710         intRetVal = 0
23720         .cmdFeeCalculations.SetFocus
23730       End With
23740     Case vbKeyDown
23750       With Me
23760         intRetVal = 0
23770         .cmdChecksPrint.SetFocus
23780       End With
23790     End Select
23800   End If

EXITP:
23810   KeyCode = intRetVal
23820   Exit Sub

ERRH:
23830   intRetVal = 0
23840   Select Case ERR.Number
        Case Else
23850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23860   End Select
23870   Resume EXITP

End Sub

Private Sub cmdChecksPreview_LostFocus()

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChecksPreview_LostFocus"

23910   With Me
23920     If (.GoToReport_arw_menu_r_img.Visible = True) Then
23930       If (.GoToReport_arw_menu_r_img.Top = .cmdChecksPreview.Top + (3& * lngTpp)) Then
23940         .GoToReport_arw_menu_r_img.Visible = False
23950       End If
23960     End If
23970   End With

EXITP:
23980   Exit Sub

ERRH:
23990   Select Case ERR.Number
        Case Else
24000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24010   End Select
24020   Resume EXITP

End Sub

Private Sub cmdChecksPrint_Click()
' ** Check for gblnAdmin vs. others.

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChecksPrint_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strOpenArgs As String

24110   With Me
24120     If blnFormOpening = False Then

24130       DoCmd.Hourglass True
24140       DoEvents

24150       .TimerInterval = 0&

24160       blnFormOpening = True

24170       UpdateImported  ' ** Procedure: Below.

24180       Set dbs = CurrentDb
24190       With dbs
              ' ** Empty tblJournal_Check_Staging1.
24200         Set qdf = .QueryDefs("qryPrintChecks_16a")
24210         qdf.Execute
24220         Set qdf = Nothing
24230         DoEvents
              ' ** Empty tblJournal_Check_Staging2.
24240         Set qdf = .QueryDefs("qryPrintChecks_16b")
24250         qdf.Execute
24260         Set qdf = Nothing
24270         DoEvents
              ' ** Append qryPrintChecks_17a (qryPrintChecks_01c (Journal, just 'Misc.', 'Paid', 'Received',
              ' ** for PrintCheck = True, with Payee), with add'l fields) to tblJournal_Check_Staging1.
24280         Set qdf = .QueryDefs("qryPrintChecks_18a")
24290         qdf.Execute
24300         Set qdf = Nothing
24310         DoEvents
              ' ** Append qryPrintChecks_17c (qryPrintChecks_17b (qryPrintChecks_17a (qryPrintChecks_01c
              ' ** (Journal, just 'Misc.', 'Paid', 'Received', for PrintCheck = True, with Payee), with
              ' ** add'l fields), grouped and summed, by Account Number, with Payeex), with Payee)
              ' ** to tblJournal_Check_Staging2.
24320         Set qdf = .QueryDefs("qryPrintChecks_18b")
24330         qdf.Execute
24340         Set qdf = Nothing
24350         DoEvents
24360         .Close
24370       End With  ' ** dbs.
24380       blnChecksCalled = True

24390       gblnSetFocus = True
24400       blnJustClose = True
24410       gstrReturningForm = vbNullString
24420       gstrReportCallingForm = vbNullString
24430       strDocName = "frmRpt_Checks"
24440       strOpenArgs = THIS_NAME & "~" & Format(gblnAdmin, "True/False")
24450       DoCmd.OpenForm strDocName, , , , , , strOpenArgs
24460       If gblnGoToReport = True Then
24470         Forms(strDocName).TimerInterval = 50&
24480         blnGoingToReport = False
24490         blnGTR_Emblem = False
24500       End If
24510       DoEvents
24520       DoCmd.Close acForm, THIS_NAME

24530       DoCmd.Hourglass False

24540     End If
24550   End With

EXITP:
24560   Set qdf = Nothing
24570   Set dbs = Nothing
24580   Exit Sub

ERRH:
24590   blnFormOpening = False
24600   DoCmd.Hourglass False
24610   Select Case ERR.Number
        Case Else
24620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24630   End Select
24640   Resume EXITP

End Sub

Private Sub cmdChecksPrint_KeyDown(KeyCode As Integer, Shift As Integer)

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChecksPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24710   intRetVal = KeyCode

24720   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
24730   intShiftDown = (Shift And acShiftMask) > 0
24740   intAltDown = (Shift And acAltMask) > 0
24750   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24760   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24770     Select Case intRetVal
          Case vbKeyLeft
24780       With Me
24790         intRetVal = 0
24800         strMenuArrow = THIS_PROC
24810         .cmdTransaction.SetFocus
24820       End With
24830     Case vbKeyRight
24840       With Me
24850         intRetVal = 0
24860         strMenuArrow = THIS_PROC
24870         .cmdTransaction.SetFocus
24880       End With
24890     Case vbKeyUp
24900       With Me
24910         intRetVal = 0
24920         .cmdChecksPreview.SetFocus
24930       End With
24940     Case vbKeyDown
24950       With Me
24960         intRetVal = 0
24970         .cmdChecksClear.SetFocus
24980       End With
24990     End Select
25000   End If

EXITP:
25010   KeyCode = intRetVal
25020   Exit Sub

ERRH:
25030   intRetVal = 0
25040   Select Case ERR.Number
        Case Else
25050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25060   End Select
25070   Resume EXITP

End Sub

Private Sub cmdChecksClear_Click()
' ** Check for gblnAdmin vs. others.

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChecksClear_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCnt As Long
        Dim lngUsrs As Long, arr_varUsr As Variant
        Dim lngX As Long

        ' ** Array: arr_varUsr().
        Const U_USR As Integer = 0
        Const U_CNT As Integer = 1

25110   With Me
25120     If blnFormOpening = False Then

25130       DoCmd.Hourglass True
25140       DoEvents

25150       .TimerInterval = 0&

25160       blnFormOpening = True

25170       gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

25180       Set dbs = CurrentDb
25190       With dbs
25200         Select Case gblnAdmin
              Case True
                ' ** Journal, grouped by journal_USER, for PrintCheck = True, with cnt.
25210           Set qdf = .QueryDefs("qryPrintChecks_11c")
25220           Set rst = qdf.OpenRecordset
25230           With rst
25240             If .BOF = True And .EOF = True Then
25250               lngCnt = 0&
25260             Else
25270               .MoveLast
25280               lngUsrs = .RecordCount
25290               .MoveFirst
25300               arr_varUsr = .GetRows(lngUsrs)
                    ' *************************************************
                    ' ** Array: arr_varUsr()
                    ' **
                    ' **   Field  Element  Name            Constant
                    ' **   =====  =======  ==============  ==========
                    ' **     1       0     journal_USER    U_USR
                    ' **     2       1     cnt             U_CNT
                    ' **
                    ' *************************************************
25310               lngCnt = 0&
25320               For lngX = 0& To (lngUsrs - 1&)
25330                 lngCnt = lngCnt + arr_varUsr(U_CNT, lngX)
25340               Next
25350             End If
25360             .Close
25370           End With
25380         Case False
                ' ** Journal, grouped, PrintCheck = True, with cnt, by specified [usr].
25390           Set qdf = .QueryDefs("qryPrintChecks_11b")
25400           With qdf.Parameters
25410             ![usr] = gstrJournalUser
25420           End With
25430           Set rst = qdf.OpenRecordset
25440           With rst
25450             If .BOF = True And .EOF Then
25460               lngCnt = 0&
25470             Else
25480               .MoveFirst
25490               lngCnt = ![cnt]
25500             End If
25510             .Close
25520           End With
25530         End Select
25540         .Close
25550       End With

25560       With Me
25570         .CheckCnt = lngCnt
25580         DoEvents
25590       End With

25600       If lngCnt > 0& Then
25610         Select Case gblnAdmin
              Case True
25620           If lngUsrs = 1& Then
25630             DoCmd.Hourglass False
25640             strMsg = "Are you sure you want to clear these checks?" & vbCrLf & vbCrLf & _
                    "Checks for User " & arr_varUsr(U_USR, 0) & ": " & CStr(lngCnt)
25650             msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Clear Checks")
25660           Else
25670             gstrJournalUser = vbNullString  ' ** Don't clear this on the form!
25680             strDocName = "frmMenu_Post_Clear_Multi"
25690             DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME & "~" & "ClearChecks"
25700             If gstrJournalUser = vbNullString Then
25710               msgResponse = vbNo
25720             Else
25730               msgResponse = vbYes
25740             End If
25750           End If
25760         Case False
25770           DoCmd.Hourglass False
25780           strMsg = "Are you sure you want to clear these checks?" & vbCrLf & vbCrLf & _
                  "Checks for User " & gstrJournalUser & ": " & CStr(lngCnt)
25790           msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Clear Checks")
25800         End Select

25810         If msgResponse = vbYes Then
25820           DoCmd.Hourglass True
25830           DoEvents
25840           Set dbs = CurrentDb
25850           With dbs
25860             Select Case gstrJournalUser
                  Case "All"
                    ' ** Update Journal, for PrintCheck = False, CheckNum = Null.
25870               Set qdf = .QueryDefs("qryPrintChecks_12a")
25880               qdf.Execute
                    ' ** Empty tblJournal_Memo.
25890               Set qdf = .QueryDefs("qryPrintChecks_13a")
25900               qdf.Execute
                    ' ** Empty tblMark.
25910               Set qdf = .QueryDefs("qryPrintChecks_14a")
25920               qdf.Execute
25930             Case Else
                    ' ** Update Journal, for PrintCheck = False, CheckNum = Null, by specified [usr].
25940               Set qdf = .QueryDefs("qryPrintChecks_12b")
25950               With qdf.Parameters
25960                 ![usr] = gstrJournalUser
25970               End With
25980               qdf.Execute
                    ' ** Delete tblJournal_Memo, by DLookups() to qryPrintChecks_13c (tblJournal_Memo, linked to Journal, by specified [usr]).
25990               Set qdf = .QueryDefs("qryPrintChecks_13b")
26000               With qdf.Parameters
26010                 ![usr] = gstrJournalUser
26020               End With
26030               qdf.Execute
                    ' ** Delete tblMark, by DLookups() to qryPrintChecks_14c (tblMark, linked to Journal, by specified [usr]).
26040               Set qdf = .QueryDefs("qryPrintChecks_14b")
26050               With qdf.Parameters
26060                 ![usr] = gstrJournalUser
26070               End With
26080               qdf.Execute
26090             End Select
26100             .Close
26110           End With
26120           DoCmd.Hourglass False
26130           Select Case gstrJournalUser
                Case "All"
26140             MsgBox "All checks have been cleared.", vbInformation + vbOKOnly, ("Checks Cleared" & Space(40))
26150           Case Else
26160             MsgBox "Checks have been cleared for User:" & gstrJournalUser & ".", vbInformation + vbOKOnly, ("Checks Cleared" & Space(40))
26170           End Select
26180         Else
26190           DoCmd.Hourglass False
26200         End If

26210       Else
26220         DoCmd.Hourglass False
26230         MsgBox "There are no checks to Clear in the Posting Journal.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
26240       End If

26250       blnFormOpening = False

26260     End If
26270   End With

EXITP:
26280   Set rst = Nothing
26290   Set qdf = Nothing
26300   Set dbs = Nothing
26310   Exit Sub

ERRH:
26320   blnFormOpening = False
26330   DoCmd.Hourglass False
26340   Select Case ERR.Number
        Case Else
26350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26360   End Select
26370   Resume EXITP

End Sub

Private Sub cmdChecksClear_KeyDown(KeyCode As Integer, Shift As Integer)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChecksClear_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26410   intRetVal = KeyCode

26420   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
26430   intShiftDown = (Shift And acShiftMask) > 0
26440   intAltDown = (Shift And acAltMask) > 0
26450   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26460   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26470     Select Case intRetVal
          Case vbKeyLeft
26480       With Me
26490         intRetVal = 0
26500         .cmdPostingJournalPreview.SetFocus
26510       End With
26520     Case vbKeyRight
26530       With Me
26540         intRetVal = 0
26550         .cmdPostingJournalPreview.SetFocus
26560       End With
26570     Case vbKeyUp
26580       With Me
26590         intRetVal = 0
26600         .cmdChecksPrint.SetFocus
26610       End With
26620     Case vbKeyDown
26630       With Me
26640         intRetVal = 0
26650         .cmdPOSPay.SetFocus
26660       End With
26670     End Select
26680   End If

EXITP:
26690   KeyCode = intRetVal
26700   Exit Sub

ERRH:
26710   intRetVal = 0
26720   Select Case ERR.Number
        Case Else
26730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26740   End Select
26750   Resume EXITP

End Sub

Private Sub cmdPOSPay_Click()

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPOSPay_Click"

26810   With Me
26820     If blnFormOpening = False Then

26830       DoCmd.Hourglass True
26840       DoEvents

26850       .TimerInterval = 0&

26860       gblnSetFocus = True
26870       blnJustClose = True: blnFormOpening = True
26880       gstrReturningForm = vbNullString
26890       strDocName = "frmCheckPOSPay"
26900       DoCmd.Hourglass False
26910       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
26920       DoCmd.Close acForm, THIS_NAME

26930       DoCmd.Hourglass False

26940     End If
26950   End With

EXITP:
26960   Exit Sub

ERRH:
26970   DoCmd.Hourglass False
26980   Select Case ERR.Number
        Case Else
26990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27000   End Select
27010   Resume EXITP

End Sub

Private Sub cmdPOSPay_KeyDown(KeyCode As Integer, Shift As Integer)

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPOSPay_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27110   intRetVal = KeyCode

27120   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
27130   intShiftDown = (Shift And acShiftMask) > 0
27140   intAltDown = (Shift And acAltMask) > 0
27150   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27160   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27170     Select Case intRetVal
          Case vbKeyLeft
27180       With Me
27190         intRetVal = 0
27200         .cmdPostingJournalPrint.SetFocus
27210       End With
27220     Case vbKeyRight
27230       With Me
27240         intRetVal = 0
27250         .cmdPostingJournalPrint.SetFocus
27260       End With
27270     Case vbKeyUp
27280       With Me
27290         intRetVal = 0
27300         .cmdChecksClear.SetFocus
27310       End With
27320     Case vbKeyDown
27330       With Me
27340         intRetVal = 0
27350         .cmdCheckMemos.SetFocus
27360       End With
27370     End Select
27380   End If

EXITP:
27390   KeyCode = intRetVal
27400   Exit Sub

ERRH:
27410   intRetVal = 0
27420   Select Case ERR.Number
        Case Else
27430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27440   End Select
27450   Resume EXITP

End Sub

Private Sub cmdCheckMemos_Click()

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckMemos_Click"

27510   With Me
27520     If blnFormOpening = False Then

27530       DoCmd.Hourglass True
27540       DoEvents

27550       .TimerInterval = 0&

27560       blnJustClose = True: blnFormOpening = True
27570       gstrReturningForm = vbNullString
27580       strDocName = "frmCheckMemo"
27590       DoCmd.Hourglass False
27600       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
27610       DoCmd.Close acForm, THIS_NAME

27620       DoCmd.Hourglass False

27630     End If
27640   End With

EXITP:
27650   Exit Sub

ERRH:
27660   blnFormOpening = False
27670   DoCmd.Hourglass False
27680   Select Case ERR.Number
        Case Else
27690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27700   End Select
27710   Resume EXITP

End Sub

Private Sub cmdCheckMemos_KeyDown(KeyCode As Integer, Shift As Integer)

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckMemos_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27810   intRetVal = KeyCode

27820   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
27830   intShiftDown = (Shift And acShiftMask) > 0
27840   intAltDown = (Shift And acAltMask) > 0
27850   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27860   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27870     Select Case intRetVal
          Case vbKeyLeft
27880       With Me
27890         intRetVal = 0
27900         .cmdPost.SetFocus
27910       End With
27920     Case vbKeyRight
27930       With Me
27940         intRetVal = 0
27950         .cmdPost.SetFocus
27960       End With
27970     Case vbKeyUp
27980       With Me
27990         intRetVal = 0
28000         .cmdPOSPay.SetFocus
28010       End With
28020     Case vbKeyDown
28030       With Me
28040         intRetVal = 0
28050         .cmdCheckReconcile.SetFocus
28060       End With
28070     End Select
28080   End If

EXITP:
28090   KeyCode = intRetVal
28100   Exit Sub

ERRH:
28110   intRetVal = 0
28120   Select Case ERR.Number
        Case Else
28130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28140   End Select
28150   Resume EXITP

End Sub

Private Sub cmdCheckReconcile_Click()

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckReconcile_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

28210   With Me
28220     If blnFormOpening = False Then

28230       DoCmd.Hourglass True
28240       DoEvents

28250       .TimerInterval = 0&

28260       blnFormOpening = True

28270       UpdateImported  ' ** Procedure: Below.
28280       DoEvents

28290       Set dbs = CurrentDb
28300       Set rst = dbs.OpenRecordset("tblCheckReconcile_Form_Graphics", dbOpenDynaset, dbReadOnly)
28310       If rst.BOF = True And rst.EOF = True Then
28320         rst.Close
28330         Set rst = Nothing
              ' ** Append qryCheckReconcile_19 (tblForm_Graphics, just frmCheckReconcile) to tblCheckReconcile_Form_Graphics.
28340         Set qdf = dbs.QueryDefs("qryCheckReconcile_20")
28350         qdf.Execute
28360         Set qdf = Nothing
28370         DoEvents
28380       Else
28390         rst.Close
28400         Set rst = Nothing
28410       End If
            ' ** Empty tmpAccount.
28420       Set qdf = dbs.QueryDefs("qryCheckReconcile_03_01")
28430       qdf.Execute
28440       Set qdf = Nothing
28450       DoEvents
            ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
            ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
            ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
            ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
            ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
            ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
            ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
            ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
28460       Set qdf = dbs.QueryDefs("qryCheckReconcile_03_02")
28470       qdf.Execute
28480       Set qdf = Nothing
28490       DoEvents
            ' ** Append qryCheckReconcile_03_03 (Account, just 'CRTC01', with alphasort) to tmpAccount.
28500       Set qdf = dbs.QueryDefs("qryCheckReconcile_03_04")
28510       qdf.Execute
28520       Set qdf = Nothing
28530       DoEvents
28540       dbs.Close
28550       Set dbs = Nothing

28560       gblnSetFocus = True
28570       blnJustClose = True
28580       gstrReturningForm = vbNullString
28590       strDocName = "frmCheckReconcile"
28600       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
28610       Select Case gblnGoToReport
            Case True
28620         Forms(strDocName).TimerInterval = 50&
28630         blnGoingToReport = False
28640         blnGTR_Emblem = False
28650       Case False
28660         DoCmd.Hourglass False
28670         Forms(strDocName).TimerInterval = 500&
28680       End Select
28690       DoCmd.Close acForm, THIS_NAME

28700       DoCmd.Hourglass False

28710     End If
28720   End With

EXITP:
28730   Set rst = Nothing
28740   Set qdf = Nothing
28750   Set dbs = Nothing
28760   Exit Sub

ERRH:
28770   blnFormOpening = False
28780   DoCmd.Hourglass False
28790   Select Case ERR.Number
        Case Else
28800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28810   End Select
28820   Resume EXITP

End Sub

Private Sub cmdCheckReconcile_KeyDown(KeyCode As Integer, Shift As Integer)

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckReconcile_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28910   intRetVal = KeyCode

28920   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
28930   intShiftDown = (Shift And acShiftMask) > 0
28940   intAltDown = (Shift And acAltMask) > 0
28950   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28960   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28970     Select Case intRetVal
          Case vbKeyLeft
28980       With Me
28990         intRetVal = 0
29000         .cmdClearPostingJournal.SetFocus
29010       End With
29020     Case vbKeyRight
29030       With Me
29040         intRetVal = 0
29050         .cmdClearPostingJournal.SetFocus
29060       End With
29070     Case vbKeyUp
29080       With Me
29090         intRetVal = 0
29100         .cmdCheckMemos.SetFocus
29110       End With
29120     Case vbKeyDown
29130       With Me
29140         intRetVal = 0
29150         .cmdSweep.SetFocus
29160       End With
29170     End Select
29180   End If

EXITP:
29190   KeyCode = intRetVal
29200   Exit Sub

ERRH:
29210   intRetVal = 0
29220   Select Case ERR.Number
        Case Else
29230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29240   End Select
29250   Resume EXITP

End Sub

Private Sub cmdSweep_Click()

29300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSweep_Click"

29310   With Me
29320     If blnFormOpening = False Then

29330       DoCmd.Hourglass True
29340       DoEvents

29350       .TimerInterval = 0&

29360       UpdateImported  ' ** Procedure: Below.

29370       DoCmd.Hourglass True  ' ** Make sure it's still running.
29380       DoEvents

29390       blnJustClose = True: blnFormOpening = True
29400       gstrReturningForm = vbNullString
29410       strDocName = "frmSweeper"
29420       DoCmd.Hourglass False
29430       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
29440       DoCmd.Close acForm, THIS_NAME

29450       DoCmd.Hourglass False

29460     End If
29470   End With

EXITP:
29480   Exit Sub

ERRH:
29490   blnFormOpening = False
29500   DoCmd.Hourglass False
29510   Select Case ERR.Number
        Case Else
29520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29530   End Select
29540   Resume EXITP

End Sub

Private Sub cmdSweep_KeyDown(KeyCode As Integer, Shift As Integer)

29600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSweep_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29610   intRetVal = KeyCode

29620   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
29630   intShiftDown = (Shift And acShiftMask) > 0
29640   intAltDown = (Shift And acAltMask) > 0
29650   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29660   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29670     Select Case intRetVal
          Case vbKeyLeft
29680       With Me
29690         intRetVal = 0
29700         strMenuArrow = THIS_PROC
29710         .cmdClearPostingJournal.SetFocus
29720       End With
29730     Case vbKeyRight
29740       With Me
29750         intRetVal = 0
29760         strMenuArrow = THIS_PROC
29770         .cmdClearPostingJournal.SetFocus
29780       End With
29790     Case vbKeyUp
29800       With Me
29810         intRetVal = 0
29820         .cmdCheckReconcile.SetFocus
29830       End With
29840     Case vbKeyDown
29850       With Me
29860         intRetVal = 0
29870         .cmdAccrue.SetFocus
29880       End With
29890     End Select
29900   End If

EXITP:
29910   KeyCode = intRetVal
29920   Exit Sub

ERRH:
29930   intRetVal = 0
29940   Select Case ERR.Number
        Case Else
29950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29960   End Select
29970   Resume EXITP

End Sub

Private Sub cmdAccrue_Click()

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccrue_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngJrnlCnt As Long

30010   With Me
30020     If blnFormOpening = False Then

30030       DoCmd.Hourglass True
30040       DoEvents

30050       .TimerInterval = 0&

30060       blnFormOpening = True

30070       UpdateImported  ' ** Procedure: Below.

30080       DoCmd.Hourglass True  ' ** Make sure it's still running.
30090       DoEvents

30100       lngJrnlCnt = DCount("*", "journal")

            ' ** Empty tmpEdit07.
30110       Set dbs = CurrentDb
30120       With dbs
30130         Set qdf = .QueryDefs("qryAccruedIncome_05")
30140         qdf.Execute
30150         .Close
30160       End With

30170       DoCmd.Hourglass True  ' ** Make sure it's still running.
30180       DoEvents

30190       blnJustClose = True
30200       gstrReturningForm = vbNullString
30210       strDocName = "frmAccruedIncome"
30220       DoCmd.Hourglass False
30230       DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & CStr(lngJrnlCnt)
30240       DoCmd.Close acForm, THIS_NAME

30250       DoCmd.Hourglass False

30260     End If
30270   End With

EXITP:
30280   Set qdf = Nothing
30290   Set dbs = Nothing
30300   Exit Sub

ERRH:
30310   blnFormOpening = False
30320   DoCmd.Hourglass False
30330   Select Case ERR.Number
        Case Else
30340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30350   End Select
30360   Resume EXITP

End Sub

Private Sub cmdAccrue_KeyDown(KeyCode As Integer, Shift As Integer)

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccrue_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30410   intRetVal = KeyCode

30420   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
30430   intShiftDown = (Shift And acShiftMask) > 0
30440   intAltDown = (Shift And acAltMask) > 0
30450   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30460   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30470     Select Case intRetVal
          Case vbKeyLeft
30480       With Me
30490         intRetVal = 0
30500         strMenuArrow = THIS_PROC
30510         .cmdClearPostingJournal.SetFocus
30520       End With
30530     Case vbKeyRight
30540       With Me
30550         intRetVal = 0
30560         strMenuArrow = THIS_PROC
30570         .cmdClearPostingJournal.SetFocus
30580       End With
30590     Case vbKeyUp
30600       With Me
30610         intRetVal = 0
30620         .cmdSweep.SetFocus
30630       End With
30640     Case vbKeyDown
30650       With Me
30660         intRetVal = 0
30670         .cmdFeeCalculations.SetFocus
30680       End With
30690     End Select
30700   End If

EXITP:
30710   KeyCode = intRetVal
30720   Exit Sub

ERRH:
30730   intRetVal = 0
30740   Select Case ERR.Number
        Case Else
30750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30760   End Select
30770   Resume EXITP

End Sub

Private Sub cmdFeeCalculations_Click()

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdFeeCalculations_Click"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngUsrs As Long, arr_varUsr() As Variant
        Dim strThisUser As String
        Dim blnFound As Boolean
        Dim lngX As Long, lngY As Long, lngE As Long

        ' ** Array: arr_varUsr().
        Const U_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const U_ID   As Integer = 0
        Const U_NAM  As Integer = 1
        Const U_ANAM As Integer = 2
        Const U_DAT  As Integer = 3
        Const U_ADAT As Integer = 4
        Const U_DEL  As Integer = 5

30810   With Me
30820     If blnFormOpening = False Then

30830       DoCmd.Hourglass True
30840       DoEvents

30850       .TimerInterval = 0&

30860       blnFormOpening = True

30870       UpdateImported  ' ** Procedure: Below.

30880       lngUsrs = 0&
30890       ReDim arr_varUsr(U_ELEMS, 0)

30900       DoCmd.Hourglass True  ' ** Make sure it's still running.
30910       DoEvents

30920       strThisUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
30930       Set dbs = CurrentDb
30940       With dbs
30950         Set rst = .OpenRecordset("PostingDate", dbOpenDynaset, dbConsistent)
30960         With rst
30970           If .BOF = True And .EOF = True Then
                  ' ** This shouldn't happen!
30980             .AddNew
30990             ![Posting_Date] = Date
31000             ![Username] = strThisUser
31010             .Update
31020           Else
31030             .MoveLast
31040             lngRecs = .RecordCount
31050             .MoveFirst
31060             For lngX = 1& To lngRecs
31070               lngUsrs = lngUsrs + 1&
31080               lngE = lngUsrs - 1&
31090               ReDim Preserve arr_varUsr(U_ELEMS, lngE)
31100               arr_varUsr(U_ID, lngE) = ![PostingDate_ID]
31110               arr_varUsr(U_NAM, lngE) = ![Username]
31120               arr_varUsr(U_ANAM, lngE) = CBool(False)
31130               arr_varUsr(U_DAT, lngE) = ![Posting_Date]
31140               arr_varUsr(U_ADAT, lngE) = CBool(False)
31150               arr_varUsr(U_DEL, lngE) = CBool(False)
31160               If lngX < lngRecs Then .MoveNext
31170             Next
31180             For lngX = 0& To (lngUsrs - 1&)
31190               If IsNull(arr_varUsr(U_NAM, lngX)) = True Then
                      ' ** This shouldn't be able to happen! Check if this user is present, otherwise delete it!
31200                 blnFound = False
31210                 For lngY = 0& To (lngUsrs - 1&)
31220                   If IsNull(arr_varUsr(U_NAM, lngY)) = False Then
31230                     If arr_varUsr(U_NAM, lngY) = strThisUser Then
31240                       blnFound = True
31250                       Exit For
31260                     End If
31270                   End If
31280                 Next
31290                 Select Case blnFound
                      Case True
31300                   arr_varUsr(U_DEL, lngX) = CBool(True)
31310                 Case False
31320                   arr_varUsr(U_ANAM, lngX) = CBool(True)
31330                 End Select
31340               End If
31350               If IsNull(arr_varUsr(U_DAT, lngX)) = True Then
                      ' ** This shouldn't be able to happen either!
31360                 arr_varUsr(U_ADAT, lngX) = CBool(True)
31370               End If
31380             Next
31390             For lngX = 0& To (lngUsrs - 1&)
31400               If arr_varUsr(U_ADAT, lngX) = True And arr_varUsr(U_DEL, lngX) = False Then
31410                 .FindFirst "[PostingDate_ID] = " & CStr(arr_varUsr(U_ID, lngX))
31420                 If .NoMatch = False Then
31430                   .Edit
31440                   ![Posting_Date] = Date
31450                   .Update
31460                 End If
31470               End If
31480             Next
31490             For lngX = (lngUsrs - 1&) To 0& Step -1&
31500               If arr_varUsr(U_DEL, lngX) = True Then
31510                 .FindFirst "[PostingDate_ID] = " & CStr(arr_varUsr(U_ID, lngX))
31520                 If .NoMatch = False Then
31530                   .Delete
31540                 End If
31550               End If
31560             Next
31570           End If
31580         End With
31590         .Close
31600       End With

31610       DoCmd.Hourglass True  ' ** Make sure it's still running.
31620       DoEvents

31630       gblnSetFocus = True
31640       blnJustClose = True
31650       gstrReturningForm = vbNullString
31660       strDocName = "frmFeeCalculations"
31670       DoCmd.Hourglass False
31680       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
31690       If gblnGoToReport = True Then
31700         Forms(strDocName).TimerInterval = 50&
31710         blnGoingToReport = False
31720         blnGTR_Emblem = False
31730       End If
31740       DoCmd.Close acForm, THIS_NAME

31750       DoCmd.Hourglass False

31760     End If
31770   End With

EXITP:
31780   Set rst = Nothing
31790   Set dbs = Nothing
31800   Exit Sub

ERRH:
31810   blnFormOpening = False
31820   DoCmd.Hourglass False
31830   Select Case ERR.Number
        Case Else
31840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31850   End Select
31860   Resume EXITP

End Sub

Private Sub cmdFeeCalculations_KeyDown(KeyCode As Integer, Shift As Integer)

31900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdFeeCalculations_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31910   intRetVal = KeyCode

31920   strMenuArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
31930   intShiftDown = (Shift And acShiftMask) > 0
31940   intAltDown = (Shift And acAltMask) > 0
31950   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31960   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31970     Select Case intRetVal
          Case vbKeyTab
31980       With Me
31990         intRetVal = 0
32000         If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
32010           .chkLoadTimer.SetFocus
32020         Else
32030           .cmdClose.SetFocus
32040         End If
32050       End With
32060     Case vbKeyLeft
32070       With Me
32080         intRetVal = 0
32090         strMenuArrow = THIS_PROC
32100         .cmdClearPostingJournal.SetFocus
32110       End With
32120     Case vbKeyRight
32130       With Me
32140         intRetVal = 0
32150         strMenuArrow = THIS_PROC
32160         .cmdClearPostingJournal.SetFocus
32170       End With
32180     Case vbKeyUp
32190       With Me
32200         intRetVal = 0
32210         .cmdAccrue.SetFocus
32220       End With
32230     Case vbKeyDown
32240       With Me
32250         intRetVal = 0
32260         .cmdChecksPreview.SetFocus
32270       End With
32280     End Select
32290   End If

        ' ** Shift keys.
32300   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32310     Select Case intRetVal
          Case vbKeyTab
32320       With Me
32330         intRetVal = 0
32340         .cmdAccrue.SetFocus
32350       End With
32360     End Select
32370   End If

EXITP:
32380   KeyCode = intRetVal
32390   Exit Sub

ERRH:
32400   intRetVal = 0
32410   Select Case ERR.Number
        Case Else
32420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32430   End Select
32440   Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

32510   With Me
32520     If IsLoaded("frmLoadTimer", acForm) = False Then  ' ** Module Function: modFileUtilities.
            ' ** Don't turn it off if it's still playing out a timing sequence.
32530       gblnTimer = False
32540     End If
32550     Select Case .chkLoadTimer
          Case True
32560       .chkLoadTimer_lbl.FontBold = True
32570       If .chkLoadTimer.Visible = True Then  ' ** This may get turned off by the Form_Timer() sequence.
32580         gblnTimer = True
32590       End If
32600     Case False
32610       .chkLoadTimer_lbl.FontBold = False
32620     End Select
32630   End With

EXITP:
32640   Exit Sub

ERRH:
32650   Select Case ERR.Number
        Case Else
32660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32670   End Select
32680   Resume EXITP

End Sub

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
32720   intShiftDown = (Shift And acShiftMask) > 0
32730   intAltDown = (Shift And acAltMask) > 0
32740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32760     Select Case intRetVal
          Case vbKeyTab
32770       With Me
32780         intRetVal = 0
32790         .cmdClose.SetFocus
32800       End With
32810     End Select
32820   End If

        ' ** Shift keys.
32830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32840     Select Case intRetVal
          Case vbKeyTab
32850       With Me
32860         intRetVal = 0
32870         .cmdFeeCalculations.SetFocus
32880       End With
32890     End Select
32900   End If

EXITP:
32910   KeyCode = intRetVal
32920   Exit Sub

ERRH:
32930   intRetVal = 0
32940   Select Case ERR.Number
        Case Else
32950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32960   End Select
32970   Resume EXITP

End Sub

Public Function UpdateTransCnt() As Boolean

33000 On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateTransCnt"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCnt1 As Long, lngCnt2 As Long, lngCnt3 As Long, lngCnt4 As Long
        Dim blnRetVal As Boolean

33010   blnRetVal = True

33020   gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

33030   Set dbs = CurrentDb
33040   With dbs
          ' ** Journal, grouped, with cnt, all Users.
33050     Set qdf = .QueryDefs("qryJournal_Clear_01a")
33060     Set rst = qdf.OpenRecordset
33070     With rst
33080       .MoveFirst
33090       lngCnt1 = ![cnt]
33100       .Close
33110     End With
          ' ** Journal, grouped, with cnt, by specified [usr].
33120     Set qdf = .QueryDefs("qryJournal_Clear_01b")
33130     With qdf.Parameters
33140       ![usr] = gstrJournalUser
33150     End With
33160     Set rst = qdf.OpenRecordset
33170     With rst
33180       .MoveFirst
33190       lngCnt3 = ![cnt]
33200       .Close
33210     End With
          ' ** tblJournal_Column, grouped, just uncommitted, with cnt.
33220     Set qdf = .QueryDefs("qryJournal_Clear_01g")
33230     Set rst = qdf.OpenRecordset
33240     With rst
33250       .MoveFirst
33260       lngCnt2 = ![cnt]
33270       .Close
33280     End With
          ' ** tblJournal_Column, grouped, just uncommitted, with cnt, by specified [usr].
33290     Set qdf = .QueryDefs("qryJournal_Clear_01h")
33300     With qdf.Parameters
33310       ![usr] = gstrJournalUser
33320     End With
33330     Set rst = qdf.OpenRecordset
33340     With rst
33350       .MoveFirst
33360       lngCnt4 = ![cnt]
33370       .Close
33380     End With
33390     .Close
33400   End With
33410   With Me
33420     If gblnAdmin = True Then
33430       If lngCnt1 > 0& And lngCnt2 > 0& Then
33440         .cmdClearPostingJournal_lbl.Caption = CStr(lngCnt1) & " (" & CStr(lngCnt2) & ")"
33450       ElseIf lngCnt1 > 0& Then
33460         .cmdClearPostingJournal_lbl.Caption = CStr(lngCnt1)
33470       ElseIf lngCnt2 > 0& Then
33480         .cmdClearPostingJournal_lbl.Caption = "(" & CStr(lngCnt2) & ")"
33490       Else
33500         .cmdClearPostingJournal_lbl.Caption = vbNullString
33510       End If
33520     Else
33530       If lngCnt3 > 0& And lngCnt4 > 0& Then
33540         .cmdClearPostingJournal_lbl.Caption = CStr(lngCnt3) & " (" & CStr(lngCnt4) & ")"
33550       ElseIf lngCnt3 > 0& Then
33560         .cmdClearPostingJournal_lbl.Caption = CStr(lngCnt3)
33570       ElseIf lngCnt4 > 0& Then
33580         .cmdClearPostingJournal_lbl.Caption = "(" & CStr(lngCnt4) & ")"
33590       Else
33600         .cmdClearPostingJournal_lbl.Caption = vbNullString
33610       End If
33620     End If
33630     .TransCnt = lngCnt1
33640     .UncommitCnt = lngCnt2
33650     .TransCnt_User = lngCnt3
33660     .UncommitCnt_User = lngCnt4
33670     .cmdPost_msg_lbl.Visible = False
33680   End With

EXITP:
33690   Set rst = Nothing
33700   Set qdf = Nothing
33710   Set dbs = Nothing
33720   UpdateTransCnt = blnRetVal
33730   Exit Function

ERRH:
33740   DoCmd.Hourglass False
33750   blnRetVal = False
33760   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Ignore.
33770   Case Else
33780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33790   End Select
33800   Resume EXITP

End Function

Private Sub UpdateImported()

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateImported"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnHasImport As Boolean

33910   blnHasImport = False

33920   Set dbs = CurrentDb
33930   With dbs

          ' ** Trust Import; Journal, with non-Null assettype, CheckNum, for date between 2010 and 2030 and IsTI = True.
33940     Set qdf = .QueryDefs("qryJournal_13")
33950     Set rst = qdf.OpenRecordset
33960     With rst
33970       If .BOF = True And .EOF = True Then
              ' ** No imported records.
33980       Else
33990         blnHasImport = True
34000       End If
34010       .Close
34020     End With

34030     If blnHasImport = True Then
            ' ** Trust Import; Append qryJournal_13 (Journal, with non-Null assettype, CheckNum, for date
            ' ** between 2010 and 2030 and IsTI = True) to tblJournal_Import. (Not already in tblJournal_Import.)
34040       Set qdf = .QueryDefs("qryJournal_14")
34050       qdf.Execute
            ' ** Trust Import; Update qryJournal_15 (Journal, linked to tblJournal_Import),
            ' ** for CheckNum = Null.  (Only where it matches tblJournal_Import.)
34060       Set qdf = .QueryDefs("qryJournal_16")
34070       qdf.Execute
34080     End If

34090     .Close
34100   End With

EXITP:
34110   Set rst = Nothing
34120   Set qdf = Nothing
34130   Set dbs = Nothing
34140   Exit Sub

ERRH:
34150   DoCmd.Hourglass False
34160   Select Case ERR.Number
        Case Else
34170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34180   End Select
34190   Resume EXITP

End Sub

Private Sub Journal_OptLoad()

34200 On Error GoTo ERRH

        Const THIS_PROC As String = "Journal_OptLoad"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long, lngE As Long

34210   If lngOpts = 0& Or IsEmpty(arr_varOpt) Then

34220     lngOpts = 0&
34230     ReDim arr_varOpt(OPT_ELEMS, 0)

34240     Set dbs = CurrentDb
34250     With dbs
            ' ** tblForm_Control, just frmJournal, with ctlspec_optionvalue.
34260       Set qdf = .QueryDefs("qryReport_03")
34270       Set rst = qdf.OpenRecordset
34280       With rst
34290         If .BOF = True And .EOF = True Then
                ' ** Oops! Some tables didn't get updated!
34300           lngRecs = 0&
34310         Else
34320           .MoveLast
34330           lngRecs = .RecordCount
34340           .MoveFirst
34350           For lngX = 1& To lngRecs
34360             lngOpts = lngOpts + 1&
34370             lngE = lngOpts - 1&
34380             ReDim Preserve arr_varOpt(OPT_ELEMS, lngE)
34390             arr_varOpt(OPT_NAM, lngE) = ![ctl_name]
34400             arr_varOpt(OPT_VAL, lngE) = ![ctlspec_optionvalue]
34410             If lngX < lngRecs Then .MoveNext
34420           Next
34430         End If
34440         .Close
34450       End With
34460       .Close
34470     End With

34480     For lngX = 0& To (lngOpts - 1&)
34490       If InStr(arr_varOpt(OPT_NAM, lngX), "Dividend") > 0 Then
34500         JRN_DIV = arr_varOpt(OPT_VAL, lngX)
34510       ElseIf InStr(arr_varOpt(OPT_NAM, lngX), "Interest") > 0 Then
34520         JRN_INT = arr_varOpt(OPT_VAL, lngX)
34530       ElseIf InStr(arr_varOpt(OPT_NAM, lngX), "Purchase") > 0 Then
34540         JRN_PUR = arr_varOpt(OPT_VAL, lngX)
34550       ElseIf InStr(arr_varOpt(OPT_NAM, lngX), "Sale") > 0 Or InStr(arr_varOpt(OPT_NAM, lngX), "Sold") > 0 Then
34560         JRN_SAL = arr_varOpt(OPT_VAL, lngX)
34570       ElseIf InStr(arr_varOpt(OPT_NAM, lngX), "Misc") > 0 Then
34580         JRN_MIS = arr_varOpt(OPT_VAL, lngX)
34590       End If
34600     Next

34610   End If

EXITP:
34620   Set rst = Nothing
34630   Set qdf = Nothing
34640   Set dbs = Nothing
34650   Exit Sub

ERRH:
34660   DoCmd.Hourglass False
34670   Select Case ERR.Number
        Case Else
34680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34690   End Select
34700   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

34810   With Me
34820     blnGTR_Emblem = False
34830     For lngX = 1& To 24&
34840 On Error Resume Next
34850       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
34860       If ERR.Number <> 0 Then
34870 On Error GoTo ERRH
34880         Exit For
34890       Else
34900 On Error GoTo ERRH
34910       End If
34920     Next
34930   End With

EXITP:
34940   Exit Sub

ERRH:
34950   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Ignore.
34960   Case Else
34970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34980   End Select
34990   Resume EXITP

End Sub
