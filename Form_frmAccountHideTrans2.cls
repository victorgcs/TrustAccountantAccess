VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans2"

'VGC 09/23/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Show All:         A {chkShowAll}
' **   Close:            C {cmdClose}
' **   ledger_HIDDEN:    D {ledger_HIDDEN on frmAccountHideTrans2_Sub_List}
' **   Hidden Groups:    G {cmdHidden}
' **   Cancel:           E {cmdCancel}
' **   Include Archive:  H {chkIncludeArchive}
' **   UnHide:           I {cmdUnHide}
' **   Journal No:       J {chkShowJournalNo}
' **   ledger_HIDDEN:    N {ledger_HIDDEN on frmAccountHideTrans2_Sub_Pick}
' **   OK:               O {cmdOK}
' **   Print Report:     P {cmdPrintReport}
' **   Show Hidden Only: S {chkShowHiddenOnly}
' **   Include Currency: U {chkIncludeCurrency}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAccountHideTrans2_Sub_List, frmAccountHideTrans2_Sub_Pick}

'OHANA: 01:07:08

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private lngSubRow_Height As Long, lngNavBtns_Height As Long, lngSubHeader_Height As Long, lngSubFooter_Height As Long
Private lngSubList_Height As Long, lngSubPick_Height As Long, lngSubList_Rows As Long, lngSubPick_Rows As Long
Private lngShortcutLbl_Offset As Long

Private lngICashSub_Left As Long, lngPCashSub_Left As Long, lngCostSub_Left As Long, lngSubTotLbl_Offset As Long
Private lngOK_Left As Long, lngCancel_Left As Long, lngPTrans_Left As Long, lngPTransLbl_Offset As Long
Private lngSharefaceSub_Left As Long, lngCurrID_Left As Long, lngPrintReportLbl_Left As Long

Private blnHasForEx_Chk As Boolean ', blnHasForExThis As Boolean
Private lngHidType As Long, lngCurrID_Width As Long
Private lngHids As Long, blnHasCurrent As Boolean, blnHasCurrentMsgShown As Boolean, blnProgBar As Boolean
Private blnViewOnly As Boolean, blnCheckArchive As Boolean, blnHasArchive As Boolean, blnHasArchiveThis As Boolean
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long, blnWasGTR As Boolean
Private blnJustClose As Boolean, blnUpdateTotals As Boolean, blnEnableSub_True As Boolean, blnEnableSub_False As Boolean
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String, strCallersCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim usr As DAO.User, grp As DAO.Group
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp01 As String, dblTmp02 As Double
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If strTmp01 <> vbNullString Then
150           If CharCnt(strTmp01, "~") = 4 Then  ' ** Module Function: modStringFuncs.

160             intPos01 = InStr(strTmp01, "~")
170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strTmp01 = Mid(strTmp01, (intPos01 + 1))
190             intPos01 = InStr(strTmp01, "~")
200             strCallersCallingForm = Left(strTmp01, (intPos01 - 1))
210             strTmp01 = Mid(strTmp01, (intPos01 + 1))
220             intPos01 = InStr(strTmp01, "~")
230             gstrAccountNo = Left(strTmp01, (intPos01 - 1))
240             strTmp01 = Mid(strTmp01, (intPos01 + 1))
250             intPos01 = InStr(strTmp01, "~")
260             blnHasCurrent = CBool(Left(strTmp01, (intPos01 - 1)))  ' ** Has transactions in current statement period, after last statement.
270             lngHids = Val(Mid(strTmp01, (intPos01 + 1)))

280             DoCmd.Hourglass True
290             DoEvents

                'I DON'T UNDERSTAND HOW blnHasCurrent AND blnNoTrans WERE SUPPOSED TO WORK!

                ' ** See if we can get accountno showing right away.
300             .Refresh
310             DoEvents

320             blnWasGTR = gblnGoToReport
330             lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left
340             If gblnGoToReport = True Then
                  ' ** Move the progress bar, even if it's the 2nd report.
350               dblTmp02 = .GoToReport_arw_pr_img.Width
360               ProgBar_Width_Hide Me, dblTmp02, 3  ' ** Module Procedure: modHideTransactions1.
370               .ProgBar_box.Left = (.ProgBar_box.Left + dblTmp02)
380               .ProgBar_box2.Left = (.ProgBar_box2.Left + dblTmp02)
390               .ProgBar_lbl1.Left = (.ProgBar_lbl1.Left + dblTmp02)
400               .Status1_lbl.Left = (.Status1_lbl.Left + dblTmp02)
410               .Status2_lbl.Left = (.Status2_lbl.Left + dblTmp02)
420             End If

430             gblnHasForExThis = False: blnJustClose = False: blnProgBar = False
440             blnUpdateTotals = False: blnCheckArchive = False: blnHasArchive = False: blnHasArchiveThis = False
450             blnEnableSub_True = False: blnEnableSub_False = False
460             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
470             blnGoingToReport = False: blnHasCurrentMsgShown = False
480             blnGTR_Emblem = False: lngGTR_Stat = 0&

                ' ** We know whether this account has archived,
                ' ** but we don't know whether anyone has archived.
490             varTmp00 = DCount("*", "LedgerArchive")
500             If IsNull(varTmp00) = False Then
510               If varTmp00 > 0 Then
520                 blnHasArchive = True
530               End If
540             End If

550             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
560             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
570             DoEvents

580             .cmdOK.Enabled = False
590             .cmdCancel.Enabled = False
600             .cmdUnHide.Enabled = False
610             .cmdUnHide.Visible = False
620             TotalsEnable False  ' ** Procedure: Below.

                ' ** Check user status, and lock all fields if ViewOnly.
630             blnViewOnly = False
640             For Each grp In DBEngine.Workspaces(0).Groups
650               If grp.Name = "ViewOnly" Then
660                 For Each usr In grp.Users
670                   If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
680                     blnViewOnly = True
690                     Exit For
700                   End If
710                 Next
720               End If
730             Next

740             If blnViewOnly = True Then
750               .frmAccountHideTrans2_Sub_List.Form.ledger_HIDDEN.Locked = True
760               .frmAccountHideTrans2_Sub_Pick.Enabled = False
770             End If
780             DoEvents

790             .chkIncludeArchive.Enabled = False

                ' ** Make sure these previous choices are honored below.
800             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

810             .shareface_subtot_box1.Visible = False
820             .shareface_subtot_box2.Visible = False
830             .icash_subtot_box1.Visible = False
840             .icash_subtot_box2.Visible = False
850             .pcash_subtot_box1.Visible = False
860             .pcash_subtot_box2.Visible = False
870             .cost_subtot_box1.Visible = False
880             .cost_subtot_box2.Visible = False
890             .curr_id_box1.Visible = False
900             .curr_id_box2.Visible = False

910             .frmAccountHideTrans2_Sub_Pick.Requery
920             .hidden_trans_display = lngHids '.hidden_trans
930             DoEvents

940             lngShortcutLbl_Offset = .ShortcutMenu_lbl.Top - (.frmAccountHideTrans2_Sub_List.Top + .frmAccountHideTrans2_Sub_List.Height)

950             lngSubList_Height = .frmAccountHideTrans2_Sub_List.Height
960             lngSubPick_Height = .frmAccountHideTrans2_Sub_Pick.Height
970             lngSubHeader_Height = .frmAccountHideTrans2_Sub_List.Form.FormHeader.Height
980             lngSubFooter_Height = .frmAccountHideTrans2_Sub_List.Form.FormFooter.Height
990             lngSubRow_Height = .frmAccountHideTrans2_Sub_List.Form.Detail.Height
1000            lngNavBtns_Height = (17& * lngTpp)  ' ** This is mine, at 15 Twips-Per-Pixel.
1010            dblTmp02 = lngSubList_Height - (lngSubHeader_Height + lngSubFooter_Height + lngNavBtns_Height)
1020            If (dblTmp02 Mod lngSubRow_Height) = 0 Then  ' ** Mine should be an even 11.
1030              lngSubList_Rows = CLng(dblTmp02 / lngSubRow_Height)
1040            Else
1050              lngSubList_Rows = CLng(Round((dblTmp02 / lngSubRow_Height), 0))
1060            End If
1070            dblTmp02 = lngSubPick_Height - (lngSubHeader_Height + lngSubFooter_Height + lngNavBtns_Height)
1080            If (dblTmp02 Mod lngSubRow_Height) = 0 Then  ' ** Mine should be an even 3.
1090              lngSubPick_Rows = CLng(dblTmp02 / lngSubRow_Height)
1100            Else
1110              lngSubPick_Rows = CLng(Round((dblTmp02 / lngSubRow_Height), 0))
1120            End If
1130            DoEvents

                ' ** Shorten by 1 row.
1140            PickSubExpand False  ' ** Procedure: Below.
1150            DoEvents
                ' ** Then by a 2nd.
1160            PickSubExpand False  ' ** Procedure: Below.

1170            If .archive_trans > 0 Then  ' ** This is just this account.
1180              blnHasArchiveThis = True
1190            Else
1200              blnHasArchiveThis = False
1210            End If
1220            ArchiveSet blnHasArchiveThis  ' ** Procedure: Below.
1230            DoEvents

                ' ** Form_Load() hasn't yet triggered as of this point.
1240            If .chkShowAll = False And .chkIncludeArchive = True Then
                  ' ** Prefs shouldn't have done this, but who knows?
1250              .chkIncludeArchive = False
1260            End If

1270            If .archive_trans = 0 Then
                  ' ** If the pref has include trans, and there are none, change it!
1280              If .chkIncludeArchive = True Then
1290                .chkIncludeArchive = False
1300              End If
1310            End If

                ' ** Load, and a couple of Timer's have already gone by the time it gets here.
1320            .archive_trans_lbl.Visible = False
1330            .archive_trans_lbl_box.Visible = False
1340            .LastStmt_Trans_lbl.Visible = True
1350            If InStr(.frmAccountHideTrans2_Sub_List.Form.Filter, "transdate") = 0 Then
1360              If .LastStmt_Trans_lbl.Visible = True Then
1370                .LastStmt_Trans_lbl.Visible = False
1380                If .chkIncludeArchive = True Then
1390                  .archive_trans_lbl.Visible = True
1400                  .archive_trans_lbl_box.Visible = True
1410                End If
1420              End If
1430            End If
1440            DoEvents

                ' *********************************************************
                ' *********************************************************
                ' ** This is the sometimes very long hidden load process.
1450            If .hidden_trans > 0 Then
1460              .cmdHidden.Enabled = False
1470              .cmdPrintReport_lbl2.Visible = False
1480              LedgerHiddenLoad Me  ' ** Module Procedure: modHideTransactions1.
1490              ProgBar_Width_Hide Me, .ProgBar_box.Width, 2  ' ** Module Procedure: modHideTransactions1.
1500              strTmp01 = Format((.ProgBar_bar01.Width / .ProgBar_box.Width), "##0%")
1510              .ProgBar_lbl1.Caption = strTmp01
1520              .cmdHidden.Enabled = True
1530              DoEvents
1540              blnProgBar = True
1550              Beep
1560              .TimerInterval = 500&
1570            Else
                  ' ** If the pref has hidden only, and there are none, change it!
1580              If .chkShowHiddenOnly = True Then
1590                .chkShowHiddenOnly = False
1600              End If
1610              .chkShowHiddenOnly.Enabled = False
1620              .cmdHidden.Enabled = False
1630              DoEvents
1640            End If
                ' *********************************************************
                ' *********************************************************

1650            If lngHids = 0& Then
1660              .cmdPrintReport.Enabled = False
1670              .cmdPrintReport_raised_img_dis.Visible = True
1680              .cmdPrintReport_raised_img.Visible = False
1690            Else
1700              .cmdPrintReport.Enabled = True
1710              .cmdPrintReport_raised_img.Visible = True
1720              .cmdPrintReport_raised_img_dis.Visible = False
1730            End If
1740            .cmdPrintReport_raised_semifocus_dots_img.Visible = False
1750            .cmdPrintReport_raised_focus_img.Visible = False
1760            .cmdPrintReport_raised_focus_dots_img.Visible = False
1770            .cmdPrintReport_sunken_focus_dots_img.Visible = False
1780            DoEvents

                '.chkIncludeArchive = False
                '.chkIncludeArchive.Enabled = False
                '.chkIncludeArchive_lbl.Visible = False
                '.chkIncludeArchive_lbl2.Visible = True
                '.chkIncludeArchive_lbl2_dim_hi.Visible = True
                '.chkIncludeArchive_lbl3.Visible = True
                '.chkIncludeArchive_lbl3_dim_hi.Visible = True
                '.chkIncludeArchive_lbl4.Visible = True
                '.chkIncludeArchive_lbl4_dim_hi.Visible = True
                '.chkIncludeArchive_lbl5.Visible = True
                '.chkIncludeArchive_lbl5_dim_hi.Visible = True

1790          Else
1800            Cancel = -1
1810          End If
1820        Else
1830          Cancel = -1
1840        End If
1850      Else
1860        Cancel = -1
1870      End If
1880    End With

1890    If Cancel = -1 Then
1900      DoCmd.Hourglass False
1910      Beep
1920      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1930    End If

EXITP:
1940    Set usr = Nothing
1950    Set grp = Nothing
1960    Exit Sub

ERRH:
1970    DoCmd.Hourglass False
1980    Select Case ERR.Number
        Case Else
1990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2000    End Select
2010    Resume EXITP

End Sub

Private Sub Form_Load()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

2110    With Me

2120      CLR_DISABLED_FG = CLR_DKGRY
2130      CLR_DISABLED_BG = CLR_LTTEAL

2140      .accountno.ForeColor = CLR_DISABLED_FG
2150      .accountno.BackColor = CLR_DISABLED_BG
2160      .shortname.ForeColor = CLR_DISABLED_FG
2170      .shortname.BackColor = CLR_DISABLED_BG
2180      .hidden_trans_display.ForeColor = CLR_DISABLED_FG
2190      .hidden_trans_display.BackColor = CLR_DISABLED_BG
2200      .curr_id.ForeColor = CLR_DISABLED_FG
2210      .curr_id.BackColor = CLR_DISABLED_BG
2220      DoEvents

2230      lngCurrID_Width = (.frmAccountHideTrans2_Sub_List.Form.ledger_HIDDEN_lbl_line.Left - .frmAccountHideTrans2_Sub_List.Form.curr_id.Left)
2240      lngSharefaceSub_Left = .shareface_subtot.Left
2250      lngICashSub_Left = .icash_subtot.Left
2260      lngPCashSub_Left = .pcash_subtot.Left
2270      lngCostSub_Left = .cost_subtot.Left
2280      lngCurrID_Left = .curr_id.Left
2290      lngSubTotLbl_Offset = (.shareface_subtot.Left - .subtot_lbl.Left)
2300      lngOK_Left = .cmdOK.Left
2310      lngCancel_Left = .cmdCancel.Left
2320      lngPTrans_Left = .pending_trans.Left
2330      lngPTransLbl_Offset = (.pending_trans.Left - .pending_trans_lbl.Left)

2340      chkShowAll_AfterUpdate  ' ** Procedure: Below.
2350      If .chkIncludeArchive.Enabled = True Then
            ' ** Wait before turning this on.
2360        .chkIncludeArchive.Enabled = False
2370      End If
2380      chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
2390      If .chkIncludeArchive.Enabled = True Then
2400        .chkIncludeArchive.Enabled = False
2410      End If
2420      chkShowHiddenOnly_AfterUpdate  ' ** Procedure: Below.
2430      If .chkIncludeArchive.Enabled = True Then
2440        .chkIncludeArchive.Enabled = False
2450      End If
2460      chkShowJournalNo_AfterUpdate  ' ** Procedure: Below.
2470      If .chkIncludeArchive.Enabled = True Then
2480        .chkIncludeArchive.Enabled = False
2490      End If
2500      chkShowAll_AfterUpdate  ' ** Procedure: Below.  ' ** I know it's twice!
2510      If .chkIncludeArchive.Enabled = True Then
2520        .chkIncludeArchive.Enabled = False
2530      End If

2540      If InStr(.frmAccountHideTrans2_Sub_List.Form.Filter, "transdate") = 0 Then
2550        If .LastStmt_Trans_lbl.Visible = True Then
2560          .LastStmt_Trans_lbl.Visible = False
2570        End If
2580      End If

2590      chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.

2600      .frmAccountHideTrans2_Sub_List.Form.FilterOn = True
2610      .frmAccountHideTrans2_Sub_List.SetFocus

2620      If .chkIncludeArchive = True Then
2630        blnCheckArchive = True
2640      End If

2650      blnHasForEx_Chk = True

2660      DoCmd.Hourglass True  ' ** Make sure it's still running.
2670      DoEvents
2680      .TimerInterval = 250&

2690    End With

EXITP:
2700    Exit Sub

ERRH:
2710    DoCmd.Hourglass False
2720    Select Case ERR.Number
        Case Else
2730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2740    End Select
2750    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
2810    intShiftDown = (Shift And acShiftMask) > 0
2820    intAltDown = (Shift And acAltMask) > 0
2830    intCtrlDown = (Shift And acCtrlMask) > 0

2840    intRetVal = KeyCode

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   ledger_HIDDEN:    D {ledger_HIDDEN on frmAccountHideTrans2_Sub_List}
        ' **   ledger_HIDDEN:    N {ledger_HIDDEN on frmAccountHideTrans2_Sub_Pick}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAccountHideTrans2_Sub_List, frmAccountHideTrans2_Sub_Pick}

        ' ** Alt keys.
2850    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyD
2870        With Me
2880          intRetVal = 0
2890          .frmAccountHideTrans2_Sub_List.SetFocus
2900          .frmAccountHideTrans2_Sub_List.Form.ledger_HIDDEN.SetFocus  ' ** Don't invoke it!
2910        End With
2920      Case vbKeyN
2930        With Me
2940          intRetVal = 0
2950          If .running_pick_cnt > 0& Then
2960            .frmAccountHideTrans2_Sub_Pick.SetFocus
2970            .frmAccountHideTrans2_Sub_Pick.Form.ledger_HIDDEN.SetFocus  ' ** Don't invoke it!
2980          Else
2990            Beep
3000          End If
3010        End With
3020      Case vbKeyX
3030        intRetVal = 0
3040        cmdClose_Click  ' ** Procedure: Below.
3050      End Select
3060    End If

        ' ** Ctrl keys.
3070    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3080      Select Case intRetVal
          Case vbKeyS
            ' ** Do nothing.
3090        intRetVal = 0
3100      End Select
3110    End If

        ' ** Ctrl-Shift keys.
3120    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3130      Select Case intRetVal
          Case vbKeyF
3140        With Me
3150          intRetVal = 0
3160          .FocusHolder.SetFocus
3170        End With
3180      End Select
3190    End If

EXITP:
3200    KeyCode = intRetVal
3210    Exit Sub

ERRH:
3220    intRetVal = 0
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub Form_Timer()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long

3310    With Me
3320      .TimerInterval = 0&
3330      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
3340        lngGTR_Stat = lngGTR_Stat + 1&
3350        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
3360        lngTmp01 = lngTmp01 + 1&
3370        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
3380        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
3390        If lngTmp01 > 1& Then
3400          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
3410        End If
3420        If lngTmp01 = 24& Then
3430          lngTmp01 = 0&
3440          lngTmp02 = lngTmp02 + 1&
3450          If lngTmp02 < 5& Then
3460            .TimerInterval = 50&
3470          Else
3480            blnGTR_Emblem = False
3490          End If
3500        Else
3510          .TimerInterval = 50&
3520        End If
3530      End If  ' ** blnGTR_Emblem.
3540      If gblnSetFocus = True Then
3550        gblnSetFocus = False
3560        strFormName = vbNullString
3570  On Error Resume Next
3580        Do Until strFormName = THIS_NAME
3590          DoCmd.SelectObject acForm, THIS_NAME, False
3600          strFormName = Screen.ActiveForm.Name
3610        Loop
3620  On Error GoTo ERRH
3630      End If
3640      If blnHasCurrent = False And gblnGoToReport = False And blnHasCurrentMsgShown = False Then
3650        blnHasCurrentMsgShown = True
3660        Beep
3670        DoCmd.Hourglass False
3680        msgResponse = MsgBox("There are no transactions since the last Statement." & vbCrLf & vbCrLf & _
              "Whould you like to see all the transactions?", vbQuestion + vbOKCancel, "No Current Transactions")
3690        If msgResponse = vbOK Then
3700          .chkShowAll = True
3710          chkShowAll_AfterUpdate  ' ** Procedure: Below.
3720        End If
3730        DoCmd.Hourglass True
3740        DoEvents
3750        blnEnableSub_False = True
3760        .TimerInterval = 250&
3770      ElseIf blnUpdateTotals = True Then
3780        blnUpdateTotals = False
3790        varTmp00 = DCount("*", "tmpAccountHideTrans_03")
3800        lngRecsCur = Nz(varTmp00, 0)
3810        .pending_trans = lngRecsCur
3820        If lngRecsCur > 0& Then
3830          varTmp00 = DSum("[sharefacex]", "tmpAccountHideTrans_03")
3840          .shareface_subtot = CDbl(Nz(varTmp00, 0))
3850          varTmp00 = DSum("[shareface_divint]", "tmpAccountHideTrans_03")
3860          .shareface_divint_subtot = CDbl(Nz(varTmp00, 0))
3870          varTmp00 = DSum("[icash]", "tmpAccountHideTrans_03")
3880          .icash_subtot = CCur(Nz(varTmp00, 0))
3890          varTmp00 = DSum("[pcash]", "tmpAccountHideTrans_03")
3900          .pcash_subtot = CCur(Nz(varTmp00, 0))
3910          varTmp00 = DSum("[cost]", "tmpAccountHideTrans_03")
3920          .cost_subtot = CCur(Nz(varTmp00, 0))
3930          TotalsEnable True  ' ** Procedure: Below.
3940        Else
3950          .shareface_subtot = 0#
3960          .shareface_divint_subtot = 0#
3970          .icash_subtot = 0@
3980          .pcash_subtot = 0@
3990          .cost_subtot = 0@
4000          TotalsEnable False  ' ** Procedure: Below.
4010        End If
4020        ValidateGroup  ' ** Procedure: Below.
4030        If .hidden_trans > 0& Then
4040          If .cmdPrintReport.Enabled = False Then
4050            .cmdPrintReport.Enabled = True
4060            .cmdPrintReport_raised_img.Visible = True
4070            .cmdPrintReport_raised_img_dis.Visible = False
4080          End If
4090        Else
4100          If .cmdPrintReport.Enabled = True Then
4110            .cmdPrintReport.Enabled = False
4120            .cmdPrintReport_raised_img_dis.Visible = True
4130            .cmdPrintReport_raised_img.Visible = False
4140          End If
4150        End If
4160        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4170        .cmdPrintReport_raised_focus_img.Visible = False
4180        .cmdPrintReport_raised_focus_dots_img.Visible = False
4190        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4200      ElseIf blnEnableSub_True = True Then
4210        blnEnableSub_True = False
4220        .frmAccountHideTrans2_Sub_Pick.Form.SubEnable True  ' ** Form Procedure: frmAccountHideTrans2_Sub_Pick.
4230        DoCmd.Hourglass False
4240      ElseIf blnEnableSub_False = True Then
4250        blnEnableSub_False = False
4260        .frmAccountHideTrans2_Sub_Pick.Form.SubEnable False  ' ** Form Procedure: frmAccountHideTrans2_Sub_Pick.
4270        DoCmd.Hourglass False
4280      End If
4290      If gblnGoToReport = True Then
4300        DoCmd.Hourglass True  ' ** Make sure it's still running.
4310        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or (lngGTR_Stat > 96&)) Then
              ' ** Just let the emblem go.
4320        Else
4330          DoCmd.Hourglass True  ' ** Make sure it's still running.
4340          If blnGTR_Emblem = False And IsNull(garr_varGoToReport(GTR_CTL3)) = False Then
4350            blnGTR_Emblem = True
4360            lngTmp01 = 0&: lngTmp02 = 0&
4370          End If
4380          DoEvents
4390          Select Case blnGoingToReport
              Case True
4400            Select Case IsNull(garr_varGoToReport(GTR_CTL3))
                Case True
                  ' ** Just sit there, and we're done.
4410              DoCmd.Hourglass False  ' ** It should already be off.
4420              .GoToReport_arw_pr_img.Visible = False
4430              .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
4440              blnGoingToReport = False
4450              gblnGoToReport = False
4460              blnGTR_Emblem = False
4470              DoEvents
4480            Case False
4490              If .Controls(garr_varGoToReport(GTR_CTL3)).Enabled = True Then
4500                .TimerInterval = 50&
4510                cmdHidden_Click  ' ** Procedure: Below
4520              Else
                    ' ** Just sit there, and we're done.
4530                .GoToReport_arw_hide_img.Visible = False
4540                blnGoingToReport = False
4550                DoCmd.Hourglass False
4560                gblnGoToReport = False
4570                blnGTR_Emblem = False
4580                GTREmblem_Off  ' ** Procedure: Below.
4590                Beep
4600              End If
4610            End Select
4620          Case False
4630            blnGoingToReport = True
4640            If blnHasCurrent = False Then
4650              blnHasCurrentMsgShown = True
4660              .chkShowAll = True
4670              chkShowAll_AfterUpdate  ' ** Procedure: Below.
4680              DoEvents
4690              .frmAccountHideTrans2_Sub_Pick.Form.SubEnable False  ' ** Form Procedure: frmAccountHideTrans2_Sub_Pick.
4700              DoEvents
4710            End If
4720            Select Case IsNull(garr_varGoToReport(GTR_CTL3))
                Case True
4730              If .cmdPrintReport.Enabled = True Then
4740                .cmdPrintReport.SetFocus
4750              End If
4760              .cmdPrintReport_lbl2.Left = (lngPrintReportLbl_Left + .GoToReport_arw_pr_img.Width)
4770              .GoToReport_arw_pr_img.Visible = True
4780              Beep
4790              DoCmd.Hourglass False
4800              DoEvents
4810              .TimerInterval = (GTR_WAIT * 2&)
4820            Case False
4830              lngTmp03 = 0&
4840              If IsNull(.hidden_trans) = False Then
4850                lngTmp03 = .hidden_trans
4860              End If
4870              .GoToReport_arw_hide_img.Visible = True
4880              DoEvents
4890              If .Controls(garr_varGoToReport(GTR_CTL3)).Enabled = True Then
4900                .Controls(garr_varGoToReport(GTR_CTL3)).SetFocus
4910              End If
4920              If lngTmp03 > 0& Then
4930                .TimerInterval = 50&
4940              Else
4950                .TimerInterval = 0&
4960                Beep
4970                DoCmd.Hourglass False
4980                MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                      "There are no hidden transactions to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
4990                blnGoingToReport = False
5000                .GoToReport_arw_hide_img.Visible = False
5010                gblnGoToReport = False
5020                blnGTR_Emblem = False
5030                GTREmblem_Off  ' ** Procedure: Below.
5040              End If
5050            End Select
5060          End Select
5070        End If
5080      End If
5090      If blnCheckArchive = True Then
5100        blnCheckArchive = False
5110        chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
5120        DoEvents
5130        blnHasForEx_Chk = True
5140      End If
5150      If blnHasForEx_Chk = True Then
5160        blnHasForEx_Chk = False
5170        gblnHasForExThis = HasForEx_Acct(gstrAccountNo, "L")  ' ** Module Function: modCurrencyFuncs.
5180        Select Case gblnHasForEx
            Case True
5190          .chkIncludeCurrency.Enabled = True
5200          Select Case gblnHasForExThis
              Case True
5210            .chkIncludeCurrency = True
5220            .chkIncludeCurrency.Locked = True
5230            chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
5240          Case False
5250            .chkIncludeCurrency.Locked = False
5260          End Select
5270        Case False
5280          .chkIncludeCurrency = False
5290          .chkIncludeCurrency.Enabled = False
5300          chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
5310        End Select
5320        DoCmd.Hourglass False
5330      End If
5340      If blnProgBar = True Then
5350        blnProgBar = False
5360        .ProgBar_lbl1.Visible = False
5370        ProgBar_Width_Hide Me, False, 1  ' ** Module Procedure: modHideTransactions1.
5380        .ProgBar_box.Visible = False
5390        .ProgBar_box2.Visible = False
5400        .Status1_lbl.Visible = False
5410        .Status2_lbl.Visible = False
5420        .cmdPrintReport_lbl2.Visible = True
5430        .chkShowAll.Enabled = True
5440        .chkShowJournalNo.Enabled = True
5450        .chkShowHiddenOnly.Enabled = True
5460        If gblnGoToReport = True Then
5470          If .Controls(garr_varGoToReport(GTR_CTL3)).Enabled = True Then
5480            .Controls(garr_varGoToReport(GTR_CTL3)).SetFocus
5490          End If
5500        End If
5510        DoCmd.Hourglass False
5520        DoEvents
5530      End If
5540      If blnHasArchiveThis = True Then
5550        .chkIncludeArchive.Enabled = True
5560      Else
5570        .chkIncludeArchive.Enabled = False
5580      End If
5590      If gblnHasForExThis = False And .chkIncludeCurrency.Enabled = False And .chkIncludeCurrency = True Then
5600        .chkIncludeCurrency = False
5610        chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
5620        DoCmd.Hourglass False
5630      End If
5640    End With

EXITP:
5650    Exit Sub

ERRH:
5660    DoCmd.Hourglass False
5670    Select Case ERR.Number
        Case Else
5680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5690    End Select
5700    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String

5810    With Me

5820      DoCmd.Hourglass True
5830      DoEvents

          ' ** Make sure all reports are closed.
5840      If Reports.Count > 0 Then
5850        Do While Reports.Count > 0
5860          DoCmd.Close acReport, Reports(0).Name
5870          DoEvents
5880        Loop
5890      End If

5900      If .pending_trans > 0& Then
5910        strTmp01 = vbNullString
5920        If .cmdUnHide.Visible = True Then
5930          strTmp01 = "unhide"
5940        Else
5950          strTmp01 = "hide"
5960        End If
5970        Beep
5980        DoCmd.Hourglass False
5990        msgResponse = MsgBox("You have transactions chosen to " & strTmp01 & " that are still pending." & vbCrLf & _
              "Are you sure you want to abandon those changes?", vbQuestion + vbYesNo, "Changes Not Saved")
6000        If msgResponse <> vbYes Then
6010          Cancel = -1
6020        End If
6030      Else
6040        msgResponse = vbYes
6050      End If

6060      If msgResponse = vbYes Then

6070        DoCmd.Hourglass True
6080        DoEvents

6090        EmptyTemps False  ' ** Procedure: Below.

6100        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

6110        Select Case strCallersCallingForm
            Case "frmMenu_Account"
6120          gblnSetFocus = True
6130          gstrReturningForm = THIS_NAME
6140          DoCmd.OpenForm strCallersCallingForm, , , , , , "frmMenu_Main"
6150        End Select

6160      End If

6170    End With

EXITP:
6180    Exit Sub

ERRH:
6190    DoCmd.Hourglass False
6200    Select Case ERR.Number
        Case Else
6210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6220    End Select
6230    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

6310    With Me
6320      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
6330        Select Case blnPrintReport_Focus
            Case True
6340          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
6350          .cmdPrintReport_raised_img.Visible = False
6360        Case False
6370          .cmdPrintReport_raised_img.Visible = True
6380          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6390        End Select
6400        .cmdPrintReport_raised_focus_img.Visible = False
6410        .cmdPrintReport_raised_focus_dots_img.Visible = False
6420        .cmdPrintReport_sunken_focus_dots_img.Visible = False
6430        .cmdPrintReport_raised_img_dis.Visible = False
6440      End If
6450    End With

EXITP:
6460    Exit Sub

ERRH:
6470    Select Case ERR.Number
        Case Else
6480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6490    End Select
6500    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

6610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6620    Exit Sub

ERRH:
6630    Select Case ERR.Number
        Case Else
6640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6650    End Select
6660    Resume EXITP

End Sub

Public Sub cmdClose_Click()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

6710    With Me

6720      .TimerInterval = 0&

6730      DoCmd.Hourglass True
6740      DoEvents

6750      blnJustClose = False
6760      DoCmd.Close acForm, THIS_NAME

6770    End With

EXITP:
6780    Exit Sub

ERRH:
6790    DoCmd.Hourglass False
6800    Select Case ERR.Number
        Case 2455  ' ** You entered an expressions that has an invalid reference to the property |.
          ' ** Do nothing.
6810    Case 2501  ' ** The Close action was canceled.
          ' ** Ignore.
6820    Case Else
6830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6840    End Select
6850    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
6990          .frmAccountHideTrans2_Sub_List.SetFocus
7000          .frmAccountHideTrans2_Sub_List.Form.ledger_HIDDEN.SetFocus
7010        End With
7020      End Select
7030    End If

        ' ** Shift keys.
7040    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7050      Select Case intRetVal
          Case vbKeyTab
7060        With Me
7070          intRetVal = 0
7080          If .cmdHidden.Enabled = True Then
7090            .cmdHidden.SetFocus
7100          ElseIf .cmdPrintReport.Enabled = True Then
7110            .cmdPrintReport.SetFocus
7120          ElseIf .chkShowHiddenOnly.Enabled = True Then
7130            .chkShowHiddenOnly.SetFocus
7140          ElseIf .chkShowJournalNo.Enabled = True Then
7150            .chkShowJournalNo.SetFocus
7160          ElseIf .chkShowAll.Enabled = True Then
7170            .chkShowAll.SetFocus
7180          ElseIf .chkIncludeCurrency.Enabled = True Then
7190            .chkIncludeCurrency.SetFocus
7200          ElseIf .chkIncludeArchive.Enabled = True Then
7210            .chkIncludeArchive.SetFocus
7220          ElseIf .cmdCancel.Enabled = True Then
7230            .cmdCancel.SetFocus
7240          Else
7250            Beep
7260          End If
7270        End With
7280      End Select
7290    End If

EXITP:
7300    KeyCode = intRetVal
7310    Exit Sub

ERRH:
7320    intRetVal = 0
7330    Select Case ERR.Number
        Case Else
7340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7350    End Select
7360    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
7410    intShiftDown = (Shift And acShiftMask) > 0
7420    intAltDown = (Shift And acAltMask) > 0
7430    intCtrlDown = (Shift And acCtrlMask) > 0

7440    intRetVal = KeyCode

        ' ** Plain keys.
7450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7470        With Me
7480          intRetVal = 0
7490          If .cmdPrintReport.Enabled = True Then
7500            .cmdPrintReport.SetFocus
7510          ElseIf .cmdHidden.Enabled = True Then
7520            .cmdHidden.SetFocus
7530          Else
7540            .cmdClose.SetFocus
7550          End If
7560        End With
7570      End Select
7580    End If

        ' ** Shift keys.
7590    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7600      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7610        With Me
7620          intRetVal = 0
7630          If .chkShowHiddenOnly.Enabled = True Then
7640            .chkShowHiddenOnly.SetFocus
7650          ElseIf .chkShowJournalNo.Enabled = True Then
7660            .chkShowJournalNo.SetFocus
7670          ElseIf .chkShowAll.Enabled = True Then
7680            .chkShowAll.SetFocus
7690          ElseIf .chkIncludeCurrency.Enabled = True Then
7700            .chkIncludeCurrency.SetFocus
7710          ElseIf .chkIncludeArchive.Enabled = True Then
7720            .chkIncludeArchive.SetFocus
7730          ElseIf .cmdCancel.Enabled = True Then
7740            .cmdCancel.SetFocus
7750          ElseIf .cmdOK.Visible = True And .cmdOK.Enabled = True Then
7760            .cmdOK.SetFocus
7770          ElseIf .cmdUnHide.Visible = True And .cmdUnHide.Enabled = True Then
7780            .cmdUnHide.SetFocus
7790          Else
7800            .frmAccountHideTrans2_Sub_List.SetFocus
7810            .frmAccountHideTrans2_Sub_List.Form.ledger_HIDDEN.SetFocus
7820          End If
7830        End With
7840      End Select
7850    End If

EXITP:
7860    KeyCode = intRetVal
7870    Exit Sub

ERRH:
7880    intRetVal = 0
7890    Select Case ERR.Number
        Case Else
7900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7910    End Select
7920    Resume EXITP

End Sub

Private Sub cmdOK_Click()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngCurList_Height As Long, lngCurPick_Height As Long
        Dim lngCurList_Rows As Long, lngCurPick_Rows As Long
        Dim lngPendingTrans As Long
        Dim strTmp01 As String, dblTmp02 As Double

8010    With Me
8020      DoCmd.Hourglass True
8030      DoEvents
8040      lngHidType = 0&
8050      If ValidateGroup = True Then  ' ** Function: Below.

8060        lngPendingTrans = .pending_trans

8070        lngCurList_Height = .frmAccountHideTrans2_Sub_List.Height
8080        lngCurPick_Height = .frmAccountHideTrans2_Sub_Pick.Height

8090        Set dbs = CurrentDb
8100        With dbs

              ' ** Update Ledger, for ledger_HIDDEN = True, via subquery to tmpAccountHideTrans_03.
8110          Set qdf = .QueryDefs("qryAccountHideTrans2_19a")
8120          qdf.Execute
8130          Set qdf = Nothing
              ' ** Update LedgerArchive, for ledger_HIDDEN = True, via subquery to tmpAccountHideTrans_03.
8140          Set qdf = .QueryDefs("qryAccountHideTrans2_19b")
8150          qdf.Execute
8160          Set qdf = Nothing

              ' ** Examples:
              ' **   accountno      astno jno1   jno2   jtype1    jtype2
              ' **   =============== ==== ====== ====== ========= =========
              ' **   000000000002300_0412_020353_020350_Withdrawn_Deposit__...
              ' **   000000000002300_0000_040395_000000_Misc.______________
              ' **          15         4     6      6       9         9      ' ** Might someone get more than 9,999 assets?
              ' ** Whichever one may have an assetno gets the slot.

              ' ** Append qryAccountHideTrans2_23_02_01 (qryAccountHideTrans2_21 (tmpAccountHideTrans_03, linked to
              ' ** qryAccountHideTrans2_20 (tmpAccountHideTrans_03, grouped, with ledghid_cnt, Max(assetno)), with
              ' ** ledghid_uniqueidx, jno, jtype), linked to qryAccountHideTrans2_22_02_22 (qryAccountHideTrans2_22_02_21
              ' ** (qryAccountHideTrans2_22_02_20 (Union of qryAccountHideTrans2_22_02_01 (qryAccountHideTrans2_21
              ' ** (tmpAccountHideTrans_03, linked to qryAccountHideTrans2_20 (tmpAccountHideTrans_03, grouped, with
              ' ** ledghid_cnt, Max(assetno)), with ledghid_uniqueidx, jno, jtype), just ledghid_ord = 1 of 2),
              ' ** qryAccountHideTrans2_22_02_02 (qryAccountHideTrans2_21 (tmpAccountHideTrans_03, linked to
              ' ** qryAccountHideTrans2_20 (tmpAccountHideTrans_03, grouped, with ledghid_cnt, Max(assetno)), with
              ' ** ledghid_uniqueidx, jno, jtype), just ledghid_ord = 2)), grouped to get 1 record), with
              ' ** ledghid_uniqueid), for ledghid_cnt = 2) to tblLedgerHidden, by specified [hidtyp].
8170          Set qdf = .QueryDefs("qryAccountHideTrans2_23_" & Right("00" & CStr(lngPendingTrans), 2) & "_02")
8180          With qdf.Parameters
8190            ![hidtyp] = lngHidType
8200          End With
8210          qdf.Execute dbFailOnError
8220          Set qdf = Nothing

              ' ** LedgHidType enumeration:
              ' **   0  GRP_NONE    Unmatched hidden transactions. This should only be a temporary designation.
              ' **   1  NORM        2 entries in hidden group, with matching assetno (which could both be zero).
              ' **   2  NORM_MISC   2 entries in hidden group, where both are 'Misc.', to be treated like a normal pair.
              ' **   3  MISC_2_GRP  2 entries in hidden group, one 'Misc.' and one other.
              ' **   4  MISC_3_GRP  3 entries in hidden group, one 'Misc.' and two other matching assetno.
              ' **   5  MULTI_GRP   3 or more entries in hidden group, with matching assetno, multi-lot group

8230          EmptyTemps False  ' ** Procedure: Below.

              ' ** Append qryAccountHideTrans2_10 (qryAccountHideTrans2_08 (Account, linked to qryAccountHideTrans2_05
              ' ** (LedgerArchive, grouped by accountno, with cnt, by specified [actno]), qryAccountHideTrans2_06
              ' ** (Union of qryAccountHideTrans2_06a (Ledger, grouped by accountno, just ledger_HIDDEN = True, with cnt,
              ' ** by specified [actno]), qryAccountHideTrans2_06b (LedgerArchive, grouped by accountno, just ledger_HIDDEN = True,
              ' ** with cnt, by specified [actno])), by specified [actno]), linked to qryAccountHideTrans2_09 (tblForm_Graphics,
              ' ** for frmAccountHideTrans2, by specified [dbid]), with add'l fields) to tmpAccountHideTrans_01.
8240          Set qdf = .QueryDefs("qryAccountHideTrans2_11")
8250          With qdf.Parameters
8260            ![dbid] = 1
8270            ![actno] = gstrAccountNo
8280          End With
8290          qdf.Execute
8300          Set qdf = Nothing

              ' ** Append qryAccountHideTrans2_12 (Union of qryAccountHideTrans2_12a (Ledger, by specified [actno]),
              ' ** qryAccountHideTrans2_12b (LedgerArchive, by specified [actno])) to tmpAccountHideTrans_02.
8310          Set qdf = .QueryDefs("qryAccountHideTrans2_13")
8320          With qdf.Parameters
8330            ![actno] = gstrAccountNo
8340          End With
8350          qdf.Execute
8360          Set qdf = Nothing

8370          EmptyTemps True  ' ** Procedure: Below.

8380          .Close
8390        End With
8400        DoEvents

8410        .assetno = 0&
8420        .curr_id = 0&

8430        .cmdClose.SetFocus
8440        .cmdOK.Enabled = False
8450        .cmdCancel.Enabled = False

8460        .shareface_subtot_box1.Visible = False
8470        .shareface_subtot_box2.Visible = False
8480        .icash_subtot_box1.Visible = False
8490        .icash_subtot_box2.Visible = False
8500        .pcash_subtot_box1.Visible = False
8510        .pcash_subtot_box2.Visible = False
8520        .cost_subtot_box1.Visible = False
8530        .cost_subtot_box2.Visible = False
8540        .curr_id_box1.Visible = False
8550        .curr_id_box2.Visible = False

8560        .Requery
8570        .frmAccountHideTrans2_Sub_List.Form.Requery
8580        .frmAccountHideTrans2_Sub_Pick.Form.Requery

8590        .hidden_trans_display = .hidden_trans
8600        .pending_trans = 0&
8610        .shareface_subtot = 0#
8620        .shareface_divint_subtot = 0#
8630        .icash_subtot = 0@
8640        .pcash_subtot = 0@
8650        .cost_subtot = 0@
8660        strTmp01 = "$#,##0.00;($#,##0.00);$0.00;$0.00"
8670        .icash_subtot.Format = strTmp01
8680        .pcash_subtot.Format = strTmp01
8690        .cost_subtot.Format = strTmp01

8700        TotalsEnable False  ' ** Procedure: Below.
8710        UpdateHiddenTrans  ' ** Procedure: Below.
8720        CheckBoxLock False  ' ** Procedure: Below.
8730        .running_pick_cnt = 0&
8740        DoEvents

8750        dblTmp02 = lngCurList_Height - (lngSubHeader_Height + lngSubFooter_Height + lngNavBtns_Height)
8760        If (dblTmp02 Mod lngSubRow_Height) = 0 Then
8770          lngCurList_Rows = CLng(dblTmp02 / lngSubRow_Height)
8780        Else
8790          lngCurList_Rows = CLng(Round((dblTmp02 / lngSubRow_Height), 0))
8800        End If
8810        dblTmp02 = lngCurPick_Height - (lngSubHeader_Height + lngSubFooter_Height + lngNavBtns_Height)
8820        If (dblTmp02 Mod lngSubRow_Height) = 0 Then
8830          lngCurPick_Rows = CLng(dblTmp02 / lngSubRow_Height)
8840        Else
8850          lngCurPick_Rows = CLng(Round((dblTmp02 / lngSubRow_Height), 0))
8860        End If

            ' ** Pick starts out 2 rows shorter than design view.
8870        If lngCurPick_Rows > (lngSubPick_Rows - 2&) Or lngCurList_Rows < (lngSubList_Rows + 2&) Then
8880          Do Until lngCurPick_Rows = (lngSubPick_Rows - 2&)
8890            PickSubExpand False  ' ** Procedure: Below
8900            DoEvents
8910            lngCurPick_Rows = lngCurPick_Rows - 1&
8920          Loop
8930        End If

8940        If .journalno1 > 0& Then
              ' ** Move to 1st journalno in just-hidden group.
8950          .frmAccountHideTrans2_Sub_List.Form.MoveRec 0, .journalno1  ' ** Form Procedure: frmAccountHideTrans2_Sub_List.
8960          DoEvents
8970          .journalno1 = 0&
8980        End If

8990        If .chkShowHiddenOnly.Enabled = False Then
9000          .chkShowHiddenOnly.Enabled = True
9010        End If

9020        If .cmdPrintReport.Enabled = False Then
9030          .cmdPrintReport.Enabled = True
9040          .cmdPrintReport_raised_img.Visible = True
9050          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9060          .cmdPrintReport_raised_focus_img.Visible = False
9070          .cmdPrintReport_raised_focus_dots_img.Visible = False
9080          .cmdPrintReport_sunken_focus_dots_img.Visible = False
9090          .cmdPrintReport_raised_img_dis.Visible = False
9100        End If

9110      Else
9120        Beep
9130      End If
9140      Beep
9150      DoCmd.Hourglass False
9160    End With

EXITP:
9170    Set qdf = Nothing
9180    Set dbs = Nothing
9190    Exit Sub

ERRH:
9200    DoCmd.Hourglass False
9210    Select Case ERR.Number
        Case Else
9220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9230    End Select
9240    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
9310    intShiftDown = (Shift And acShiftMask) > 0
9320    intAltDown = (Shift And acAltMask) > 0
9330    intCtrlDown = (Shift And acCtrlMask) > 0

9340    intRetVal = KeyCode

        ' ** Plain keys.
9350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9360      Select Case intRetVal
          Case vbKeyTab
9370        With Me
9380          intRetVal = 0
9390          .cmdCancel.SetFocus
9400        End With
9410      End Select
9420    End If

        ' ** Shift keys.
9430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9440      Select Case intRetVal
          Case vbKeyTab
9450        With Me
9460          intRetVal = 0
9470          lngRecsCur = .frmAccountHideTrans2_Sub_Pick.Form.RecCnt
9480          If lngRecsCur > 0& Then
9490            .frmAccountHideTrans2_Sub_Pick.SetFocus
9500          Else
9510            .frmAccountHideTrans2_Sub_List.SetFocus
9520          End If
9530        End With
9540      End Select
9550    End If

EXITP:
9560    KeyCode = intRetVal
9570    Exit Sub

ERRH:
9580    intRetVal = 0
9590    Select Case ERR.Number
        Case Else
9600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9610    End Select
9620    Resume EXITP

End Sub

Private Sub cmdUnHide_Click()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnHide_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngCurList_Height As Long, lngCurPick_Height As Long
        Dim lngCurList_Rows As Long, lngCurPick_Rows As Long
        Dim lngPendingTrans As Long, lngUnHiddenTrans As Long
        Dim strTmp01 As String, dblTmp02 As Double

9710    With Me

9720      DoCmd.Hourglass True
9730      DoEvents

9740      lngPendingTrans = .pending_trans
9750      lngUnHiddenTrans = .unhidden_trans  ' ** The size of the original group.

9760      lngCurList_Height = .frmAccountHideTrans2_Sub_List.Height
9770      lngCurPick_Height = .frmAccountHideTrans2_Sub_Pick.Height

9780      Set dbs = CurrentDb
9790      With dbs

            ' ** Update Ledger, for ledger_HIDDEN = False, via subquery to tmpAccountHideTrans_03.
9800        Set qdf = .QueryDefs("qryAccountHideTrans2_19c")
9810        qdf.Execute
9820        Set qdf = Nothing
            ' ** Update LedgerArchive, for ledger_HIDDEN = False, via subquery to tmpAccountHideTrans_03.
9830        Set qdf = .QueryDefs("qryAccountHideTrans2_19d")
9840        qdf.Execute
9850        Set qdf = Nothing

            ' ** Delete tblLedgerHidden, via subquery to tmpAccountHideTrans_03.
9860        Set qdf = .QueryDefs("qryAccountHideTrans2_29")
9870        qdf.Execute
9880        Set qdf = Nothing

9890        EmptyTemps False  ' ** Procedure: Below

            ' ** Append qryAccountHideTrans2_10 (qryAccountHideTrans2_08 (Account, linked to qryAccountHideTrans2_05
            ' ** (LedgerArchive, grouped by accountno, with cnt, by specified [actno]), qryAccountHideTrans2_06
            ' ** (Union of qryAccountHideTrans2_06a (Ledger, grouped by accountno, just ledger_HIDDEN = True, with cnt,
            ' ** by specified [actno]), qryAccountHideTrans2_06b (LedgerArchive, grouped by accountno, just ledger_HIDDEN = True,
            ' ** with cnt, by specified [actno])), by specified [actno]), linked to qryAccountHideTrans2_09 (tblForm_Graphics,
            ' ** for frmAccountHideTrans2, by specified [dbid]), with add'l fields) to tmpAccountHideTrans_01.
9900        Set qdf = .QueryDefs("qryAccountHideTrans2_11")
9910        With qdf.Parameters
9920          ![dbid] = 1
9930          ![actno] = gstrAccountNo
9940        End With
9950        qdf.Execute
9960        Set qdf = Nothing

            ' ** Append qryAccountHideTrans2_12 (Union of qryAccountHideTrans2_12a (Ledger, by specified [actno]),
            ' ** qryAccountHideTrans2_12b (LedgerArchive, by specified [actno])) to tmpAccountHideTrans_02.
9970        Set qdf = .QueryDefs("qryAccountHideTrans2_13")
9980        With qdf.Parameters
9990          ![actno] = gstrAccountNo
10000       End With
10010       qdf.Execute
10020       Set qdf = Nothing

10030       .Close
10040     End With

10050     .assetno = 0&
10060     .curr_id = 0&

10070     .cmdClose.SetFocus
10080     .cmdOK.Enabled = False
10090     .cmdOK.Visible = True
10100     .cmdUnHide.Enabled = False
10110     .cmdUnHide.Visible = False
10120     .pending_trans_lbl.Caption = "Pending Matches"
10130     .cmdCancel.Enabled = False

10140     .shareface_subtot_box1.Visible = False
10150     .shareface_subtot_box2.Visible = False
10160     .icash_subtot_box1.Visible = False
10170     .icash_subtot_box2.Visible = False
10180     .pcash_subtot_box1.Visible = False
10190     .pcash_subtot_box2.Visible = False
10200     .cost_subtot_box1.Visible = False
10210     .cost_subtot_box2.Visible = False
10220     .curr_id_box1.Visible = False
10230     .curr_id_box2.Visible = False

10240     .Requery
10250     .frmAccountHideTrans2_Sub_List.Form.Requery
10260     .frmAccountHideTrans2_Sub_Pick.Form.Requery

10270     .hidden_trans_display = .hidden_trans
10280     .pending_trans = 0&
10290     .shareface_subtot = 0#
10300     .shareface_divint_subtot = 0#
10310     .icash_subtot = 0@
10320     .pcash_subtot = 0@
10330     .cost_subtot = 0@
10340     strTmp01 = "$#,##0.00;($#,##0.00);$0.00;$0.00"
10350     .icash_subtot.Format = strTmp01
10360     .pcash_subtot.Format = strTmp01
10370     .cost_subtot.Format = strTmp01

10380     .unhidden_trans = 0&

10390     TotalsEnable False  ' ** Procedure: Below.
10400     UpdateHiddenTrans  ' ** Procedure: Below.
10410     CheckBoxLock False  ' ** Procedure: Below.
10420     .running_pick_cnt = 0&
10430     DoEvents

10440     If .hidden_trans = 0 Then
10450       If .chkShowHiddenOnly = True Then
10460         .chkShowHiddenOnly = False
10470         DoEvents
10480         chkShowHiddenOnly_AfterUpdate  ' ** Procedure: Below.
10490         DoEvents
10500       End If
10510       If .chkShowHiddenOnly.Enabled = True Then
10520         .chkShowHiddenOnly.Enabled = False
10530       End If
10540     End If

10550     dblTmp02 = lngCurList_Height - (lngSubHeader_Height + lngSubFooter_Height + lngNavBtns_Height)
10560     If (dblTmp02 Mod lngSubRow_Height) = 0 Then
10570       lngCurList_Rows = CLng(dblTmp02 / lngSubRow_Height)
10580     Else
10590       lngCurList_Rows = CLng(Round((dblTmp02 / lngSubRow_Height), 0))
10600     End If
10610     dblTmp02 = lngCurPick_Height - (lngSubHeader_Height + lngSubFooter_Height + lngNavBtns_Height)
10620     If (dblTmp02 Mod lngSubRow_Height) = 0 Then
10630       lngCurPick_Rows = CLng(dblTmp02 / lngSubRow_Height)
10640     Else
10650       lngCurPick_Rows = CLng(Round((dblTmp02 / lngSubRow_Height), 0))
10660     End If

          ' ** Pick starts out 2 rows shorter than design view.
10670     If lngCurPick_Rows > (lngSubPick_Rows - 2&) Or lngCurList_Rows < (lngSubList_Rows + 2&) Then
10680       Do Until lngCurPick_Rows = (lngSubPick_Rows - 2&)
10690         PickSubExpand False  ' ** Procedure: Below
10700         DoEvents
10710         lngCurPick_Rows = lngCurPick_Rows - 1&
10720       Loop
10730     End If

10740     EmptyTemps True  ' ** Procedure: Below.

10750     If .journalno1 > 0& Then
            ' ** Move to 1st journalno in just-hidden group.
10760       .frmAccountHideTrans2_Sub_List.Form.MoveRec 0, .journalno1  ' ** Form Procedure: frmAccountHideTrans2_Sub_List.
10770       DoEvents
10780       .journalno1 = 0&
10790     End If

10800     Beep
10810     DoCmd.Hourglass False

10820   End With

EXITP:
10830   Set qdf = Nothing
10840   Set dbs = Nothing
10850   Exit Sub

ERRH:
10860   DoCmd.Hourglass False
10870   Select Case ERR.Number
        Case Else
10880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10890   End Select
10900   Resume EXITP

End Sub

Private Sub cmdUnHide_KeyDown(KeyCode As Integer, Shift As Integer)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUnHide_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
11010   intShiftDown = (Shift And acShiftMask) > 0
11020   intAltDown = (Shift And acAltMask) > 0
11030   intCtrlDown = (Shift And acCtrlMask) > 0

11040   intRetVal = KeyCode

        ' ** Plain keys.
11050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11060     Select Case intRetVal
          Case vbKeyTab
11070       With Me
11080         intRetVal = 0
11090         .cmdCancel.SetFocus
11100       End With
11110     End Select
11120   End If

        ' ** Shift keys.
11130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11140     Select Case intRetVal
          Case vbKeyTab
11150       With Me
11160         intRetVal = 0
11170         lngRecsCur = .frmAccountHideTrans2_Sub_Pick.Form.RecCnt
11180         If lngRecsCur > 0& Then
11190           .frmAccountHideTrans2_Sub_Pick.SetFocus
11200         Else
11210           .frmAccountHideTrans2_Sub_List.SetFocus
11220         End If
11230       End With
11240     End Select
11250   End If

EXITP:
11260   KeyCode = intRetVal
11270   Exit Sub

ERRH:
11280   intRetVal = 0
11290   Select Case ERR.Number
        Case Else
11300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11310   End Select
11320   Resume EXITP

End Sub

Private Sub cmdCancel_Click()

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngCurList_Height As Long, lngCurPick_Height As Long
        Dim lngCurList_Rows As Long, lngCurPick_Rows As Long
        Dim strFormat As String
        Dim dblTmp01 As Double

11410   With Me

11420     DoCmd.Hourglass True
11430     DoEvents

11440     .shareface_subtot_box1.Visible = False
11450     .shareface_subtot_box2.Visible = False
11460     .icash_subtot_box1.Visible = False
11470     .icash_subtot_box2.Visible = False
11480     .pcash_subtot_box1.Visible = False
11490     .pcash_subtot_box2.Visible = False
11500     .cost_subtot_box1.Visible = False
11510     .cost_subtot_box2.Visible = False
11520     .curr_id_box1.Visible = False
11530     .curr_id_box2.Visible = False
11540     .frmAccountHideTrans2_Sub_Pick.Form.Sort_lbl.Visible = False
11550     .frmAccountHideTrans2_Sub_Pick.Form.Sort_line.Visible = False
11560     .frmAccountHideTrans2_Sub_Pick.Form.journalno_Sort_line1.Visible = False
11570     .frmAccountHideTrans2_Sub_Pick.Form.journalno_Sort_line2.Visible = False

11580     lngCurList_Height = .frmAccountHideTrans2_Sub_List.Height
11590     lngCurPick_Height = .frmAccountHideTrans2_Sub_Pick.Height

11600     Set dbs = CurrentDb
11610     Select Case .cmdUnHide.Visible
          Case True
            ' ** A Cancel when unhiding means all stay hidden.
            ' ** However, if they've already unhidden one or two,
            ' ** the list will no longer be accurate!
            ' ** Empty tmpAccountHideTrans_04.
11620       Set qdf = dbs.QueryDefs("qryAccountHideTrans2_01d")
11630       qdf.Execute
11640       Set qdf = Nothing
11650       DoEvents
            ' ** Append qryAccountHideTrans2_27a (tmpAccountHideTrans_02, linked to qryAccountHideTrans2_25
            ' ** (qryAccountHideTrans2_24 (Union of qryAccountHideTrans2_24a (Ledger, just needed fields),
            ' ** qryAccountHideTrans2_24b (LedgerArchive, just needed fields)), just ledger_HIDDEN = True),
            ' ** just hiddens no longer hidden) to tmpAccountHideTrans_04.
11660       Set qdf = dbs.QueryDefs("qryAccountHideTrans2_27b")
11670       qdf.Execute
11680       Set qdf = Nothing
11690       DoEvents
            ' ** Update qryAccountHideTrans2_27a (tmpAccountHideTrans_02, linked
            ' ** to qryAccountHideTrans2_25 (qryAccountHideTrans2_24 (Union of
            ' ** qryAccountHideTrans2_24a (Ledger, just needed fields),
            ' ** qryAccountHideTrans2_24b (LedgerArchive, just just needed
            ' ** fields)), just ledger_HIDDEN = True), just hiddens no longer
            ' ** hidden), for ledger_HIDDEN = True.
11700       Set qdf = dbs.QueryDefs("qryAccountHideTrans2_27c")
11710       qdf.Execute dbFailOnError
11720       Set qdf = Nothing
11730       DoEvents
            ' ** Empty tmpAccountHideTrans_03.
11740       Set qdf = dbs.QueryDefs("qryAccountHideTrans2_01c")
11750       qdf.Execute
11760       Set qdf = Nothing
11770       DoEvents
11780     Case False
            ' ** Update tmpAccountHideTrans_02, for ledger_HIDDEN = False, via subquery to tmpAccountHideTrans_03.
11790       Set qdf = dbs.QueryDefs("qryAccountHideTrans2_17")
11800       qdf.Execute
11810       Set qdf = Nothing
11820       DoEvents
            ' ** Empty tmpAccountHideTrans_03.
11830       Set qdf = dbs.QueryDefs("qryAccountHideTrans2_01c")
11840       qdf.Execute
11850       Set qdf = Nothing
11860       DoEvents
11870     End Select

11880     dbs.Close
11890     Set dbs = Nothing

11900     .frmAccountHideTrans2_Sub_Pick.Form.Requery
11910     DoEvents

11920     .assetno = 0&
11930     .curr_id = 0&
11940     .journalno1 = 0&

11950     .pending_trans = 0&
11960     .shareface_subtot = 0#
11970     .shareface_divint_subtot = 0#
11980     .icash_subtot = 0@
11990     .pcash_subtot = 0@
12000     .cost_subtot = 0@

12010     .frmAccountHideTrans2_Sub_List.SetFocus
12020     .frmAccountHideTrans2_Sub_List.Form.Refresh
12030     .cmdOK.Enabled = False
12040     .cmdOK.Visible = True
12050     .cmdUnHide.Enabled = False
12060     .cmdUnHide.Visible = False
12070     .pending_trans_lbl.Caption = "Pending Matches"
12080     .cmdCancel.Enabled = False

12090     TotalsEnable False  ' ** Procedure: Below.
12100     UpdateHiddenTrans  ' ** Procedure: Below.
12110     CheckBoxLock False  ' ** Procedure: Below.
12120     .running_pick_cnt = 0&

12130     strFormat = "$#,##0.00;($#,##0.00);$0.00;$0.00"
12140     .icash_subtot.Format = strFormat
12150     .pcash_subtot.Format = strFormat
12160     .cost_subtot.Format = strFormat

12170     dblTmp01 = lngCurList_Height - (lngSubHeader_Height + lngSubFooter_Height + lngNavBtns_Height)
12180     If (dblTmp01 Mod lngSubRow_Height) = 0 Then
12190       lngCurList_Rows = CLng(dblTmp01 / lngSubRow_Height)
12200     Else
12210       lngCurList_Rows = CLng(Round((dblTmp01 / lngSubRow_Height), 0))
12220     End If
12230     dblTmp01 = lngCurPick_Height - (lngSubHeader_Height + lngSubFooter_Height + lngNavBtns_Height)
12240     If (dblTmp01 Mod lngSubRow_Height) = 0 Then
12250       lngCurPick_Rows = CLng(dblTmp01 / lngSubRow_Height)
12260     Else
12270       lngCurPick_Rows = CLng(Round((dblTmp01 / lngSubRow_Height), 0))
12280     End If

          ' ** Pick starts out 2 rows shorter than design view.
12290     If lngCurPick_Rows > (lngSubPick_Rows - 2&) Or lngCurList_Rows < (lngSubList_Rows + 2&) Then
12300       Do Until lngCurPick_Rows = (lngSubPick_Rows - 2&)
12310         PickSubExpand False  ' ** Procedure: Below
12320         DoEvents
12330         lngCurPick_Rows = lngCurPick_Rows - 1&
12340       Loop
12350     End If

12360   End With

12370   DoCmd.Hourglass False

EXITP:
12380   Set qdf = Nothing
12390   Set dbs = Nothing
12400   Exit Sub

ERRH:
12410   DoCmd.Hourglass False
12420   Select Case ERR.Number
        Case Else
12430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12440   End Select
12450   Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
12510   intShiftDown = (Shift And acShiftMask) > 0
12520   intAltDown = (Shift And acAltMask) > 0
12530   intCtrlDown = (Shift And acCtrlMask) > 0

12540   intRetVal = KeyCode

        ' ** Plain keys.
12550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12560     Select Case intRetVal
          Case vbKeyTab
12570       With Me
12580         intRetVal = 0
12590         If .chkIncludeArchive.Enabled = True Then
12600           .chkIncludeArchive.SetFocus
12610         ElseIf .chkIncludeCurrency.Enabled = True Then
12620           .chkIncludeCurrency.SetFocus
12630         ElseIf .chkShowAll.Enabled = True Then
12640           .chkShowAll.SetFocus
12650         ElseIf .chkShowJournalNo.Enabled = True Then
12660           .chkShowJournalNo.SetFocus
12670         ElseIf .chkShowHiddenOnly.Enabled = True Then
12680           .chkShowHiddenOnly.SetFocus
12690         ElseIf .cmdPrintReport.Enabled = True Then
12700           .cmdPrintReport.SetFocus
12710         ElseIf .cmdHidden.Enabled = True Then
12720           .cmdHidden.SetFocus
12730         Else
12740           .cmdClose.SetFocus
12750         End If
12760       End With
12770     End Select
12780   End If

        ' ** Shift keys.
12790   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12800     Select Case intRetVal
          Case vbKeyTab
12810       With Me
12820         intRetVal = 0
12830         If .cmdOK.Visible = True And .cmdOK.Enabled = True Then
12840           .cmdOK.SetFocus
12850         ElseIf .cmdUnHide.Visible = True And .cmdUnHide.Enabled = True Then
12860           .cmdUnHide.SetFocus
12870         Else
12880           Beep
12890         End If
12900       End With
12910     End Select
12920   End If

EXITP:
12930   KeyCode = intRetVal
12940   Exit Sub

ERRH:
12950   intRetVal = 0
12960   Select Case ERR.Number
        Case Else
12970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12980   End Select
12990   Resume EXITP

End Sub

Private Sub chkIncludeArchive_BeforeUpdate(Cancel As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_BeforeUpdate"

13010   With Me
13020     If .archive_trans = 0 Then
            ' ** Should already be disabled!
13030       .Undo
13040       Cancel = -1
13050       .cmdClose.SetFocus
13060       .chkIncludeArchive.Enabled = False
13070       .chkIncludeArchive_lbl.Visible = False
13080       Select Case blnHasArchive
            Case True
13090         .chkIncludeArchive_lbl4.Visible = True
13100         .chkIncludeArchive_lbl4_dim_hi.Visible = True
13110         .chkIncludeArchive_lbl5.Visible = True
13120         .chkIncludeArchive_lbl5_dim_hi.Visible = True
13130       Case False
13140         .chkIncludeArchive_lbl2.Visible = True
13150         .chkIncludeArchive_lbl2_dim_hi.Visible = True
13160         .chkIncludeArchive_lbl3.Visible = True
13170         .chkIncludeArchive_lbl3_dim_hi.Visible = True
13180       End Select
13190     End If
13200   End With

EXITP:
13210   Exit Sub

ERRH:
13220   Select Case ERR.Number
        Case Else
13230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13240   End Select
13250   Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

        Dim strFilter As String

13310   With Me
13320     If .archive_trans > 0 Then
13330       Select Case .chkIncludeArchive
            Case True
13340         .chkIncludeArchive_lbl.FontBold = True
13350         If .chkShowAll = False Then
13360           .chkShowAll = True
13370           .chkShowAll_lbl.FontBold = True
13380         End If
13390         strFilter = vbNullString
13400         .frmAccountHideTrans2_Sub_List.Form.Filter = strFilter
13410         .frmAccountHideTrans2_Sub_List.Form.FilterOn = False
13420         .LastStmt_Trans_lbl.Visible = False  ' ** Just to be sure.
13430         .archive_trans_lbl.Visible = True
13440         .archive_trans_lbl_box.Visible = True
13450       Case False
13460         .chkIncludeArchive_lbl.FontBold = False
13470         strFilter = "[aht2_source] = 'L'"
13480         .frmAccountHideTrans2_Sub_List.Form.Filter = strFilter
13490         .frmAccountHideTrans2_Sub_List.Form.FilterOn = True
13500         .archive_trans_lbl.Visible = False
13510         .archive_trans_lbl_box.Visible = False
13520       End Select
13530       .frmAccountHideTrans2_Sub_List.Form.Requery
13540       DoEvents
13550       .frmAccountHideTrans2_Sub_List.SetFocus
13560     End If
13570   End With

EXITP:
13580   Exit Sub

ERRH:
13590   Select Case ERR.Number
        Case Else
13600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13610   End Select
13620   Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
13710   intShiftDown = (Shift And acShiftMask) > 0
13720   intAltDown = (Shift And acAltMask) > 0
13730   intCtrlDown = (Shift And acCtrlMask) > 0

13740   intRetVal = KeyCode

        ' ** Plain keys.
13750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13760     Select Case intRetVal
          Case vbKeyTab
13770       With Me
13780         intRetVal = 0
13790         If .chkIncludeCurrency.Enabled = True Then
13800           .chkIncludeCurrency.SetFocus
13810         ElseIf .chkShowAll.Enabled = True Then
13820           .chkShowAll.SetFocus
13830         ElseIf .chkShowJournalNo.Enabled = True Then
13840           .chkShowJournalNo.SetFocus
13850         ElseIf .chkShowHiddenOnly.Enabled = True Then
13860           .chkShowHiddenOnly.SetFocus
13870         ElseIf .cmdPrintReport.Enabled = True Then
13880           .cmdPrintReport.SetFocus
13890         ElseIf .cmdHidden.Enabled = True Then
13900           .cmdHidden.SetFocus
13910         Else
13920           .cmdClose.SetFocus
13930         End If
13940       End With
13950     End Select
13960   End If

        ' ** Shift keys.
13970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13980     Select Case intRetVal
          Case vbKeyTab
13990       With Me
14000         intRetVal = 0
14010         If .cmdCancel.Enabled = True Then
14020           .cmdCancel.SetFocus
14030         Else
14040           .frmAccountHideTrans2_Sub_List.SetFocus
14050         End If
14060       End With
14070     End Select
14080   End If

EXITP:
14090   KeyCode = intRetVal
14100   Exit Sub

ERRH:
14110   intRetVal = 0
14120   Select Case ERR.Number
        Case Else
14130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14140   End Select
14150   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

14210   With Me
14220     DoCmd.Hourglass True
14230     DoEvents
14240     Select Case .chkIncludeCurrency
          Case True
14250       .chkIncludeCurrency_lbl.FontBold = True
14260       .frmAccountHideTrans2_Sub_List.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountHideTrans2_Sub_List.
14270       .frmAccountHideTrans2_Sub_Pick.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountHideTrans2_Sub_Pick.
14280       IncludeCurrency True  ' ** Procedure: Below.
14290     Case False
14300       .chkIncludeCurrency_lbl.FontBold = False
14310       .frmAccountHideTrans2_Sub_List.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountHideTrans2_Sub_List.
14320       .frmAccountHideTrans2_Sub_Pick.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountHideTrans2_Sub_Pick.
14330       IncludeCurrency False  ' ** Procedure: Below.
14340     End Select
14350     .frmAccountHideTrans2_Sub_List.SetFocus
14360     DoCmd.Hourglass False
14370   End With

EXITP:
14380   Exit Sub

ERRH:
14390   DoCmd.Hourglass False
14400   Select Case ERR.Number
        Case Else
14410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14420   End Select
14430   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
14510   intShiftDown = (Shift And acShiftMask) > 0
14520   intAltDown = (Shift And acAltMask) > 0
14530   intCtrlDown = (Shift And acCtrlMask) > 0

14540   intRetVal = KeyCode

        ' ** Plain keys.
14550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14560     Select Case intRetVal
          Case vbKeyTab
14570       With Me
14580         intRetVal = 0
14590         If .chkShowAll.Enabled = True Then
14600           .chkShowAll.SetFocus
14610         ElseIf .chkShowJournalNo.Enabled = True Then
14620           .chkShowJournalNo.SetFocus
14630         ElseIf .chkShowHiddenOnly.Enabled = True Then
14640           .chkShowHiddenOnly.SetFocus
14650         ElseIf .cmdPrintReport.Enabled = True Then
14660           .cmdPrintReport.SetFocus
14670         ElseIf .cmdHidden.Enabled = True Then
14680           .cmdHidden.SetFocus
14690         Else
14700           .cmdClose.SetFocus
14710         End If
14720       End With
14730     End Select
14740   End If

        ' ** Shift keys.
14750   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14760     Select Case intRetVal
          Case vbKeyTab
14770       With Me
14780         intRetVal = 0
14790         If .chkIncludeArchive.Enabled = True Then
14800           .chkIncludeArchive.SetFocus
14810         ElseIf .cmdCancel.Enabled = True Then
14820           .cmdCancel.SetFocus
14830         Else
14840           .frmAccountHideTrans2_Sub_List.SetFocus
14850         End If
14860       End With
14870     End Select
14880   End If

EXITP:
14890   KeyCode = intRetVal
14900   Exit Sub

ERRH:
14910   intRetVal = 0
14920   Select Case ERR.Number
        Case Else
14930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14940   End Select
14950   Resume EXITP

End Sub

Private Sub chkShowAll_AfterUpdate()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAll_AfterUpdate"

        Dim strFilter As String

15010   With Me
15020     If .chkShowHiddenOnly = True Then
            ' ** For simplicity, show all needs to override hidden only.
15030       .chkShowHiddenOnly = False
15040       .chkShowHiddenOnly_lbl.FontBold = False
15050     End If
15060     Select Case .chkShowAll
          Case True
15070       .chkShowAll_lbl.FontBold = True
15080       strFilter = "[aht2_source] = 'L'"
15090       .frmAccountHideTrans2_Sub_List.Form.Filter = strFilter
15100       .frmAccountHideTrans2_Sub_List.Form.FilterOn = True
15110       .LastStmt_Trans_lbl.Visible = False
15120       If .archive_trans > 0 Then
15130         .chkIncludeArchive.Enabled = True
15140       Else
15150         .chkIncludeArchive.Enabled = False
15160       End If
15170     Case False
15180       .chkShowAll_lbl.FontBold = False
15190       If .chkIncludeArchive = True Then
15200         .chkIncludeArchive = False
15210         .chkIncludeArchive_lbl.FontBold = False
15220         .archive_trans_lbl.Visible = False
15230         .archive_trans_lbl_box.Visible = False
15240       End If
15250       strFilter = "[transdate] > #" & Format(.balance_date_max, "mm/dd/yyyy") & "#"
15260       strFilter = strFilter & " And [aht2_source] = 'L'"
15270       .frmAccountHideTrans2_Sub_List.Form.Filter = strFilter
15280       .frmAccountHideTrans2_Sub_List.Form.FilterOn = True
15290       .LastStmt_Trans_lbl.Visible = True
15300       If .archive_trans > 0 Then
15310         .chkIncludeArchive.Enabled = False
15320       End If
15330     End Select
15340     .frmAccountHideTrans2_Sub_List.Form.Requery
15350     DoEvents
15360     .frmAccountHideTrans2_Sub_List.SetFocus
15370   End With

EXITP:
15380   Exit Sub

ERRH:
15390   Select Case ERR.Number
        Case Else
15400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15410   End Select
15420   Resume EXITP

End Sub

Private Sub chkShowAll_KeyDown(KeyCode As Integer, Shift As Integer)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
15510   intShiftDown = (Shift And acShiftMask) > 0
15520   intAltDown = (Shift And acAltMask) > 0
15530   intCtrlDown = (Shift And acCtrlMask) > 0

15540   intRetVal = KeyCode

        ' ** Plain keys.
15550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15560     Select Case intRetVal
          Case vbKeyTab
15570       With Me
15580         intRetVal = 0
15590         If .chkShowJournalNo.Enabled = True Then
15600           .chkShowJournalNo.SetFocus
15610         ElseIf .chkShowHiddenOnly.Enabled = True Then
15620           .chkShowHiddenOnly.SetFocus
15630         ElseIf .cmdPrintReport.Enabled = True Then
15640           .cmdPrintReport.SetFocus
15650         ElseIf .cmdHidden.Enabled = True Then
15660           .cmdHidden.SetFocus
15670         Else
15680           .cmdClose.SetFocus
15690         End If
15700       End With
15710     End Select
15720   End If

        ' ** Shift keys.
15730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15740     Select Case intRetVal
          Case vbKeyTab
15750       With Me
15760         intRetVal = 0
15770         If .chkIncludeCurrency.Enabled = True Then
15780           .chkIncludeCurrency.SetFocus
15790         ElseIf .chkIncludeArchive.Enabled = True Then
15800           .chkIncludeArchive.SetFocus
15810         ElseIf .cmdCancel.Enabled = True Then
15820           .cmdCancel.SetFocus
15830         Else
15840           .frmAccountHideTrans2_Sub_List.SetFocus
15850         End If
15860       End With
15870     End Select
15880   End If

EXITP:
15890   KeyCode = intRetVal
15900   Exit Sub

ERRH:
15910   intRetVal = 0
15920   Select Case ERR.Number
        Case Else
15930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15940   End Select
15950   Resume EXITP

End Sub

Private Sub chkShowJournalNo_AfterUpdate()

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowJournalNo_AfterUpdate"

16010   With Me
16020     Select Case .chkShowJournalNo
          Case True
16030       .chkShowJournalNo_lbl.FontBold = True
16040       .frmAccountHideTrans2_Sub_List.Form.ShowJournalNo_Sub True  ' ** Form Procedure: frmAccountHideTrans2_Sub_List.
16050       .frmAccountHideTrans2_Sub_Pick.Form.ShowJournalNo_Sub True  ' ** Form Procedure: frmAccountHideTrans2_Sub_Pick.
16060     Case False
16070       .chkShowJournalNo_lbl.FontBold = False
16080       .frmAccountHideTrans2_Sub_List.Form.ShowJournalNo_Sub False  ' ** Form Procedure: frmAccountHideTrans2_Sub_List.
16090       .frmAccountHideTrans2_Sub_Pick.Form.ShowJournalNo_Sub False  ' ** Form Procedure: frmAccountHideTrans2_Sub_Pick.
16100     End Select
16110     .frmAccountHideTrans2_Sub_List.SetFocus
16120   End With

EXITP:
16130   Exit Sub

ERRH:
16140   Select Case ERR.Number
        Case Else
16150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16160   End Select
16170   Resume EXITP

End Sub

Private Sub chkShowJournalNo_KeyDown(KeyCode As Integer, Shift As Integer)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowJournalNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
16210   intShiftDown = (Shift And acShiftMask) > 0
16220   intAltDown = (Shift And acAltMask) > 0
16230   intCtrlDown = (Shift And acCtrlMask) > 0

16240   intRetVal = KeyCode

        ' ** Plain keys.
16250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16260     Select Case intRetVal
          Case vbKeyTab
16270       With Me
16280         intRetVal = 0
16290         If .chkShowHiddenOnly.Enabled = True Then
16300           .chkShowHiddenOnly.SetFocus
16310         ElseIf .cmdPrintReport.Enabled = True Then
16320           .cmdPrintReport.SetFocus
16330         ElseIf .cmdHidden.Enabled = True Then
16340           .cmdHidden.SetFocus
16350         Else
16360           .cmdClose.SetFocus
16370         End If
16380       End With
16390     End Select
16400   End If

        ' ** Shift keys.
16410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16420     Select Case intRetVal
          Case vbKeyTab
16430       With Me
16440         intRetVal = 0
16450         If .chkShowAll.Enabled = True Then
16460           .chkShowAll.SetFocus
16470         ElseIf .chkIncludeCurrency.Enabled = True Then
16480           .chkIncludeCurrency.SetFocus
16490         ElseIf .chkIncludeArchive.Enabled = True Then
16500           .chkIncludeArchive.SetFocus
16510         ElseIf .cmdCancel.Enabled = True Then
16520           .cmdCancel.SetFocus
16530         Else
16540           .frmAccountHideTrans2_Sub_List.SetFocus
16550         End If
16560       End With
16570     End Select
16580   End If

EXITP:
16590   KeyCode = intRetVal
16600   Exit Sub

ERRH:
16610   intRetVal = 0
16620   Select Case ERR.Number
        Case Else
16630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16640   End Select
16650   Resume EXITP

End Sub

Private Sub chkShowHiddenOnly_BeforeUpdate(Cancel As Integer)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowHiddenOnly_BeforeUpdate"

16710   With Me
16720     If .hidden_trans = 0 And .hidden_trans_display = 0 Then
            ' ** Should already be disabled!
16730       chkShowHiddenOnly.Undo
16740       Cancel = -1
16750 On Error Resume Next
16760       .cmdClose.SetFocus
16770       .chkShowHiddenOnly.Enabled = False
16780 On Error GoTo ERRH
16790     End If
16800   End With

EXITP:
16810   Exit Sub

ERRH:
16820   Select Case ERR.Number
        Case Else
16830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16840   End Select
16850   Resume EXITP

End Sub

Private Sub chkShowHiddenOnly_AfterUpdate()

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowHiddenOnly_AfterUpdate"

        Dim strFilter As String

16910   With Me
16920     If .hidden_trans > 0 Or (.hidden_trans = 0 And .hidden_trans_display > 0) Then
16930       Select Case .chkShowHiddenOnly
            Case True
16940         .chkShowHiddenOnly_lbl.FontBold = True
16950         strFilter = "[ledger_HIDDEN] = True"
16960         .frmAccountHideTrans2_Sub_List.Form.Filter = strFilter
16970         .frmAccountHideTrans2_Sub_List.Form.FilterOn = True
16980         .LastStmt_Trans_lbl.Visible = False
16990         .chkShowAll = True
17000         .chkShowAll_lbl.FontBold = True
17010         .chkShowAll_lbl.ForeColor = CLR_GRY5
17020         .chkShowAll_lbl6.Visible = True
17030         .chkShowAll.Locked = True
17040         If .archive_trans > 0 Then
17050           .chkIncludeArchive.Enabled = True
17060           .chkIncludeArchive = True
17070           .chkIncludeArchive_lbl.FontBold = True
17080           .chkIncludeArchive_lbl.ForeColor = CLR_GRY5
17090           .chkIncludeArchive_lbl6.Visible = True
17100           .chkIncludeArchive.Locked = True
17110           .archive_trans_lbl.Visible = True
17120           .archive_trans_lbl_box.Visible = True
17130           .chkShowHiddenOnly_Msg1.Visible = True
17140         End If
17150       Case False
17160         .chkShowHiddenOnly_lbl.FontBold = False
17170         strFilter = vbNullString
17180         .frmAccountHideTrans2_Sub_List.Form.Filter = strFilter
17190         .frmAccountHideTrans2_Sub_List.Form.FilterOn = False
17200         If .archive_trans > 0 Then
17210           .chkIncludeArchive_lbl.ForeColor = CLR_DKGRY4
17220           .chkIncludeArchive_lbl6.Visible = False
17230           .chkIncludeArchive.Locked = False
17240         End If
17250         .chkShowHiddenOnly_Msg1.Visible = False
17260         .chkShowAll_lbl.ForeColor = CLR_DKGRY4
17270         .chkShowAll_lbl6.Visible = False
17280         .chkShowAll.Locked = False
17290       End Select
17300       .frmAccountHideTrans2_Sub_List.Form.Requery
17310       DoEvents
17320       .frmAccountHideTrans2_Sub_List.SetFocus
17330     End If
17340   End With

EXITP:
17350   Exit Sub

ERRH:
17360   Select Case ERR.Number
        Case Else
17370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17380   End Select
17390   Resume EXITP

End Sub

Private Sub chkShowHiddenOnly_KeyDown(KeyCode As Integer, Shift As Integer)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowHiddenOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
17410   intShiftDown = (Shift And acShiftMask) > 0
17420   intAltDown = (Shift And acAltMask) > 0
17430   intCtrlDown = (Shift And acCtrlMask) > 0

17440   intRetVal = KeyCode

        ' ** Plain keys.
17450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17460     Select Case intRetVal
          Case vbKeyTab
17470       With Me
17480         intRetVal = 0
17490         If .cmdPrintReport.Enabled = True Then
17500           .cmdPrintReport.SetFocus
17510         ElseIf .cmdHidden.Enabled = True Then
17520           .cmdHidden.SetFocus
17530         Else
17540           .cmdClose.SetFocus
17550         End If
17560       End With
17570     End Select
17580   End If

        ' ** Shift keys.
17590   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17600     Select Case intRetVal
          Case vbKeyTab
17610       With Me
17620         intRetVal = 0
17630         If .chkShowJournalNo.Enabled = True Then
17640           .chkShowJournalNo.SetFocus
17650         ElseIf .chkShowAll.Enabled = True Then
17660           .chkShowAll.SetFocus
17670         ElseIf .chkIncludeCurrency.Enabled = True Then
17680           .chkIncludeCurrency.SetFocus
17690         ElseIf .chkIncludeArchive.Enabled = True Then
17700           .chkIncludeArchive.SetFocus
17710         ElseIf .cmdCancel.Enabled = True Then
17720           .cmdCancel.SetFocus
17730         Else
17740           .frmAccountHideTrans2_Sub_List.SetFocus
17750         End If
17760       End With
17770     End Select
17780   End If

EXITP:
17790   KeyCode = intRetVal
17800   Exit Sub

ERRH:
17810   intRetVal = 0
17820   Select Case ERR.Number
        Case Else
17830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17840   End Select
17850   Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

17910   With Me
17920     blnPrintReport_Focus = True
17930     .cmdPrintReport_raised_semifocus_dots_img.Visible = True
17940     .cmdPrintReport_raised_img.Visible = False
17950     .cmdPrintReport_raised_focus_img.Visible = False
17960     .cmdPrintReport_raised_focus_dots_img.Visible = False
17970     .cmdPrintReport_sunken_focus_dots_img.Visible = False
17980     .cmdPrintReport_raised_img_dis.Visible = False
17990   End With

EXITP:
18000   Exit Sub

ERRH:
18010   Select Case ERR.Number
        Case Else
18020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18030   End Select
18040   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

18110   With Me
18120     blnPrintReport_MouseDown = True
18130     .cmdPrintReport_sunken_focus_dots_img.Visible = True
18140     .cmdPrintReport_raised_img.Visible = False
18150     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
18160     .cmdPrintReport_raised_focus_img.Visible = False
18170     .cmdPrintReport_raised_focus_dots_img.Visible = False
18180     .cmdPrintReport_raised_img_dis.Visible = False
18190   End With

EXITP:
18200   Exit Sub

ERRH:
18210   Select Case ERR.Number
        Case Else
18220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18230   End Select
18240   Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strSortNow As String, strSortOrig As String
        Dim intMode As Integer

18310   With Me
18320     If .hidden_trans > 0& Then
18330       strSortNow = .frmAccountHideTrans2_Sub_List.Form.SortNow_Get  ' ** Form Function: frmAccountHideTrans2_Sub_List.
18340       strSortOrig = .frmAccountHideTrans2_Sub_List.Form.SortOrig_Get  ' ** Form Function: frmAccountHideTrans2_Sub_List.
18350       If strSortNow = strSortOrig Then
18360         strDocName = "rptTransaction_Hidden"
18370       Else
18380         strDocName = "rptTransaction_Hidden_Sort"
18390       End If
18400       If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
18410         intMode = acViewPreview
18420       Else
18430         intMode = acViewNormal
18440       End If
            '##GTR_Ref: rptTransaction_Hidden
18450       DoCmd.OpenReport strDocName, intMode, , , , .accountno
18460       If intMode = acViewPreview Then
18470         DoCmd.Maximize
18480         DoCmd.RunCommand acCmdFitToWindow
18490         Forms(THIS_NAME).Visible = False
18500       End If
18510     Else
18520       Beep
18530     End If
18540   End With

EXITP:
18550   Exit Sub

ERRH:
18560   Select Case ERR.Number
        Case Else
18570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18580   End Select
18590   Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
18610   intShiftDown = (Shift And acShiftMask) > 0
18620   intAltDown = (Shift And acAltMask) > 0
18630   intCtrlDown = (Shift And acCtrlMask) > 0

18640   intRetVal = KeyCode

        ' ** Plain keys.
18650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18660     Select Case intRetVal
          Case vbKeyTab
18670       With Me
18680         intRetVal = 0
18690         If .cmdHidden.Enabled = True Then
18700           .cmdHidden.SetFocus
18710         Else
18720           .cmdClose.SetFocus
18730         End If
18740       End With
18750     End Select
18760   End If

        ' ** Shift keys.
18770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18780     Select Case intRetVal
          Case vbKeyTab
18790       With Me
18800         intRetVal = 0
18810         If .chkShowHiddenOnly.Enabled = True Then
18820           .chkShowHiddenOnly.SetFocus
18830         ElseIf .chkShowJournalNo.Enabled = True Then
18840           .chkShowJournalNo.SetFocus
18850         ElseIf .chkShowAll.Enabled = True Then
18860           .chkShowAll.SetFocus
18870         ElseIf .chkIncludeCurrency.Enabled = True Then
18880           .chkIncludeCurrency.SetFocus
18890         ElseIf .chkIncludeArchive.Enabled = True Then
18900           .chkIncludeArchive.SetFocus
18910         ElseIf .cmdCancel.Enabled = True Then
18920           .cmdCancel.SetFocus
18930         ElseIf .cmdOK.Visible = True And .cmdOK.Enabled = True Then
18940           .cmdOK.SetFocus
18950         ElseIf .cmdUnHide.Visible = True And .cmdUnHide.Enabled = True Then
18960           .cmdUnHide.SetFocus
18970         Else
18980           .frmAccountHideTrans2_Sub_List.SetFocus
18990           .frmAccountHideTrans2_Sub_List.Form.ledger_HIDDEN.SetFocus
19000         End If
19010       End With
19020     End Select
19030   End If

EXITP:
19040   KeyCode = intRetVal
19050   Exit Sub

ERRH:
19060   intRetVal = 0
19070   Select Case ERR.Number
        Case Else
19080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19090   End Select
19100   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

19210   With Me
19220     If blnPrintReport_MouseDown = False Then
19230       Select Case blnPrintReport_Focus
            Case True
19240         .cmdPrintReport_raised_focus_dots_img.Visible = True
19250         .cmdPrintReport_raised_focus_img.Visible = False
19260       Case False
19270         .cmdPrintReport_raised_focus_img.Visible = True
19280         .cmdPrintReport_raised_focus_dots_img.Visible = False
19290       End Select
19300       .cmdPrintReport_raised_img.Visible = False
19310       .cmdPrintReport_raised_semifocus_dots_img.Visible = False
19320       .cmdPrintReport_sunken_focus_dots_img.Visible = False
19330       .cmdPrintReport_raised_img_dis.Visible = False
19340     End If
19350   End With

EXITP:
19360   Exit Sub

ERRH:
19370   Select Case ERR.Number
        Case Else
19380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19390   End Select
19400   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

19510   With Me
19520     .cmdPrintReport_raised_focus_dots_img.Visible = True
19530     .cmdPrintReport_raised_img.Visible = False
19540     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
19550     .cmdPrintReport_raised_focus_img.Visible = False
19560     .cmdPrintReport_sunken_focus_dots_img.Visible = False
19570     .cmdPrintReport_raised_img_dis.Visible = False
19580     blnPrintReport_MouseDown = False
19590   End With

EXITP:
19600   Exit Sub

ERRH:
19610   Select Case ERR.Number
        Case Else
19620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19630   End Select
19640   Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

19710   With Me
19720     .cmdPrintReport_raised_img.Visible = True
19730     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
19740     .cmdPrintReport_raised_focus_img.Visible = False
19750     .cmdPrintReport_raised_focus_dots_img.Visible = False
19760     .cmdPrintReport_sunken_focus_dots_img.Visible = False
19770     .cmdPrintReport_raised_img_dis.Visible = False
19780     blnPrintReport_Focus = False
19790   End With

EXITP:
19800   Exit Sub

ERRH:
19810   Select Case ERR.Number
        Case Else
19820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19830   End Select
19840   Resume EXITP

End Sub

Private Sub cmdHidden_Click()

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHidden_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strDocName As String
        Dim lngRecs As Long
        Dim blnContinue As Boolean, blnRenumGrps As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long

19910   blnContinue = True

19920   With Me

19930     DoCmd.Hourglass True
19940     DoEvents

19950     gstrAccountNo = .accountno

19960     lngHids = .hidden_trans
19970     If lngHids > 0& Then
19980       Set dbs = CurrentDb
19990       With dbs
              ' ** Balance, grouped, by accountno, with Max(balance_date), by specified [actno].
20000         Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_02")
20010         With qdf.Parameters
20020           ![actno] = gstrAccountNo
20030         End With
20040         Set rst1 = qdf.OpenRecordset
20050         With rst1
20060           If .BOF = True And .EOF = True Then
20070             blnContinue = False
20080           Else
20090             .MoveFirst
20100             gdatEndDate = ![balance_date]
20110           End If
20120           .Close
20130         End With  ' ** rst1.
20140         Set rst1 = Nothing
20150         Set qdf = Nothing
20160       End With  ' ** dbs.
20170     Else
20180       blnContinue = False
20190     End If  ' ** hidden_trans.

20200     DoEvents

20210     If blnContinue = True Then

            ' ** qryAccountHideTrans2_Hidden_05 (tblLedgerHidden, linked to qryAccountHideTrans2_Hidden_04
            ' ** (Union of qryAccountHideTrans2_Hidden_04a (Ledger, linked to MasterAsset, just
            ' ** ledger_HIDDEN = True), qryAccountHideTrans2_Hidden_04b (LedgerArchive, linked to MasterAsset,
            ' ** just ledger_HIDDEN = True)), with totdesc), with shareface_0, shareface_1, shareface_2,
            ' ** shareface_3, shareface_4, by specified FormRef('PriorPeriod').
20220       varTmp00 = DCount("*", "qryAccountHideTrans2_Hidden_06")
20230       If IsNull(varTmp00) = False Then
20240         If lngHids > varTmp00 Then
20250           Hide_AddlMatch gstrAccountNo, False, Me, 0#  ' ** Module Procedure: modHideTransactions1.
20260         End If
20270       End If

20280       With dbs

              ' ** Empty tblMark.
20290         Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_01a")
20300         qdf.Execute
20310         Set qdf = Nothing

20320         Set rst2 = .OpenRecordset("tblMark", dbOpenDynaset, dbConsistent)
20330         blnRenumGrps = False

              ' ** qryAccountHideTrans2_Hidden_06 (qryAccountHideTrans2_Hidden_05 (tblLedgerHidden,
              ' ** linked to qryAccountHideTrans2_Hidden_04 (Union of qryAccountHideTrans2_Hidden_04a
              ' ** (Ledger, linked to MasterAsset, just ledger_HIDDEN = True), qryAccountHideTrans2_Hidden_04b
              ' ** (LedgerArchive, linked to MasterAsset, just ledger_HIDDEN = True)), with totdesc),
              ' ** with shareface_0, shareface_1, shareface_2, shareface_3, shareface_4, by specified
              ' ** FormRef('PriorPeriod')), grouped by ledghid_grpnum.
20340         Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_07")
20350         Set rst1 = qdf.OpenRecordset
20360         With rst1
20370           If .BOF = True And .EOF = True Then
                  ' ** Shouldn't be here then!
20380             blnContinue = False
20390           Else
20400             .MoveLast
20410             lngRecs = .RecordCount
20420             .MoveFirst
20430             For lngX = 1& To lngRecs
20440               rst2.AddNew
20450               rst2![unique_id] = lngX
20460               rst2![mark] = False
20470               rst2![Value] = ![ledghid_grpnum]
20480               rst2.Update
20490               If lngX < lngRecs Then .MoveNext
20500             Next
20510             rst2.MoveFirst
20520             rst2.MoveLast
20530             If rst2![unique_id] <> CLng(rst2![Value]) Then
                    ' ** Simple way to see if they need renumbering.
20540               blnRenumGrps = True
20550             Else
                    ' ** More thorough check.
20560               .MoveFirst
20570               lngY = 0&
20580               For lngX = 1& To lngRecs
20590                 lngY = lngY + 1&
20600                 If lngY <> ![ledghid_grpnum] Then
20610                   blnRenumGrps = True
20620                   Exit For
20630                 End If
20640                 If lngX < lngRecs Then .MoveNext
20650               Next
20660             End If
20670           End If
20680           .Close
20690         End With  ' ** rst1.
20700         rst2.Close
20710         Set rst1 = Nothing
20720         Set rst2 = Nothing
20730         Set qdf = Nothing

20740         .Close
20750       End With  ' ** dbs.
20760       Set dbs = Nothing
20770       DoEvents

20780       If blnContinue = True Then

20790         If blnRenumGrps = True Then
20800           RenumGrps  ' ** Procedure: Below.
20810         End If
20820         DoEvents

20830         gblnSetFocus = True
20840         strDocName = "frmAccountHideTrans2_Hidden"
20850         gstrFormQuerySpec = strDocName
20860         If FormRef("AcctNum") = gstrAccountNo Then
20870           varTmp00 = DCount("*", "qryAccountHideTrans2_Hidden_03")
20880           Select Case IsNull(varTmp00)
                Case True
20890             Beep
20900             DoCmd.Hourglass False
20910           Case False
20920             If varTmp00 > 0 Then
20930               DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & gstrAccountNo & "~" & CStr(lngHids) & "~" & Format(gdatEndDate, "mm/dd/yyyy")
                    ' ** THIS_NAME ~ .accountno ~ .hidden_trans ~ gdatEndDate
20940               If gblnGoToReport = True Then
20950                 Forms(strDocName).TimerInterval = 50&
20960                 blnGoingToReport = False
20970                 blnGTR_Emblem = False
20980               End If
20990             Else
21000               Beep
21010               DoCmd.Hourglass False
21020             End If
21030           End Select
21040         Else
21050           Beep
21060           DoCmd.Hourglass False
21070         End If
21080       Else
21090         Beep
21100         DoCmd.Hourglass False
21110       End If  ' ** blnContinue.

21120     Else
21130       blnContinue = False
21140       dbs.Close
21150       DoCmd.Hourglass False
21160       Beep
21170       MsgBox "There are no hidden transactions for this account.", vbInformation + vbOKOnly, "Nothing To Do"
21180       blnGoingToReport = False  ' ** Shouldn't be here anyway if that's the case.
21190       gblnGoToReport = False
21200     End If  ' ** blnContinue.

21210   End With

EXITP:
21220   Set rst1 = Nothing
21230   Set rst2 = Nothing
21240   Set qdf = Nothing
21250   Set dbs = Nothing
21260   Exit Sub

ERRH:
21270   DoCmd.Hourglass False
21280   Select Case ERR.Number
        Case Else
21290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21300   End Select
21310   Resume EXITP

End Sub

Private Sub cmdHidden_KeyDown(KeyCode As Integer, Shift As Integer)

21400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHidden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
21410   intShiftDown = (Shift And acShiftMask) > 0
21420   intAltDown = (Shift And acAltMask) > 0
21430   intCtrlDown = (Shift And acCtrlMask) > 0

21440   intRetVal = KeyCode

        ' ** Plain keys.
21450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21460     Select Case intRetVal
          Case vbKeyTab
21470       With Me
21480         intRetVal = 0
21490         .cmdClose.SetFocus
21500       End With
21510     End Select
21520   End If

        ' ** Shift keys.
21530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21540     Select Case intRetVal
          Case vbKeyTab
21550       With Me
21560         intRetVal = 0
21570         If .cmdPrintReport.Enabled = True Then
21580           .cmdPrintReport.SetFocus
21590         ElseIf .chkShowHiddenOnly.Enabled = True Then
21600           .chkShowHiddenOnly.SetFocus
21610         ElseIf .chkShowJournalNo.Enabled = True Then
21620           .chkShowJournalNo.SetFocus
21630         ElseIf .chkShowAll.Enabled = True Then
21640           .chkShowAll.SetFocus
21650         ElseIf .chkIncludeCurrency.Enabled = True Then
21660           .chkIncludeCurrency.SetFocus
21670         ElseIf .chkIncludeArchive.Enabled = True Then
21680           .chkIncludeArchive.SetFocus
21690         ElseIf .cmdCancel.Enabled = True Then
21700           .cmdCancel.SetFocus
21710         Else
21720           .frmAccountHideTrans2_Sub_List.SetFocus
21730         End If
21740       End With
21750     End Select
21760   End If

EXITP:
21770   KeyCode = intRetVal
21780   Exit Sub

ERRH:
21790   intRetVal = 0
21800   Select Case ERR.Number
        Case Else
21810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21820   End Select
21830   Resume EXITP

End Sub

Public Sub PickSubExpand(blnAddRow As Boolean)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "PickSubExpand"

        Dim lngCurList_Height As Long, lngCurPick_Height As Long
        Dim lngTmp01 As Long

21910   With Me
21920     lngCurList_Height = .frmAccountHideTrans2_Sub_List.Height
21930     lngCurPick_Height = .frmAccountHideTrans2_Sub_Pick.Height
21940     Select Case blnAddRow
          Case True
21950       .frmAccountHideTrans2_Sub_Pick.Top = .frmAccountHideTrans2_Sub_Pick.Top - lngSubRow_Height
21960       .frmAccountHideTrans2_Sub_Pick_box.Top = .frmAccountHideTrans2_Sub_Pick_box.Top - lngSubRow_Height
21970       .frmAccountHideTrans2_Sub_Pick.Height = .frmAccountHideTrans2_Sub_Pick.Height + lngSubRow_Height
21980       .frmAccountHideTrans2_Sub_Pick_box.Height = .frmAccountHideTrans2_Sub_Pick_box.Height + lngSubRow_Height
21990       .frmAccountHideTrans2_Sub_List.Height = .frmAccountHideTrans2_Sub_List.Height - lngSubRow_Height
22000       .frmAccountHideTrans2_Sub_List_box.Height = .frmAccountHideTrans2_Sub_List_box.Height - lngSubRow_Height
22010       .Nav_1_box01.Top = .Nav_1_box01.Top - lngSubRow_Height
22020       .Nav_1_hline01.Top = .Nav_1_hline01.Top - lngSubRow_Height
22030       .Nav_1_hline02.Top = .Nav_1_hline02.Top - lngSubRow_Height
22040       .Nav_1_vline01.Top = .Nav_1_vline01.Top - lngSubRow_Height
22050       .Nav_1_vline02.Top = .Nav_1_vline02.Top - lngSubRow_Height
22060       .Nav_1_vline03.Top = .Nav_1_vline03.Top - lngSubRow_Height
22070       .Nav_1_vline04.Top = .Nav_1_vline04.Top - lngSubRow_Height
22080       .hidden_trans_display.Top = .hidden_trans_display.Top - lngSubRow_Height
22090       .hidden_trans_display_lbl.Top = .hidden_trans_display_lbl.Top - lngSubRow_Height
22100     Case False
22110       .frmAccountHideTrans2_Sub_Pick.Height = .frmAccountHideTrans2_Sub_Pick.Height - lngSubRow_Height
22120       .frmAccountHideTrans2_Sub_Pick_box.Height = .frmAccountHideTrans2_Sub_Pick_box.Height - lngSubRow_Height
22130       .frmAccountHideTrans2_Sub_Pick.Top = .frmAccountHideTrans2_Sub_Pick.Top + lngSubRow_Height
22140       .frmAccountHideTrans2_Sub_Pick_box.Top = .frmAccountHideTrans2_Sub_Pick_box.Top + lngSubRow_Height
22150       .frmAccountHideTrans2_Sub_List.Height = .frmAccountHideTrans2_Sub_List.Height + lngSubRow_Height
22160       .frmAccountHideTrans2_Sub_List_box.Height = .frmAccountHideTrans2_Sub_List_box.Height + lngSubRow_Height
22170       .Nav_1_box01.Top = .Nav_1_box01.Top + lngSubRow_Height
22180       .Nav_1_hline01.Top = .Nav_1_hline01.Top + lngSubRow_Height
22190       .Nav_1_hline02.Top = .Nav_1_hline02.Top + lngSubRow_Height
22200       .Nav_1_vline01.Top = .Nav_1_vline01.Top + lngSubRow_Height
22210       .Nav_1_vline02.Top = .Nav_1_vline02.Top + lngSubRow_Height
22220       .Nav_1_vline03.Top = .Nav_1_vline03.Top + lngSubRow_Height
22230       .Nav_1_vline04.Top = .Nav_1_vline04.Top + lngSubRow_Height
22240       .hidden_trans_display.Top = .hidden_trans_display.Top + lngSubRow_Height
22250       .hidden_trans_display_lbl.Top = .hidden_trans_display_lbl.Top + lngSubRow_Height
22260     End Select
22270     lngTmp01 = ((.hidden_trans_display.Top + .hidden_trans_display.Height) + (2& * lngTpp))
22280     .Detail_hline01.Top = lngTmp01
22290     .Detail_hline02.Top = (.Detail_hline01.Top + lngTpp)
22300     .Detail_vline01.Top = .Detail_hline01.Top
22310     .Detail_vline02.Top = .Detail_hline01.Top
22320     .ShortcutMenu_lbl.Top = (.frmAccountHideTrans2_Sub_List.Top + .frmAccountHideTrans2_Sub_List.Height) + lngShortcutLbl_Offset
22330     .ShortcutMenu_up_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
22340     .ShortcutMenu_down_arrow_lbl.Top = (.ShortcutMenu_lbl.Top - lngTpp)
22350   End With

EXITP:
22360   Exit Sub

ERRH:
22370   Select Case ERR.Number
        Case Else
22380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22390   End Select
22400   Resume EXITP

End Sub

Public Sub TotalsEnable(blnEnable As Boolean)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "TotalsEnable"

22510   With Me
22520     Select Case blnEnable
          Case True
22530       .pending_trans.Locked = True  ' ** Doesn't look disabled.
22540       .pending_trans.BorderColor = CLR_LTBLU2
22550       .pending_trans.BackStyle = acBackStyleNormal
22560       .shareface_subtot.BackStyle = acBackStyleNormal
22570       .shareface_divint_subtot.Locked = True
22580       .shareface_divint_subtot.BorderColor = CLR_LTBLU2
22590       .shareface_divint_subtot.BackStyle = acBackStyleNormal
22600       If .cnt_asset = 0 Then
22610         .shareface_subtot.Locked = False
22620         .shareface_subtot.BorderColor = WIN_CLR_DISR
22630         .shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
22640         .shareface_subtot_lbl_dim_hi.Visible = True
22650       Else
22660         .shareface_subtot.Locked = True
22670         .shareface_subtot.BorderColor = CLR_LTBLU2
22680         .shareface_subtot_lbl.ForeColor = CLR_DKGRY3
22690         .shareface_subtot_lbl_dim_hi.Visible = False
22700       End If
22710       .icash_subtot.Locked = True
22720       .icash_subtot.BorderColor = CLR_LTBLU2
22730       .icash_subtot.BackStyle = acBackStyleNormal
22740       .icash_subtot_lbl.ForeColor = CLR_DKGRY3
22750       .icash_subtot_lbl_dim_hi.Visible = False
22760       .pcash_subtot.Locked = True
22770       .pcash_subtot.BorderColor = CLR_LTBLU2
22780       .pcash_subtot.BackStyle = acBackStyleNormal
22790       .pcash_subtot_lbl.ForeColor = CLR_DKGRY3
22800       .pcash_subtot_lbl_dim_hi.Visible = False
22810       .cost_subtot.Locked = True
22820       .cost_subtot.BorderColor = CLR_LTBLU2
22830       .cost_subtot.BackStyle = acBackStyleNormal
22840       .cost_subtot_lbl.ForeColor = CLR_DKGRY3
22850       .cost_subtot_lbl_dim_hi.Visible = False
22860       If .chkIncludeCurrency = True Then
22870         .curr_id.Locked = True
22880         .curr_id.BorderColor = CLR_LTBLU2
22890         .curr_id.BackStyle = acBackStyleNormal
22900         .curr_id_lbl.ForeColor = CLR_DKGRY3
22910         .curr_id_lbl_dim_hi.Visible = False
22920       End If
22930       .subtot_lbl.Visible = True
22940       .subtot_lbl_dim.Visible = False
22950       .subtot_lbl_dim_hi.Visible = False
22960       blnEnableSub_True = True
22970       .TimerInterval = 500&
22980     Case False
22990       .pending_trans.Locked = False  ' ** Looks disabled.
23000       .pending_trans.BorderColor = WIN_CLR_DISR
23010       .pending_trans.BackStyle = acBackStyleTransparent
23020       .shareface_subtot.Locked = False
23030       .shareface_subtot.BorderColor = WIN_CLR_DISR
23040       .shareface_subtot.BackStyle = acBackStyleTransparent
23050       .shareface_divint_subtot.Locked = False
23060       .shareface_divint_subtot.BorderColor = WIN_CLR_DISR
23070       .shareface_divint_subtot.BackStyle = acBackStyleTransparent
23080       .shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
23090       .shareface_subtot_lbl_dim_hi.Visible = True
23100       .icash_subtot.Locked = False
23110       .icash_subtot.BorderColor = WIN_CLR_DISR
23120       .icash_subtot.BackStyle = acBackStyleTransparent
23130       .icash_subtot_lbl.ForeColor = WIN_CLR_DISF
23140       .icash_subtot_lbl_dim_hi.Visible = True
23150       .pcash_subtot.Locked = False
23160       .pcash_subtot.BorderColor = WIN_CLR_DISR
23170       .pcash_subtot.BackStyle = acBackStyleTransparent
23180       .pcash_subtot_lbl.ForeColor = WIN_CLR_DISF
23190       .pcash_subtot_lbl_dim_hi.Visible = True
23200       .cost_subtot.Locked = False
23210       .cost_subtot.BorderColor = WIN_CLR_DISR
23220       .cost_subtot.BackStyle = acBackStyleTransparent
23230       .cost_subtot_lbl.ForeColor = WIN_CLR_DISF
23240       .cost_subtot_lbl_dim_hi.Visible = True
23250       If .chkIncludeCurrency = True Then
23260         .curr_id.Locked = False
23270         .curr_id.BorderColor = WIN_CLR_DISR
23280         .curr_id.BackStyle = acBackStyleTransparent
23290         .curr_id_lbl.ForeColor = WIN_CLR_DISF
23300         .curr_id_lbl_dim_hi.Visible = True
23310       End If
23320       .subtot_lbl.Visible = False
23330       .subtot_lbl_dim.Visible = True
23340       .subtot_lbl_dim_hi.Visible = True
23350       blnEnableSub_False = True
23360       .TimerInterval = 500&
23370     End Select
23380     DoEvents
23390   End With

EXITP:
23400   Exit Sub

ERRH:
23410   Select Case ERR.Number
        Case Else
23420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23430   End Select
23440   Resume EXITP

End Sub

Private Sub IncludeCurrency(blnInclude As Boolean)

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngCurrID As Long

23510   With Me
23520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
23530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
23540     End If
23550     Select Case blnInclude
          Case True
23560       .curr_id.Left = lngCurrID_Left
23570       .curr_id_lbl.Left = .curr_id.Left
23580       .curr_id_lbl_dim_hi.Left = (.curr_id_lbl.Left + lngTpp)
23590       .curr_id_box1.Left = (.curr_id.Left - (2& * lngTpp))
23600       .curr_id_box2.Left = (.curr_id.Left - (1& * lngTpp))
23610       .curr_id.Visible = True
23620       .curr_id_lbl.Visible = True
23630       If .pending_trans > 0 Then
23640         lngCurrID = .frmAccountHideTrans2_Sub_Pick.Form.curr_id
23650         If lngCurrID = 150& Then
23660           .curr_id.BackColor = CLR_DISABLED_BG
23670         Else
23680           .curr_id.BackColor = CLR_LTGRN
23690         End If
23700         .curr_id.BackStyle = acBackStyleNormal
23710         .curr_id.BorderColor = CLR_LTBLU2
23720         .curr_id_lbl.ForeColor = CLR_DKGRY3
23730         .curr_id_lbl_dim_hi.Visible = False
23740       Else
23750         .curr_id.BackStyle = acBackStyleTransparent
23760         .curr_id.BorderColor = WIN_CLR_DISR
23770         .curr_id_lbl.ForeColor = WIN_CLR_DISF
23780         .curr_id_lbl_dim_hi.Visible = True
23790       End If
23800       .cost_subtot.Left = lngCostSub_Left
23810       .cost_subtot_lbl.Left = .cost_subtot.Left
23820       .cost_subtot_lbl_dim_hi.Left = (.cost_subtot_lbl.Left + lngTpp)
23830       .cost_subtot_box1.Left = (.cost_subtot.Left - (2& * lngTpp))
23840       .cost_subtot_box2.Left = (.cost_subtot.Left - (1& * lngTpp))
23850       .pcash_subtot.Left = lngPCashSub_Left
23860       .pcash_subtot_lbl.Left = .pcash_subtot.Left
23870       .pcash_subtot_lbl_dim_hi.Left = (.pcash_subtot.Left + lngTpp)
23880       .pcash_subtot_box1.Left = (.pcash_subtot.Left - (2& * lngTpp))
23890       .pcash_subtot_box2.Left = (.pcash_subtot.Left - (1& * lngTpp))
23900       .icash_subtot.Left = lngICashSub_Left
23910       .icash_subtot_lbl.Left = .icash_subtot.Left
23920       .icash_subtot_lbl_dim_hi.Left = (.icash_subtot.Left + lngTpp)
23930       .icash_subtot_box1.Left = (.icash_subtot.Left - (2& * lngTpp))
23940       .icash_subtot_box2.Left = (.icash_subtot.Left - (1& * lngTpp))
23950       .shareface_subtot.Left = lngSharefaceSub_Left
23960       .shareface_divint_subtot.Left = .shareface_subtot.Left
23970       .shareface_subtot_lbl.Left = .shareface_subtot.Left
23980       .shareface_subtot_lbl_dim_hi.Left = (.shareface_subtot.Left + lngTpp)
23990       .shareface_subtot_box1.Left = (.shareface_subtot.Left - (2& * lngTpp))
24000       .shareface_subtot_box2.Left = (.shareface_subtot.Left - (1& * lngTpp))
24010       .subtot_lbl.Left = (.shareface_subtot.Left - lngSubTotLbl_Offset)
24020       .subtot_lbl_dim.Left = .subtot_lbl.Left
24030       .subtot_lbl_dim_hi.Left = (.subtot_lbl_dim.Left + lngTpp)
24040       If .pending_trans > 0 Then
24050         .subtot_lbl.Visible = True
24060         .subtot_lbl_dim.Visible = False
24070         .subtot_lbl_dim_hi.Visible = False
24080       Else
24090         .subtot_lbl.Visible = False
24100         .subtot_lbl_dim.Visible = True
24110         .subtot_lbl_dim_hi.Visible = True
24120       End If
24130       .pending_trans.Left = lngPTrans_Left
24140       .pending_trans_lbl.Left = (.pending_trans.Left - lngPTransLbl_Offset)
24150       .cmdCancel.Left = lngCancel_Left
24160       .cmdOK.Left = lngOK_Left
24170       .cmdUnHide.Left = .cmdOK.Left
24180     Case False
24190       .curr_id.Visible = False
24200       .curr_id_box1.Visible = False
24210       .curr_id_box2.Visible = False
24220       .curr_id_lbl.Visible = False
24230       .curr_id_lbl_dim_hi.Visible = False
24240       .curr_id.Left = 0&
24250       .curr_id_box1.Left = 0&
24260       .curr_id_box2.Left = 0&
24270       .curr_id_lbl.Left = 0&
24280       .curr_id_lbl_dim_hi.Left = 0&
24290       .cost_subtot.Left = (lngCostSub_Left + lngCurrID_Width)
24300       .cost_subtot_lbl.Left = .cost_subtot.Left
24310       .cost_subtot_lbl_dim_hi.Left = (.cost_subtot.Left + lngTpp)
24320       .cost_subtot_box1.Left = (.cost_subtot.Left - (2& * lngTpp))
24330       .cost_subtot_box2.Left = (.cost_subtot.Left - (1& * lngTpp))
24340       .pcash_subtot.Left = (lngPCashSub_Left + lngCurrID_Width)
24350       .pcash_subtot_lbl.Left = .pcash_subtot.Left
24360       .pcash_subtot_lbl_dim_hi.Left = (.pcash_subtot.Left + lngTpp)
24370       .pcash_subtot_box1.Left = (.pcash_subtot.Left - (2& * lngTpp))
24380       .pcash_subtot_box2.Left = (.pcash_subtot.Left - (1& * lngTpp))
24390       .icash_subtot.Left = (lngICashSub_Left + lngCurrID_Width)
24400       .icash_subtot_lbl.Left = .icash_subtot.Left
24410       .icash_subtot_lbl_dim_hi.Left = (.icash_subtot.Left + lngTpp)
24420       .icash_subtot_box1.Left = (.icash_subtot.Left - (2& * lngTpp))
24430       .icash_subtot_box2.Left = (.icash_subtot.Left - (1& * lngTpp))
24440       .shareface_subtot.Left = (lngSharefaceSub_Left + lngCurrID_Width)
24450       .shareface_subtot_lbl.Left = .shareface_subtot.Left
24460       .shareface_subtot_lbl_dim_hi.Left = (.shareface_subtot.Left + lngTpp)
24470       .shareface_divint_subtot.Left = .shareface_subtot.Left
24480       .shareface_subtot_box1.Left = (.shareface_subtot.Left - (2& * lngTpp))
24490       .shareface_subtot_box2.Left = (.shareface_subtot.Left - (1& * lngTpp))
24500       .subtot_lbl.Left = (.shareface_subtot.Left - lngSubTotLbl_Offset)
24510       .subtot_lbl_dim.Left = .subtot_lbl.Left
24520       .subtot_lbl_dim_hi.Left = (.subtot_lbl_dim.Left + lngTpp)
24530       If .pending_trans > 0 Then
24540         .subtot_lbl.Visible = True
24550         .subtot_lbl_dim.Visible = False
24560         .subtot_lbl_dim_hi.Visible = False
24570       Else
24580         .subtot_lbl.Visible = False
24590         .subtot_lbl_dim.Visible = True
24600         .subtot_lbl_dim_hi.Visible = True
24610       End If
24620       .pending_trans.Left = (lngPTrans_Left + lngCurrID_Width)
24630       .pending_trans_lbl.Left = (.pending_trans.Left - lngPTransLbl_Offset)
24640       .cmdCancel.Left = (lngCancel_Left + lngCurrID_Width)
24650       .cmdOK.Left = (lngOK_Left + lngCurrID_Width)
24660       .cmdUnHide.Left = .cmdOK.Left
24670     End Select
24680   End With

EXITP:
24690   Exit Sub

ERRH:
24700   Select Case ERR.Number
        Case Else
24710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24720   End Select
24730   Resume EXITP

End Sub

Public Sub CheckBoxLock(blnLock As Boolean)

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "CheckBoxLock"

24810   With Me
24820     Select Case blnLock
          Case True
24830       .chkIncludeArchive.Locked = True
24840       .chkIncludeArchive_lbl.ForeColor = CLR_DKGRY
24850       .chkShowHiddenOnly.Locked = True
24860       .chkShowHiddenOnly_lbl.ForeColor = CLR_DKGRY
24870       .chkShowAll.Locked = True
24880       .chkShowAll_lbl.ForeColor = CLR_DKGRY
24890       .cmdHidden.Enabled = False
24900     Case False
24910       Select Case .chkShowHiddenOnly
            Case True
24920         .chkIncludeArchive.Locked = True
24930         .chkIncludeArchive_lbl.ForeColor = CLR_DKGRY
24940         .chkShowAll.Locked = True
24950         .chkShowAll_lbl.ForeColor = CLR_DKGRY
24960       Case False
24970         .chkIncludeArchive.Locked = False
24980         .chkIncludeArchive_lbl.ForeColor = CLR_BLK
24990         .chkShowAll.Locked = False
25000         .chkShowAll_lbl.ForeColor = CLR_BLK
25010       End Select
25020       .chkShowHiddenOnly.Locked = False
25030       .chkShowHiddenOnly_lbl.ForeColor = CLR_BLK
25040       If .hidden_trans > 0 Then
25050         .cmdHidden.Enabled = True
25060       End If
25070     End Select
25080   End With

EXITP:
25090   Exit Sub

ERRH:
25100   Select Case ERR.Number
        Case Else
25110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25120   End Select
25130   Resume EXITP

End Sub

Public Sub UpdateHiddenTrans()

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateHiddenTrans"

        Dim varTmp00 As Variant

25210   With Me
          ' ** tmpAccountHideTrans_02, grouped, with hidden_trans.
25220     varTmp00 = DLookup("[hidden_trans]", "qryAccountHideTrans2_18")
25230     If IsNull(varTmp00) = True Then varTmp00 = 0&
25240     .hidden_trans_display = CLng(varTmp00)
25250   End With

EXITP:
25260   Exit Sub

ERRH:
25270   Select Case ERR.Number
        Case Else
25280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25290   End Select
25300   Resume EXITP

End Sub

Public Function ValidateGroup() As Boolean

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "ValidateGroup"

        Dim curICash As Currency, curPCash As Currency, curCost As Currency
        Dim dblShareface As Double, dblDivInt As Double
        Dim lngCurrID As Long, blnCurrID As Boolean
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

25410   blnRetVal = False

25420   With Me

25430     curICash = 0@: curPCash = 0@: curCost = 0@
25440     dblShareface = 0#: dblDivInt = 0#
25450     blnCurrID = True

25460     .FocusHolder.SetFocus
25470     .cmdOK.Enabled = False
25480     .shareface_subtot_box1.Visible = False
25490     .shareface_subtot_box2.Visible = False
25500     .icash_subtot_box1.Visible = False
25510     .icash_subtot_box2.Visible = False
25520     .pcash_subtot_box1.Visible = False
25530     .pcash_subtot_box2.Visible = False
25540     .cost_subtot_box1.Visible = False
25550     .cost_subtot_box2.Visible = False
25560     .curr_id_box1.Visible = False
25570     .curr_id_box2.Visible = False

25580     varTmp00 = DCount("*", "tmpAccountHideTrans_03")
25590     lngRecsCur = CLng(Nz(varTmp00, 0))

25600     shareface_chk  ' ** Procedure: Below.
25610     DoEvents

25620     If lngRecsCur > 0 Then

25630       varTmp00 = DSum("[sharefacex]", "tmpAccountHideTrans_03")
25640       dblShareface = CDbl(Nz(varTmp00, 0))
25650       If .shareface_subtot <> dblShareface Then .shareface_subtot = dblShareface
25660       varTmp00 = DSum("[shareface_divint]", "tmpAccountHideTrans_03")
25670       dblDivInt = CDbl(Nz(varTmp00, 0))
25680       If .shareface_divint_subtot <> dblDivInt Then .shareface_divint_subtot = dblDivInt
25690       varTmp00 = DSum("[icash]", "tmpAccountHideTrans_03")
25700       curICash = CCur(Nz(varTmp00, 0))
25710       If .icash_subtot <> curICash Then .icash_subtot = curICash
25720       varTmp00 = DSum("[pcash]", "tmpAccountHideTrans_03")
25730       curPCash = CCur(Nz(varTmp00, 0))
25740       If .pcash_subtot <> curPCash Then .pcash_subtot = curPCash
25750       varTmp00 = DSum("[cost]", "tmpAccountHideTrans_03")
25760       curCost = CCur(Nz(varTmp00, 0))
25770       If .cost_subtot <> curCost Then .cost_subtot = curCost

25780       If .chkIncludeCurrency = True Then
25790         lngCurrID = DLookup("[curr_id]", "tmpAccountHideTrans_03")  ' ** Arbitrary.
25800         varTmp00 = DCount("*", "tmpAccountHideTrans_03", "[curr_id] = " & CStr(lngCurrID))
25810         If varTmp00 <> lngRecsCur Then  ' ** Records have mixed curr_id.
25820           blnCurrID = False
25830         End If
25840       End If

25850       If dblShareface = 0# And curICash = 0@ And curPCash = 0@ And curCost = 0@ And blnCurrID = True Then
25860         blnRetVal = True
25870         .cmdOK.Enabled = True
25880         If .cnt_asset > 0 Then
25890           .shareface_subtot_box1.Visible = True
25900           .shareface_subtot_box2.Visible = True
25910         Else
25920           .shareface_subtot_box1.Visible = False
25930           .shareface_subtot_box2.Visible = False
25940         End If
25950         .icash_subtot_box1.Visible = True
25960         .icash_subtot_box2.Visible = True
25970         .pcash_subtot_box1.Visible = True
25980         .pcash_subtot_box2.Visible = True
25990         .cost_subtot_box1.Visible = True
26000         .cost_subtot_box2.Visible = True
26010         If .chkIncludeCurrency = True Then
26020           .curr_id_box1.Visible = True
26030           .curr_id_box2.Visible = True
26040         End If
26050       Else
26060         If dblShareface = 0# And .cnt_asset > 0 Then
26070           .shareface_subtot_box1.Visible = True
26080           .shareface_subtot_box2.Visible = True
26090         End If
26100         If curICash = 0@ Then
26110           .icash_subtot_box1.Visible = True
26120           .icash_subtot_box2.Visible = True
26130         End If
26140         If curPCash = 0@ Then
26150           .pcash_subtot_box1.Visible = True
26160           .pcash_subtot_box2.Visible = True
26170         End If
26180         If curCost = 0@ Then
26190           .cost_subtot_box1.Visible = True
26200           .cost_subtot_box2.Visible = True
26210         End If
26220         If .chkIncludeCurrency = True And blnCurrID = True Then
                ' ** Only turn these on with the whole group.
                '.curr_id_box1.Visible = True
                '.curr_id_box2.Visible = True
26230         End If
26240       End If

            ' ** LedgHidType enumeration:
            ' **   0  GRP_NONE    Unmatched hidden transactions. This should only be a temporary designation.
            ' **   1  NORM        2 entries in hidden group, with matching assetno (which could both be zero).
            ' **   2  NORM_MISC   2 entries in hidden group, where both are 'Misc.', to be treated like a normal pair.
            ' **   3  MISC_2_GRP  2 entries in hidden group, one 'Misc.' and one other.
            ' **   4  MISC_3_GRP  3 entries in hidden group, one 'Misc.' and two other matching assetno.
            ' **   5  MULTI_GRP   3 or more entries in hidden group, with matching assetno, multi-lot group

            ' ** Determine the Hidden Type.
26250       If blnRetVal = True Then
26260         lngHidType = 0&
26270         If lngRecsCur = 2& Then
                ' ** tmpAccountHideTrans_03, grouped by assetno.
26280           varTmp00 = DCount("*", "qryAccountHideTrans2_40a")
26290           If varTmp00 = 1 Then
                  ' ** Two matching assetno's, though both could be 0.
                  ' ** qryAccountHideTrans2_40b (tmpAccountHideTrans_03, with IsMisc), grouped and summed.
26300             varTmp00 = DLookup("[IsMisc]", "qryAccountHideTrans2_40c")
26310             Select Case Abs(varTmp00)
                  Case 0
26320               lngHidType = 1&
26330             Case 1  ' ** I don't think this'll ever hit.
26340               lngHidType = 3&
26350             Case 2
26360               lngHidType = 2&
26370             End Select
26380           Else
                  ' ** Mismatched assetno's.
                  ' ** qryAccountHideTrans2_40b (tmpAccountHideTrans_03, with IsMisc), grouped and summed.
26390             varTmp00 = DLookup("[IsMisc]", "qryAccountHideTrans2_40c")
26400             Select Case Abs(varTmp00)
                  Case 0  ' ** I don't think this'll ever hit.
26410               lngHidType = 1&
26420             Case 1
26430               lngHidType = 3&
26440             Case 2  ' ** I don't think this'll ever hit either!
26450               lngHidType = 2&
26460             End Select
26470           End If
26480         ElseIf lngRecsCur = 3& Then
                ' ** tmpAccountHideTrans_03, grouped by assetno.
26490           varTmp00 = DCount("*", "qryAccountHideTrans2_40a")
26500           Select Case varTmp00
                Case 1
26510             lngHidType = 5&
26520           Case 2
                  ' ** qryAccountHideTrans2_40b (tmpAccountHideTrans_03, with IsMisc), grouped and summed.
26530             varTmp00 = DLookup("[IsMisc]", "qryAccountHideTrans2_40c")
26540             If Abs(varTmp00) = 1 Then
26550               lngHidType = 4&
26560             Else
                    ' ** I'm not sure if 2 Misc. and 1 assetno is legitimate.
26570             End If
26580           Case Else
                  ' ** Shouldn't get here!
26590           End Select
26600         ElseIf lngRecsCur > 3& Then
26610           lngHidType = 5&
26620         End If
              ' ** Now get lngHidType to where we need it!
26630       End If  ' ** blnRetVal.

26640     End If  ' ** lngRecsCur.

26650   End With

EXITP:
26660   ValidateGroup = blnRetVal
26670   Exit Function

ERRH:
26680   blnRetVal = False
26690   Select Case ERR.Number
        Case Else
26700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26710   End Select
26720   Resume EXITP

End Function

Private Sub ArchiveSet(blnHasArchs As Boolean)

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "ArchiveSet"

26810   With Me
26820     Select Case blnHasArchive  ' ** blnHasArchive is whether there are ANY archive records.
          Case True
26830       Select Case blnHasArchs
            Case True
26840         .chkIncludeArchive_lbl.Visible = True
26850         .chkIncludeArchive_lbl2.Visible = False
26860         .chkIncludeArchive_lbl2_dim_hi.Visible = False
26870         .chkIncludeArchive_lbl3.Visible = False
26880         .chkIncludeArchive_lbl3_dim_hi.Visible = False
26890         .chkIncludeArchive_lbl4.Visible = False
26900         .chkIncludeArchive_lbl4_dim_hi.Visible = False
26910         .chkIncludeArchive_lbl5.Visible = False
26920         .chkIncludeArchive_lbl5_dim_hi.Visible = False
26930       Case False
              ' ** This Account Has No Archived Transactions.
26940         .chkIncludeArchive_lbl.Visible = False
26950         .chkIncludeArchive_lbl2.Visible = False
26960         .chkIncludeArchive_lbl2_dim_hi.Visible = False
26970         .chkIncludeArchive_lbl3.Visible = False
26980         .chkIncludeArchive_lbl3_dim_hi.Visible = False
26990         .chkIncludeArchive_lbl4.Visible = True
27000         .chkIncludeArchive_lbl4_dim_hi.Visible = True
27010         .chkIncludeArchive_lbl5.Visible = True
27020         .chkIncludeArchive_lbl5_dim_hi.Visible = True
27030       End Select
27040     Case False
            ' ** There Are No Archived Transactions.
27050       .chkIncludeArchive_lbl.Visible = False
27060       .chkIncludeArchive_lbl2.Visible = True
27070       .chkIncludeArchive_lbl2_dim_hi.Visible = True
27080       .chkIncludeArchive_lbl3.Visible = True
27090       .chkIncludeArchive_lbl3_dim_hi.Visible = True
27100       .chkIncludeArchive_lbl4.Visible = False
27110       .chkIncludeArchive_lbl4_dim_hi.Visible = False
27120       .chkIncludeArchive_lbl5.Visible = False
27130       .chkIncludeArchive_lbl5_dim_hi.Visible = False
27140     End Select
27150   End With

EXITP:
27160   Exit Sub

ERRH:
27170   Select Case ERR.Number
        Case Else
27180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27190   End Select
27200   Resume EXITP

End Sub

Public Sub UpdateTotals_Set(blnUpdate As Boolean, lngTimer As Long)

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateTotals_Set"

27310   With Me
27320     blnUpdateTotals = blnUpdate
27330     .TimerInterval = lngTimer
27340   End With

EXITP:
27350   Exit Sub

ERRH:
27360   Select Case ERR.Number
        Case Else
27370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27380   End Select
27390   Resume EXITP

End Sub

Public Sub shareface_chk()
' ** Determine whether this has asset transactions.

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_chk"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long
        Dim lngX As Long

27410   With Me
27420     lngTmp01 = 0&: lngTmp02 = 0&: lngTmp03 = 0&: lngTmp04 = 0&
27430     lngRecsCur = .frmAccountHideTrans2_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_Sub_Pick.
27440     If lngRecsCur > 0& Then
27450       Set rst = .frmAccountHideTrans2_Sub_Pick.Form.RecordsetClone
27460       With rst
27470         .MoveLast
27480         lngRecs = .RecordCount
27490         .MoveFirst
27500         For lngX = 1& To lngRecs
27510           Select Case ![journaltype]
                Case "Deposit", "Purchase", "Withdrawn", "Sold", "Liability"
27520             lngTmp01 = lngTmp01 + 1&
27530           Case "Cost Adj."
27540             lngTmp02 = lngTmp02 + 1&
27550           Case "Dividend", "Interest"
27560             lngTmp03 = lngTmp03 + 1&
27570           Case Else
27580             lngTmp04 = lngTmp04 + 1&
27590           End Select
27600           If lngX < lngRecs Then .MoveNext
27610         Next
27620         .Close
27630       End With
27640       Set rst = Nothing
27650       If lngTmp01 = 0& Then
27660         If lngTmp03 > 0& Then
                ' ** Though shareface is listed, it has no bearing on the match.
27670           .shareface_divint_subtot.Visible = True
27680           .shareface_subtot.Visible = False
27690         ElseIf lngTmp02 > 0& Then
                ' ** This option can't occur, because it has an asset, and lngTmp01 = 0&.
27700           .shareface_subtot.Visible = True
27710           .shareface_divint_subtot.Visible = False
27720         Else
                ' ** These will be Zero.
27730           .shareface_subtot.Visible = True
27740           .shareface_divint_subtot.Visible = False
27750         End If
27760       Else
              ' ** This must total to Zero.
27770         .shareface_subtot.Visible = True
27780         .shareface_divint_subtot.Visible = False
27790       End If
27800     End If
27810     .cnt_asset = lngTmp01
27820     .cnt_costadj = lngTmp02
27830     .cnt_divint = lngTmp03
27840     .cnt_misc = lngTmp04
27850     DoEvents
27860   End With

EXITP:
27870   Set rst = Nothing
27880   Exit Sub

ERRH:
27890   Select Case ERR.Number
        Case Else
27900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27910   End Select
27920   Resume EXITP

End Sub

Private Sub RenumGrps()

28000 On Error GoTo ERRH

        Const THIS_PROC As String = "RenumGrps"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long

28010   Set dbs = CurrentDb
28020   With dbs

          ' ** Check for and fix those high-number groups.
28030     GrpNum_Chk  ' ** Function: Below.

          'THIS SEEMS TO BE RENUMBERING ALL GROUPS FOR ALL ACCOUNTS!
          'SHOULDN'T IT JUST BE DOING JUST THIS ACCOUNT?
          ' ** Update qryAccountHideTrans2_Hidden_09_01 (tblLedgerHidden,
          ' ** with ledghid_grpnum_new (+10000)), by specified [actno].
28040     Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_09_02")
28050     With qdf.Parameters
28060       ![actno] = gstrAccountNo
28070     End With
28080     qdf.Execute dbFailOnError
28090     Set qdf = Nothing

          ' ** Update tblMark, for value = value + 10000.
28100     Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_09_03")
28110     qdf.Execute
28120     Set qdf = Nothing

28130     Set rst = .OpenRecordset("tblMark", dbOpenDynaset, dbConsistent)
28140     With rst
28150       .MoveLast
28160       lngRecs = .RecordCount
28170       .MoveFirst
28180       For lngX = 1& To lngRecs
              ' ** Update qryAccountHideTrans2_Hidden_09_04 (tblLedgerHidden, by specified [grpnum], [ngrpnum]).
28190         Set qdf = dbs.QueryDefs("qryAccountHideTrans2_Hidden_09_05")
28200         With qdf.Parameters
28210           ![grpnum] = CLng(rst![Value])
28220           ![ngrpnum] = rst![unique_id]
28230         End With
28240         qdf.Execute
28250         Set qdf = Nothing
28260         If lngX < lngRecs Then .MoveNext
28270       Next  ' ** lngX.
28280       .Close
28290     End With  ' ** rst.

28300     .Close
28310   End With  ' ** dbs.

EXITP:
28320   Set rst = Nothing
28330   Set qdf = Nothing
28340   Set dbs = Nothing
28350   Exit Sub

ERRH:
28360   DoCmd.Hourglass False
28370   Select Case ERR.Number
        Case Else
28380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28390   End Select
28400   Resume EXITP

End Sub

Private Function GrpNum_Chk() As Boolean
' ** There always seems to be groups with 10000 range numbers
' ** that didn't get put back to their new low number!

28500 On Error GoTo ERRH

        Const THIS_PROC As String = "GrpNum_Chk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngGrps As Long, arr_varGrp As Variant
        Dim lngCnt As Long
        Dim blnContinue As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varGrp().
        Const G_ACTNO As Integer = 0
        Const G_GNUM  As Integer = 1
        Const G_GNUMN As Integer = 2

28510   blnRetVal = True
28520   blnContinue = True

28530   Set dbs = CurrentDb
28540   With dbs

          ' ** This handles high-number groups for all accounts.

          ' ** qryAccountHideTrans2_Hidden_09_10 (tblLedgerHidden, just
          ' ** ledghid_grpnum > 1000, with ledghid_grpnum_new), grouped, with cnt_xhid.
28550     Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_09_11")
28560     Set rst = qdf.OpenRecordset
28570     With rst
28580       If .BOF = True And .EOF = True Then
              ' ** No stragglers.
28590         blnContinue = False
28600       Else
28610         .MoveFirst
28620         If ![cnt_xhid] = 0 Then
28630           blnContinue = False
28640         Else
                ' ** Yup, gotta fix 'em.
28650         End If
28660       End If
28670       .Close
28680     End With  ' ** rst.
28690     Set rst = Nothing
28700     Set qdf = Nothing

28710     If blnContinue = True Then
            ' ** We've got stragglers to renumber!

            ' ** Update qryAccountHideTrans2_Hidden_09_22 (qryAccountHideTrans2_Hidden_09_10
            ' ** (tblLedgerHidden, just ledghid_grpnum > 1000, with ledghid_grpnum_new), not
            ' ** in qryAccountHideTrans2_Hidden_09_12 (qryAccountHideTrans2_Hidden_09_10
            ' ** (tblLedgerHidden, just ledghid_grpnum > 1000, with ledghid_grpnum_new),
            ' ** linked back to tblLedgerHidden, via ledghid_grpnum_new, conflict)).
28720       Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_09_23")
28730       qdf.Execute
28740       Set qdf = Nothing

            ' ** qryAccountHideTrans2_Hidden_09_15 (qryAccountHideTrans2_Hidden_09_12
            ' ** (qryAccountHideTrans2_Hidden_09_10 (tblLedgerHidden, just ledghid_grpnum > 1000,
            ' ** with ledghid_grpnum_new), linked back to tblLedgerHidden, via ledghid_grpnum_new,
            ' ** conflict), grouped, by accountno, ledghid_grpnum), grouped by accountno, with cnt_grp > 1.
28750       Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_09_16")
28760       Set rst = qdf.OpenRecordset
28770       With rst
28780         If .BOF = True And .EOF = True Then
                ' ** Good. Only 1 group per account.
28790           lngCnt = 1&
28800         Else
                ' ** At least one account has more than 1 bad group!
28810           .MoveLast
28820           lngCnt = .RecordCount
28830         End If
28840         .Close
28850       End With  ' ** rst.
28860       Set rst = Nothing
28870       Set qdf = Nothing

            ' ** qryAccountHideTrans2_Hidden_09_18 (qryAccountHideTrans2_Hidden_09_17 (qryAccountHideTrans2_Hidden_09_15 (
            ' ** qryAccountHideTrans2_Hidden_09_12 (qryAccountHideTrans2_Hidden_09_10 (tblLedgerHidden, just ledghid_grpnum > 1000,
            ' ** with ledghid_grpnum_new), linked back to tblLedgerHidden, via ledghid_grpnum_new, conflict), grouped, by
            ' ** accountno, ledghid_grpnum), not in qryAccountHideTrans2_Hidden_09_16 (qryAccountHideTrans2_Hidden_09_15
            ' ** (qryAccountHideTrans2_Hidden_09_12 (qryAccountHideTrans2_Hidden_09_10 (tblLedgerHidden, just ledghid_grpnum > 1000,
            ' ** with ledghid_grpnum_new), linked back to tblLedgerHidden, via ledghid_grpnum_new, conflict), grouped, by accountno,
            ' ** ledghid_grpnum), grouped by accountno, with cnt_grp > 1)), linked to qryAccountHideTrans2_Hidden_09_14
            ' ** (qryAccountHideTrans2_Hidden_09_13 (qryAccountHideTrans2_Hidden_09_12 (qryAccountHideTrans2_Hidden_09_10
            ' ** (tblLedgerHidden, just ledghid_grpnum > 1000, with ledghid_grpnum_new), linked back to tblLedgerHidden, via
            ' ** ledghid_grpnum_new, conflict), grouped by accountno), linked back to tblLedgerHidden, grouped, by accountno,
            ' ** ledghid_grpnum, without ledghid_grpnum > 1000), grouped, with ledghid_grpnum_max), with new ledghid_grpnum_new.
28880       Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_09_19")
28890       Set rst = qdf.OpenRecordset
28900       With rst
28910         .MoveLast
28920         lngGrps = .RecordCount
28930         .MoveFirst
28940         arr_varGrp = .GetRows(lngGrps)
              ' *******************************************************
              ' ** Array: arr_varGrp()
              ' **
              ' **   Field  Element  Name                  Constant
              ' **   =====  =======  ====================  ==========
              ' **     1       0     accountno             G_ACTNO
              ' **     1       0     ledghid_grpnum        G_GNUM
              ' **     1       0     ledghid_grpnum_new    G_GNUMN
              ' **
              ' *******************************************************
28950         .Close
28960       End With  ' ** rst.
28970       Set rst = Nothing
28980       Set qdf = Nothing

            ' ** These are just single groups per account.
28990       If lngGrps > 0& Then
29000         For lngX = 0& To (lngGrps - 1&)
                ' ** Update qryAccountHideTrans2_Hidden_09_20
                ' ** (tblLedgerHidden, by specified [actno], [ogrp], [ngrp]).
29010           Set qdf = .QueryDefs("qryAccountHideTrans2_Hidden_09_21")
29020           With qdf.Parameters
29030             ![actno] = arr_varGrp(G_ACTNO, lngX)
29040             ![ogrp] = arr_varGrp(G_GNUM, lngX)
29050             ![ngrp] = arr_varGrp(G_GNUMN, lngX)
29060           End With
29070           qdf.Execute
29080         Next  ' ** lngX.
29090         DoEvents
29100         Set qdf = Nothing
29110       End If  ' ** lngGrps

29120     End If  ' ** blnContinue.

29130     .Close
29140   End With  ' ** dbs.

EXITP:
29150   Set rst = Nothing
29160   Set qdf = Nothing
29170   Set dbs = Nothing
29180   GrpNum_Chk = blnRetVal
29190   Exit Function

ERRH:
29200   blnRetVal = False
29210   Select Case ERR.Number
        Case Else
29220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29230   End Select
29240   Resume EXITP

End Function

Private Sub EmptyTemps(blnShow As Boolean)

29300 On Error GoTo ERRH

        Const THIS_PROC As String = "EmptyTemps"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

29310   With Me
29320     Select Case blnShow
          Case True
29330       With .frmAccountHideTrans2_Sub_List.Form
29340         .FocusHolder.SetFocus
29350         .journalno.ForeColor = CLR_DISABLED_FG
29360         .transdate.ForeColor = CLR_DISABLED_FG
29370         .journaltype.ForeColor = CLR_DISABLED_FG
29380         .shareface_0.ForeColor = CLR_DISABLED_FG
29390         .shareface_1.ForeColor = CLR_DISABLED_FG
29400         .shareface_2.ForeColor = CLR_DISABLED_FG
29410         .shareface_3.ForeColor = CLR_DISABLED_FG
29420         .shareface_4.ForeColor = CLR_DISABLED_FG
29430         .totdesc.ForeColor = CLR_DISABLED_FG
29440         .ICash.ForeColor = CLR_DISABLED_FG
29450         .icash_str.ForeColor = CLR_DISABLED_FG
29460         .PCash.ForeColor = CLR_DISABLED_FG
29470         .pcash_str.ForeColor = CLR_DISABLED_FG
29480         .Cost.ForeColor = CLR_DISABLED_FG
29490         .cost_str.ForeColor = CLR_DISABLED_FG
29500         .curr_id.ForeColor = CLR_DISABLED_FG
29510         .ledger_HIDDEN_box.ForeColor = MY_CLR_VLTBGE
29520         .ledger_HIDDEN.Visible = True
29530         .ledger_HIDDEN_dummy.Visible = False
29540         DoEvents
29550       End With
29560     Case False
            ' ** Blank these out so we don't see '#DELETED'.
29570       With .frmAccountHideTrans2_Sub_List.Form
29580         .FocusHolder.SetFocus
29590         .journalno.ForeColor = CLR_DISABLED_BG
29600         .transdate.ForeColor = CLR_DISABLED_BG
29610         .journaltype.ForeColor = CLR_DISABLED_BG
29620         .shareface_0.ForeColor = CLR_DISABLED_BG
29630         .shareface_1.ForeColor = CLR_DISABLED_BG
29640         .shareface_2.ForeColor = CLR_DISABLED_BG
29650         .shareface_3.ForeColor = CLR_DISABLED_BG
29660         .shareface_4.ForeColor = CLR_DISABLED_BG
29670         .totdesc.ForeColor = CLR_DISABLED_BG
29680         .ICash.ForeColor = CLR_DISABLED_BG
29690         .icash_str.ForeColor = CLR_DISABLED_BG
29700         .PCash.ForeColor = CLR_DISABLED_BG
29710         .pcash_str.ForeColor = CLR_DISABLED_BG
29720         .Cost.ForeColor = CLR_DISABLED_BG
29730         .cost_str.ForeColor = CLR_DISABLED_BG
29740         .curr_id.ForeColor = CLR_DISABLED_BG
29750         .FocusHolder.SetFocus
29760         DoEvents
29770         .ledger_HIDDEN_box.ForeColor = MY_CLR_BGE
29780         .ledger_HIDDEN_dummy.Visible = True
29790         .ledger_HIDDEN.Visible = False
29800         DoEvents
29810       End With

29820       Set dbs = CurrentDb
29830       With dbs
              ' ** Empty tmpAccountHideTrans_01.
29840         Set qdf = .QueryDefs("qryAccountHideTrans2_01a")
29850         qdf.Execute
29860         Set qdf = Nothing
              ' ** Empty tmpAccountHideTrans_02.
29870         Set qdf = .QueryDefs("qryAccountHideTrans2_01b")
29880         qdf.Execute
29890         Set qdf = Nothing
              ' ** Empty tmpAccountHideTrans_03.
29900         Set qdf = .QueryDefs("qryAccountHideTrans2_01c")
29910         qdf.Execute
29920         Set qdf = Nothing
              ' ** Empty tmpAccountHideTrans_04.
29930         Set qdf = .QueryDefs("qryAccountHideTrans2_01d")
29940         qdf.Execute
29950         Set qdf = Nothing
29960         DoEvents
29970         .Close
29980       End With
29990     End Select  ' ** blnShow.

30000   End With

EXITP:
30010   Set qdf = Nothing
30020   Set dbs = Nothing
30030   Exit Sub

ERRH:
30040   Select Case ERR.Number
        Case Else
30050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30060   End Select
30070   Resume EXITP

End Sub

Public Sub GTREmblem_Set(blnSetEmblem As Boolean)

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Set"

30110   With Me
30120     blnGTR_Emblem = blnSetEmblem
30130     If blnGTR_Emblem = False Then
30140       .TimerInterval = 0&
30150       .GoToReport_arw_hide_img.Visible = False
30160       blnGoingToReport = False
30170       GTREmblem_Off  ' ** Procedure: Below.
30180     End If
30190   End With

EXITP:
30200   Exit Sub

ERRH:
30210   DoCmd.Hourglass False
30220   Select Case ERR.Number
        Case Else
30230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30240   End Select
30250   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

30310   With Me
30320     blnGTR_Emblem = False
30330     For lngX = 1& To 24&
30340       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
30350     Next
30360   End With

EXITP:
30370   Exit Sub

ERRH:
30380   Select Case ERR.Number
        Case Else
30390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30400   End Select
30410   Resume EXITP

End Sub
