VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans2_Sub_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans2_Sub_List"

'VGC 04/15/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Show All:         A {chkShowAll on frmAccountHideTrans2}
' **   Close:            C {cmdClose on frmAccountHideTrans2}
' **   ledger_HIDDEN:    D {ledger_HIDDEN}
' **   Hidden Groups:    G {cmdHidden on frmAccountHideTrans2}
' **   Cancel:           E {cmdCancel on frmAccountHideTrans2}
' **   Include Archive:  H {chkIncludeArchive on frmAccountHideTrans2}
' **   UnHide:           I {cmdUnHide on frmAccountHideTrans2}
' **   Journal No:       J {chkShowJournalNo on frmAccountHideTrans2}
' **   ledger_HIDDEN:    N {ledger_HIDDEN on frmAccountHideTrans2_Sub_Pick}
' **   OK:               O {cmdOK on frmAccountHideTrans2}
' **   Print Report:     P {cmdPrintReport on frmAccountHideTrans2}
' **   Show Hidden Only: S {chkShowHiddenOnly on frmAccountHideTrans2}
' **   Include Currency: U {chkIncludeCurrency on frmAccountHideTrans2}
' **   Exit:             X {cmdClose on frmAccountHideTrans2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[transdate] DESC, [JournalType_Order], [journalno]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngForm_Width As Long, lngJournalNo_Width As Long, lngCurrID_Width As Long
Private lngJournalNo_Left As Long, lngTransdate_Left As Long, lngShareface_Left As Long, lngTotDesc_Left As Long, lngTotDesc_Width As Long
Private lngICash_Left As Long, lngICashLbl_Offset As Long, lngPCash_Left As Long, lngCost_Left As Long
Private lngCurrID_Left As Long, lngCurrIDLbl_Offset As Long, lngJournalType_Left As Long
Private lngLedgerHid_Left As Long, lngLedgerHidLbl_Offset As Long, lngLedgerHidLblLine_Offset As Long

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Top As Long, lngSortLbl_Width As Long
Private lngSortLine_Left As Long, lngSortLine_Width As Long
Private blnCancelUpdate As Boolean, blnUnHide As Boolean
Private lngRecsCur As Long, lngJournalNo As Long
Private strCallingForm As String, lngTpp As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountHideTrans2" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

190         blnCancelUpdate = False: blnUnHide = False

200         .Sort_lbl.Visible = False
210         .Sort_line.Visible = False
220         lngSortLbl_Top = .Sort_lbl.Top
230         lngSortLbl_Width = .Sort_lbl.Width

240         lngForm_Width = .Width
250         lngJournalNo_Left = .journalno.Left
260         lngTransdate_Left = .transdate.Left
270         lngJournalType_Left = .journaltype.Left
280         lngShareface_Left = .shareface_0.Left
290         lngTotDesc_Left = .totdesc.Left
300         lngTotDesc_Width = .totdesc.Width
310         lngICash_Left = .ICash.Left
320         lngICashLbl_Offset = (.ICash.Left - .icash_lbl.Left)
330         lngPCash_Left = .PCash.Left
340         lngCost_Left = .Cost.Left
350         lngCurrID_Left = .curr_id.Left
360         lngCurrIDLbl_Offset = (.curr_id.Left - .curr_id_lbl.Left)
370         lngLedgerHid_Left = .ledger_HIDDEN.Left
380         lngLedgerHidLbl_Offset = (.ledger_HIDDEN.Left - .ledger_HIDDEN_lbl.Left)
390         lngLedgerHidLblLine_Offset = (.ledger_HIDDEN.Left - .ledger_HIDDEN_lbl_line.Left)
400         lngJournalNo_Width = (.transdate.Left - .journalno.Left)
410         lngCurrID_Width = (.ledger_HIDDEN_lbl_line.Left - .curr_id.Left)

420       Else
430         Cancel = -1
440       End If
450     End With

460     If Cancel = -1 Then
470       Beep
480       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
490     End If

EXITP:
500     Exit Sub

ERRH:
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me

620       SortNow THIS_PROC  ' ** Procedure: Below

          ' ** Though most of these are transparent.
630       .journalno.ForeColor = CLR_DISABLED_FG
640       .journalno.BackColor = CLR_DISABLED_BG
650       .transdate.ForeColor = CLR_DISABLED_FG
660       .transdate.BackColor = CLR_DISABLED_BG
670       .journaltype.ForeColor = CLR_DISABLED_FG
680       .journaltype.BackColor = CLR_DISABLED_BG
690       .shareface_0.ForeColor = CLR_DISABLED_FG
700       .shareface_1.ForeColor = CLR_DISABLED_FG
710       .shareface_2.ForeColor = CLR_DISABLED_FG
720       .shareface_3.ForeColor = CLR_DISABLED_FG
730       .shareface_4.ForeColor = CLR_DISABLED_FG
740       .shareface_box.BackColor = CLR_DISABLED_BG
750       .totdesc.ForeColor = CLR_DISABLED_FG
760       .totdesc.BackColor = CLR_DISABLED_BG
770       .ICash.ForeColor = CLR_DISABLED_FG
780       .ICash.BackColor = CLR_DISABLED_BG
790       .PCash.ForeColor = CLR_DISABLED_FG
800       .PCash.BackColor = CLR_DISABLED_BG
810       .Cost.ForeColor = CLR_DISABLED_FG
820       .Cost.BackColor = CLR_DISABLED_BG
830       .curr_id.ForeColor = CLR_DISABLED_FG
840       .curr_code.BackColor = CLR_DISABLED_BG

850     End With

EXITP:
860     Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1010    intShiftDown = (Shift And acShiftMask) > 0
1020    intAltDown = (Shift And acAltMask) > 0
1030    intCtrlDown = (Shift And acCtrlMask) > 0

1040    intRetVal = KeyCode

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   ledger_HIDDEN:    D {ledger_HIDDEN}
        ' **   ledger_HIDDEN:    N {ledger_HIDDEN on frmAccountHideTrans2_Sub_Pick}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyUp
1070        intRetVal = 0
1080        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1090      Case vbKeyDown
1100        intRetVal = 0
1110        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1120      End Select
1130    End If

        ' ** Alt keys.
1140    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1150      Select Case intRetVal
          Case vbKeyD
1160        With Me
1170          intRetVal = 0
1180          .ledger_HIDDEN.SetFocus  ' ** Don't invoke it!
1190        End With
1200      Case vbKeyN
1210        With Me
1220          intRetVal = 0
1230          If .Parent.running_pick_cnt > 0& Then
1240            DoCmd.SelectObject acForm, .Parent.Name
1250            .Parent.frmAccountHideTrans2_Sub_Pick.SetFocus
1260            .Parent.frmAccountHideTrans2_Sub_Pick.Form.ledger_HIDDEN.SetFocus  ' ** Don't invoke it!
1270          Else
1280            Beep
1290          End If
1300        End With
1310      Case vbKeyX
1320        With Me
1330          intRetVal = 0
1340          DoCmd.SelectObject acForm, .Parent.Name, False
1350          .Parent.cmdClose.SetFocus
1360          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountHideTrans2
1370        End With
1380      End Select
1390    End If

        ' ** Ctrl keys.
1400    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1410      Select Case intRetVal
          Case vbKeyS
1420        intRetVal = 0
1430        cmdSave_Click  ' ** Procedure: Below.
1440      Case vbKeyTab, vbKeyReturn
1450        With Me
1460          intRetVal = 0
1470          lngRecsCur = .Parent.frmAccountHideTrans2_Sub_Pick.Form.RecCnt
1480          If lngRecsCur > 0& Then
1490            DoCmd.SelectObject acForm, .Parent.Name, False
1500            .Parent.frmAccountHideTrans2_Sub_Pick.SetFocus
1510          Else
1520            DoCmd.SelectObject acForm, .Parent.Name, False
1530            If .Parent.chkIncludeArchive.Enabled = True Then
1540              .Parent.chkIncludeArchive.SetFocus
1550            ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
1560              .Parent.chkIncludeCurrency.SetFocus
1570            ElseIf .Parent.chkShowAll.Enabled = True Then
1580              .Parent.chkShowAll.SetFocus
1590            ElseIf .Parent.chkShowJournalNo.Enabled = True Then
1600              .Parent.chkShowJournalNo.SetFocus
1610            ElseIf .Parent.chkShowHiddenOnly.Enabled = True Then
1620              .Parent.chkShowHiddenOnly.SetFocus
1630            ElseIf .Parent.cmdPrintReport.Enabled = True Then
1640              .Parent.cmdPrintReport.SetFocus
1650            ElseIf .Parent.cmdHidden.Enabled = True Then
1660              .Parent.cmdHidden.SetFocus
1670            Else
1680              .Parent.cmdClose.SetFocus
1690            End If
1700          End If
1710        End With
1720      Case vbKeyUp, vbKeyPageUp, vbKeyHome
1730        intRetVal = 0
1740        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1750      Case vbKeyDown, vbKeyPageDown, vbKeyEnd
1760        intRetVal = 0
1770        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1780      End Select
1790    End If

        ' ** Ctrl-Shift keys.
1800    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1810      Select Case intRetVal
          Case vbKeyF
1820        With Me
1830          intRetVal = 0
1840          DoCmd.SelectObject acForm, .Parent.Name, False
1850          .Parent.FocusHolder.SetFocus
1860        End With
1870      Case vbKeyTab, vbKeyReturn
1880        With Me
1890          intRetVal = 0
1900          DoCmd.SelectObject acForm, .Parent.Name
1910          .Parent.cmdClose.SetFocus
1920        End With
1930      End Select
1940    End If

EXITP:
1950    KeyCode = intRetVal
1960    Exit Sub

ERRH:
1970    intRetVal = 0
1980    Select Case ERR.Number
        Case Else
1990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2000    End Select
2010    Resume EXITP

End Sub

Private Sub Form_Timer()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

2110    With Me
2120      .TimerInterval = 0&
2130      If blnUnHide = True Then
2140        blnUnHide = False
2150        UnHideTrans  ' ** Procedure: Below.
2160      End If
2170    End With

EXITP:
2180    Exit Sub

ERRH:
2190    Select Case ERR.Number
        Case Else
2200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2210    End Select
2220    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2310    DoCmd.RunCommand acCmdSaveRecord
2320    DoEvents

EXITP:
2330    Exit Sub

ERRH:
2340    Select Case ERR.Number
        Case Else
2350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2360    End Select
2370    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
2410    intShiftDown = (Shift And acShiftMask) > 0
2420    intAltDown = (Shift And acAltMask) > 0
2430    intCtrlDown = (Shift And acCtrlMask) > 0

2440    intRetVal = KeyCode

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2470        With Me
2480          intRetVal = 0
2490          lngRecsCur = RecCnt  ' ** Function: Below.
2500          If lngRecsCur > 0& Then
2510            .ledger_HIDDEN.SetFocus
2520          Else
2530            DoCmd.SelectObject acForm, .Parent.Name, False
2540            .Parent.cmdClose.SetFocus
2550          End If
2560        End With
2570      End Select
2580    End If

        ' ** Shift keys.
2590    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2600      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2610        With Me
2620          intRetVal = 0
2630          lngRecsCur = RecCnt  ' ** Function: Below.
2640          If lngRecsCur > 0& Then
2650            .ledger_HIDDEN.SetFocus
2660          Else
2670            DoCmd.SelectObject acForm, .Parent.Name, False
2680            .Parent.cmdClose.SetFocus
2690          End If
2700        End With
2710      End Select
2720    End If

EXITP:
2730    KeyCode = intRetVal
2740    Exit Sub

ERRH:
2750    intRetVal = 0
2760    Select Case ERR.Number
        Case Else
2770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2780    End Select
2790    Resume EXITP

End Sub

Private Sub journalno_lbl_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_lbl_DblClick"

2810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

2910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2920    Exit Sub

ERRH:
2930    Select Case ERR.Number
        Case Else
2940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950    End Select
2960    Resume EXITP

End Sub

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

3010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

3110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

3210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

3310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

3410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

3510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

3610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_lbl_DblClick"

3710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_BeforeUpdate(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_BeforeUpdate"

3810    With Me
3820      lngJournalNo = .journalno
3830      blnCancelUpdate = False
3840      If .ledger_HIDDEN = False Then  ' ** That is, if it was just changed from True to False.
3850        lngRecsCur = .Parent.frmAccountHideTrans2_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_Sub_Pick.
3860        If lngRecsCur = 0& Then
              ' ** Check the running count in case it's lying!  'WELL, WHAT IF IT'S TELLING THE TRUTH
3870          If .Parent.running_pick_cnt > 0& Then             'AND THIS ONE'S LYING?
3880            lngRecsCur = .Parent.running_pick_cnt
3890          End If
3900        End If
3910        If lngRecsCur > 0& Then
3920          If .Parent.frmAccountHideTrans2_Sub_Pick.Form.IsPicked(lngJournalNo) = False Then
                ' ** Don't let them disturb a different group while there's still one pending.
3930            .Undo
3940            DoEvents
3950            blnCancelUpdate = True
3960            Beep
3970            MsgBox "Please complete the pending hidden transactions" & vbCrLf & _
                  "before unhiding another.", vbInformation + vbOKOnly, "Pending Hidden Transactions Present"
3980          ElseIf .Parent.cmdUnHide.Visible = True Then
                ' ** Do all unhiding in the pending section only.
3990            .Undo
4000            DoEvents
4010            blnCancelUpdate = True
4020            Beep
4030            MsgBox "Transactions may only be unhidden as a group." & vbCrLf & vbCrLf & _
                  "Use the UnHide button below.", vbInformation + vbOKOnly, "Unhide Transaction Group Only"
4040          End If
4050        Else
              ' ** This would be a previously hidden one.
4060          .Undo
4070          blnCancelUpdate = True
4080          DoCmd.Hourglass True
4090          DoEvents
4100          blnUnHide = True
4110          .TimerInterval = 250&
4120        End If
4130      Else
4140        If IsNull(.Parent.curr_id) = True Then  ' ** That means this is the 1st.
              ' ** That's OK.
4150        Else
4160          If .Parent.curr_id = 0 Then  ' ** That means this is the 1st.
                ' ** That's OK.
4170          ElseIf .curr_id = .Parent.curr_id Then
                ' ** That's OK, too.
4180          Else
4190            Cancel = -1
4200            DoEvents
4210            blnCancelUpdate = True
4220            Beep
4230            MsgBox "Currency does not match.", vbInformation + vbOKOnly, "Invalid Entry"
4240          End If
4250        End If
4260        If blnCancelUpdate = False Then
4270          If .assetno = .Parent.assetno Then
                ' ** That's OK.
4280          ElseIf .assetno = 0& Or .Parent.assetno = 0& Then  ' ** If Parent.assetno = 0, that means this is the 1st.
                ' ** That's OK, too.
4290          Else
4300            Cancel = -1
4310            DoEvents
4320            blnCancelUpdate = True
4330            Beep
4340            MsgBox "Assets do not match.", vbInformation + vbOKOnly, "Invalid Entry"
4350          End If
4360        End If  ' ** blnCancelUpdate.
4370      End If  ' ** ledger_HIDDEN.
4380    End With

EXITP:
4390    Exit Sub

ERRH:
4400    DoCmd.Hourglass False
4410    Select Case ERR.Number
        Case Else
4420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4430    End Select
4440    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_AfterUpdate()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngOrder As Long, lngCurrID As Long
        Dim strCurrSym As String, strFormat As String
        Dim varTmp00 As Variant, strTmp01 As String

4510    With Me
4520      If blnCancelUpdate = False Then
4530        lngJournalNo = .journalno
4540        Select Case .ledger_HIDDEN
            Case True  ' ** Clicked from False to True.
4550          DoCmd.Hourglass True
4560          DoEvents
4570          varTmp00 = DCount("*", "tmpAccountHideTrans_03")
4580          lngRecsCur = Nz(varTmp00, 0)
4590          lngOrder = lngRecsCur + 1&
4600          Select Case .aht2_prehidden
              Case True
4610            .aht2_changed = False
4620          Case False
4630            .aht2_changed = True
4640          End Select
4650          cmdSave_Click  ' ** Procedure: Above.
4660          Set dbs = CurrentDb
4670          With dbs
                ' ** Append tmpAccountHideTrans_02 to tmpAccountHideTrans_03, by specified [jno], [ord].
4680            Set qdf = .QueryDefs("qryAccountHideTrans2_14")
4690            With qdf.Parameters
4700              ![jno] = lngJournalNo
4710              ![ord] = lngOrder
4720            End With
4730            qdf.Execute
4740            Set qdf = Nothing
4750            .Close
4760          End With
4770          .Parent.frmAccountHideTrans2_Sub_Pick.Form.Requery
4780          strTmp01 = .Parent.frmAccountHideTrans2_Sub_Pick.Form.OrderBy
4790          If strTmp01 = vbNullString Then
4800            .Parent.frmAccountHideTrans2_Sub_Pick.Form.SortNow "Form_Load"  ' ** Form Procedure: frmAccountHideTrans2_Sub_Pick.
4810          Else
4820            If Left(strTmp01, 5) = "[journalno]" Then
4830              .Parent.frmAccountHideTrans2_Sub_Pick.Form.journalno_Sort_line1.Visible = True
4840              .Parent.frmAccountHideTrans2_Sub_Pick.Form.journalno_Sort_line2.Visible = True
4850              .Parent.frmAccountHideTrans2_Sub_Pick.Form.Sort_lbl.Visible = True
4860              .Parent.frmAccountHideTrans2_Sub_Pick.Form.OrderByOn = True
4870            Else
4880              .Parent.frmAccountHideTrans2_Sub_Pick.Form.Sort_line.Visible = True
4890              .Parent.frmAccountHideTrans2_Sub_Pick.Form.Sort_lbl.Visible = True
4900              .Parent.frmAccountHideTrans2_Sub_Pick.Form.OrderByOn = True
4910            End If
4920          End If
4930          .Parent.cmdCancel.Enabled = True
4940          .Parent.cmdOK.Visible = True
4950          .Parent.cmdUnHide.Visible = False
4960          .Parent.cmdUnHide.Enabled = False
4970          If .Parent.pending_trans_lbl.Caption = "Pending Matches" Then
4980            .Parent.pending_trans_lbl.Caption = "Pending Matches"
4990          End If
5000          .Parent.TotalsEnable True  ' ** Form Procedure: frmAccountHideTrans2.
5010          If lngRecsCur = 0& Then  ' ** This'll be the first.
5020            .Parent.CheckBoxLock True  ' ** Form Procedure: frmAccountHideTrans2.
5030            .Parent.unhidden_trans = 0&
5040            .Parent.assetno = .assetno  ' ** May be Zero.
5050          End If
5060          DoEvents
5070          lngCurrID = .curr_id
5080          strCurrSym = .currsym_symbol
5090          strFormat = strCurrSym & "#,##0.00;(" & strCurrSym & "#,##0.00);" & strCurrSym & "0.00;" & strCurrSym & "0.00"
5100          varTmp00 = DCount("*", "tmpAccountHideTrans_03")
5110          lngRecsCur = Nz(varTmp00, 0)
5120          .Parent.pending_trans = lngRecsCur
5130          varTmp00 = DSum("[sharefacex]", "tmpAccountHideTrans_03")
5140          .Parent.shareface_subtot = CDbl(Nz(varTmp00, 0))
5150          varTmp00 = DSum("[shareface_divint]", "tmpAccountHideTrans_03")
5160          .Parent.shareface_divint_subtot = CDbl(Nz(varTmp00, 0))
5170          varTmp00 = DSum("[icash]", "tmpAccountHideTrans_03")
5180          .Parent.icash_subtot = CCur(Nz(varTmp00, 0))
5190          .Parent.icash_subtot.Format = strFormat
5200          varTmp00 = DSum("[pcash]", "tmpAccountHideTrans_03")
5210          .Parent.pcash_subtot = CCur(Nz(varTmp00, 0))
5220          .Parent.pcash_subtot.Format = strFormat
5230          varTmp00 = DSum("[cost]", "tmpAccountHideTrans_03")
5240          .Parent.cost_subtot = CCur(Nz(varTmp00, 0))
5250          .Parent.cost_subtot.Format = strFormat
5260          .Parent.PickSubExpand True  ' ** Form Procedure: frmAccountHideTrans2.
5270          .Parent.frmAccountHideTrans2_Sub_Pick.Form.ledger_HIDDEN.Locked = False  ' ** In case it was an unhide.
5280          .Parent.running_pick_cnt = .Parent.running_pick_cnt + 1&
5290          If .Parent.journalno1 = 0& Then
5300            .Parent.journalno1 = lngJournalNo  ' ** Just save the 1st one.
5310          End If
5320          If IsNull(.Parent.curr_id) = True Then
5330            .Parent.curr_id = lngCurrID
5340          Else
5350            If .Parent.curr_id = 0& Then
5360              .Parent.curr_id = lngCurrID
5370            End If
5380          End If
5390          If lngCurrID = 150& Then
5400            .Parent.curr_id.BackColor = CLR_DISABLED_BG
5410            .Parent.icash_subtot.BackColor = CLR_WHT
5420            .Parent.pcash_subtot.BackColor = CLR_WHT
5430            .Parent.cost_subtot.BackColor = CLR_WHT
5440          Else
5450            .Parent.curr_id.BackColor = CLR_LTGRN
5460            .Parent.icash_subtot.BackColor = CLR_VLTGRN
5470            .Parent.pcash_subtot.BackColor = CLR_VLTGRN
5480            .Parent.cost_subtot.BackColor = CLR_VLTGRN
5490          End If
5500          If .assetno = 0 Then
5510            If .Parent.cnt_asset = 0 Then
5520              .Parent.shareface_subtot.Enabled = False
5530              .Parent.shareface_subtot.Locked = False
5540              .Parent.shareface_subtot.BorderColor = WIN_CLR_DISR
5550              .Parent.shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
5560              .Parent.shareface_subtot_lbl_dim_hi.Visible = True
5570            End If
5580          Else
5590            If .Parent.shareface_subtot.Locked = False Then
5600              .Parent.shareface_subtot.Locked = True
5610              .Parent.shareface_subtot.BorderColor = CLR_LTBLU2
5620              .Parent.shareface_subtot_lbl.ForeColor = CLR_DKGRY3
5630              .Parent.shareface_subtot_lbl_dim_hi.Visible = False
5640            End If
5650          End If

5660        Case False  ' ** Clicked from True to False.
5670          DoCmd.Hourglass True
5680          DoEvents
5690          Select Case .aht2_prehidden
              Case True
5700            .aht2_changed = True
5710          Case False
5720            .aht2_changed = False
5730          End Select
5740          cmdSave_Click  ' ** Procedure: Above.
5750          Set dbs = CurrentDb
5760          With dbs
                ' ** Delete tmpAccountHideTrans_03, by specified [jno].
5770            Set qdf = .QueryDefs("qryAccountHideTrans2_15")
5780            With qdf.Parameters
5790              ![jno] = lngJournalNo
5800            End With
5810            qdf.Execute dbFailOnError
5820            Set qdf = Nothing
5830            .Close
5840          End With
5850          .Parent.frmAccountHideTrans2_Sub_Pick.Form.Requery
5860          DoEvents
5870          varTmp00 = DCount("*", "tmpAccountHideTrans_03")
5880          lngRecsCur = Nz(varTmp00, 0)
5890          .Parent.pending_trans = lngRecsCur
5900          If lngRecsCur > 0& Then
5910            .Parent.shareface_subtot = .Parent.frmAccountHideTrans2_Sub_Pick.Form.shareface_subtot
5920            .Parent.shareface_divint_subtot = .Parent.frmAccountHideTrans2_Sub_Pick.Form.shareface_divint_subtot
5930            .Parent.icash_subtot = .Parent.frmAccountHideTrans2_Sub_Pick.Form.icash_subtot
5940            .Parent.pcash_subtot = .Parent.frmAccountHideTrans2_Sub_Pick.Form.pcash_subtot
5950            .Parent.cost_subtot = .Parent.frmAccountHideTrans2_Sub_Pick.Form.cost_subtot
5960            .Parent.PickSubExpand False  ' ** Form Procedure: frmAccountHideTrans2.
5970            .Parent.frmAccountHideTrans2_Sub_Pick.Form.ledger_HIDDEN.Locked = False  ' ** In case it was an unhide.
5980            strTmp01 = .Parent.frmAccountHideTrans2_Sub_Pick.Form.OrderBy
5990            If strTmp01 = vbNullString Then
6000              .Parent.frmAccountHideTrans2_Sub_Pick.Form.SortNow "Form_Load"  ' ** Form Procedure: frmAccountHideTrans2_Sub_Pick.
6010            Else
6020              If Left(strTmp01, 5) = "[journalno]" Then
6030                .Parent.frmAccountHideTrans2_Sub_Pick.Form.journalno_Sort_line1.Visible = True
6040                .Parent.frmAccountHideTrans2_Sub_Pick.Form.journalno_Sort_line2.Visible = True
6050                .Parent.frmAccountHideTrans2_Sub_Pick.Form.Sort_lbl.Visible = True
6060                .Parent.frmAccountHideTrans2_Sub_Pick.Form.OrderByOn = True
6070              Else
6080                .Parent.frmAccountHideTrans2_Sub_Pick.Form.Sort_line.Visible = True
6090                .Parent.frmAccountHideTrans2_Sub_Pick.Form.Sort_lbl.Visible = True
6100                .Parent.frmAccountHideTrans2_Sub_Pick.Form.OrderByOn = True
6110              End If
6120            End If
6130          Else
6140            strFormat = "$#,##0.00;($#,##0.00);$0.00;$0.00"
6150            .Parent.shareface_subtot = 0#
6160            .Parent.shareface_divint_subtot = 0#
6170            .Parent.icash_subtot = 0#
6180            .Parent.icash_subtot.Format = strFormat
6190            .Parent.pcash_subtot = 0#
6200            .Parent.pcash_subtot.Format = strFormat
6210            .Parent.cost_subtot = 0#
6220            .Parent.cost_subtot.Format = strFormat
6230            .Parent.curr_id = Null
6240            .Parent.PickSubExpand False  ' ** Form Procedure: frmAccountHideTrans2.
6250          End If
6260          .Parent.running_pick_cnt = .Parent.running_pick_cnt - 1&
6270          If .Parent.journalno1 = 0& Then
6280            .Parent.journalno1 = lngJournalNo  ' ** Just save the 1st one.
6290          End If
6300        End Select
6310        DoEvents
6320        varTmp00 = DCount("*", "tmpAccountHideTrans_03")
6330        lngRecsCur = Nz(varTmp00, 0)
6340        DoEvents
6350        .Parent.shareface_chk  ' ** Form Procedure: frmAccountHideTrans2.
6360        DoEvents
6370        If lngRecsCur = 0& Then
6380          .Parent.cmdOK.Enabled = False
6390          .Parent.cmdCancel.Enabled = False
6400          .Parent.TotalsEnable False  ' ** Form Procedure: frmAccountHideTrans2.
6410          .Parent.CheckBoxLock False  ' ** Form Procedure: frmAccountHideTrans2.
6420        End If
6430        .Parent.UpdateHiddenTrans  ' ** Form Procedure: frmAccountHideTrans2.
6440        .Parent.UpdateTotals_Set True, 500&  ' ** Form Procedure: frmAccountHideTrans2.
6450      Else
6460        blnCancelUpdate = False
6470      End If  ' ** blnCancelUpdate.
6480    End With

6490    DoCmd.Hourglass False

EXITP:
6500    Set qdf = Nothing
6510    Set dbs = Nothing
6520    Exit Sub

ERRH:
6530    DoCmd.Hourglass False
6540    Select Case ERR.Number
        Case Else
6550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6560    End Select
6570    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
6610    intShiftDown = (Shift And acShiftMask) > 0
6620    intAltDown = (Shift And acAltMask) > 0
6630    intCtrlDown = (Shift And acCtrlMask) > 0

6640    intRetVal = KeyCode

        ' ** Plain keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6670        With Me
6680          intRetVal = 0
6690          lngRecsCur = RecCnt  ' ** Function: Below.
6700          If .CurrentRecord < lngRecsCur Then
6710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
6720            .ledger_HIDDEN.SetFocus
6730          Else
6740            lngRecsCur = .Parent.frmAccountHideTrans2_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_Sub_Pick.
6750            DoCmd.SelectObject acForm, .Parent.Name, False
6760            If lngRecsCur > 0& Then
6770              .Parent.frmAccountHideTrans2_Sub_Pick.SetFocus
6780              .Parent.frmAccountHideTrans2_Sub_Pick.Form.ledger_HIDDEN.SetFocus
6790            Else
6800              If .Parent.chkIncludeArchive.Enabled = True Then
6810                .Parent.chkIncludeArchive.SetFocus
6820              ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
6830                .Parent.chkIncludeCurrency.SetFocus
6840              ElseIf .Parent.chkShowAll.Enabled = True Then
6850                .Parent.chkShowAll.SetFocus
6860              ElseIf .Parent.chkShowJournalNo.Enabled = True Then
6870                .Parent.chkShowJournalNo.SetFocus
6880              ElseIf .Parent.chkShowHiddenOnly.Enabled = True Then
6890                .Parent.chkShowHiddenOnly.SetFocus
6900              ElseIf .Parent.cmdPrintReport.Enabled = True Then
6910                .Parent.cmdPrintReport.SetFocus
6920              ElseIf .Parent.cmdHidden.Enabled = True Then
6930                .Parent.cmdHidden.SetFocus
6940              Else
6950                .Parent.cmdClose.SetFocus
6960              End If
6970            End If
6980          End If
6990        End With
7000      End Select
7010    End If

        ' ** Shift keys.
7020    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7030      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7040        With Me
7050          intRetVal = 0
7060          If .CurrentRecord > 1 Then
7070            MoveRec acCmdRecordsGoToPrevious
7080            .ledger_HIDDEN.SetFocus
7090          Else
7100            DoCmd.SelectObject acForm, .Parent.Name, False
7110            .Parent.cmdClose.SetFocus
7120          End If
7130        End With
7140      End Select
7150    End If

EXITP:
7160    KeyCode = intRetVal
7170    Exit Sub

ERRH:
7180    intRetVal = 0
7190    Select Case ERR.Number
        Case Else
7200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7210    End Select
7220    Resume EXITP

End Sub

Private Sub UnHideTrans()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "UnHideTrans"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCurrID As Long, strCurrSym As String, strFormat As String
        Dim blnContinue As Boolean, blnCurrID As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long

7310    blnContinue = True

7320    With Me
7330      If lngJournalNo > 0& Then

7340        DoCmd.Hourglass True
7350        DoEvents

7360        .Parent.unhidden_trans = 0&
7370        blnCurrID = True

7380        varTmp00 = DCount("*", "tblLedgerHidden")
7390        If IsNull(varTmp00) = True Then
7400          blnContinue = False
7410        Else
7420          If varTmp00 = 0 Then
7430            blnContinue = False
7440          End If
7450        End If

7460        If blnContinue = True Then

7470          Set dbs = CurrentDb

              ' ** qryAccountHideTrans2_26b (qryAccountHideTrans2_26a (tblLedgerHidden, by specified [jno]),
              ' ** linked back to tblLedgerHidden, for the entire group), linked to tmpAccountHideTrans_02
7480          Set qdf = dbs.QueryDefs("qryAccountHideTrans2_26c")
7490          With qdf.Parameters
7500            ![jno] = lngJournalNo
7510          End With
7520          Set rst = qdf.OpenRecordset
7530          If rst.BOF = True And rst.EOF = True Then
7540            blnContinue = False
7550            rst.Close
7560            dbs.Close
7570          Else
7580            rst.Close
7590            Set rst = Nothing
7600            Set qdf = Nothing
7610          End If
7620        End If  ' ** blnContinue.

7630        If blnContinue = True Then

7640          With dbs
                ' ** Append qryAccountHideTrans2_26c (qryAccountHideTrans2_26b (qryAccountHideTrans2_26a
                ' ** (tblLedgerHidden, by specified [jno]), linked back to tblLedgerHidden, for the
                ' ** entire group), linked to tmpAccountHideTrans_02) to tmpAccountHideTrans_03.  #curr_id
7650            Set qdf = .QueryDefs("qryAccountHideTrans2_26d")
7660            With qdf.Parameters
7670              ![jno] = lngJournalNo
7680            End With
7690            qdf.Execute dbFailOnError
7700            .Close
7710          End With
7720          .Parent.frmAccountHideTrans2_Sub_Pick.Form.Requery
7730          DoEvents

7740          .Parent.cmdCancel.Enabled = True
7750          .Parent.cmdOK.Visible = False
7760          .Parent.cmdUnHide.Visible = True
7770          .Parent.cmdUnHide.Enabled = True
7780          .Parent.pending_trans_lbl.Caption = "Pending Reveals"
7790          .Parent.TotalsEnable True  ' ** Form Procedure: frmAccountHideTrans2.
7800          .Parent.CheckBoxLock True  ' ** Form Procedure: frmAccountHideTrans2.
7810          DoEvents

7820          lngCurrID = .curr_id
7830          strCurrSym = .currsym_symbol
7840          strFormat = strCurrSym & "#,##0.00;(" & strCurrSym & "#,##0.00);" & strCurrSym & "0.00;" & strCurrSym & "0.00"

7850          varTmp00 = DCount("*", "tmpAccountHideTrans_03")
7860          lngRecsCur = Nz(varTmp00, 0)
7870          .Parent.pending_trans = lngRecsCur
7880          varTmp00 = DSum("[sharefacex]", "tmpAccountHideTrans_03")
7890          .Parent.shareface_subtot = CDbl(Nz(varTmp00, 0))
7900          varTmp00 = DSum("[shareface_divint]", "tmpAccountHideTrans_03")
7910          .Parent.shareface_divint_subtot = CDbl(Nz(varTmp00, 0))
7920          varTmp00 = DSum("[icash]", "tmpAccountHideTrans_03")
7930          .Parent.icash_subtot = CCur(Nz(varTmp00, 0))
7940          .Parent.icash_subtot.Format = strFormat
7950          varTmp00 = DSum("[pcash]", "tmpAccountHideTrans_03")
7960          .Parent.pcash_subtot = CCur(Nz(varTmp00, 0))
7970          .Parent.pcash_subtot.Format = strFormat
7980          varTmp00 = DSum("[cost]", "tmpAccountHideTrans_03")
7990          .Parent.cost_subtot = CCur(Nz(varTmp00, 0))
8000          .Parent.cost_subtot.Format = strFormat
8010          .Parent.curr_id = lngCurrID
8020          varTmp00 = DCount("*", "tmpAccountHideTrans_03", "[curr_id] = " & CStr(lngCurrID))
8030          If varTmp00 <> lngRecsCur Then
8040            blnCurrID = False
8050          End If
8060          DoEvents

8070          If .Parent.chkIncludeCurrency = True Then
8080            If lngCurrID = 150& Then
8090              .Parent.curr_id.BackColor = CLR_DISABLED_BG
8100              .Parent.icash_subtot.BackColor = CLR_WHT
8110              .Parent.pcash_subtot.BackColor = CLR_WHT
8120              .Parent.cost_subtot.BackColor = CLR_WHT
8130            Else
8140              .Parent.curr_id.BackColor = CLR_LTGRN
8150              .Parent.icash_subtot.BackColor = CLR_VLTGRN
8160              .Parent.pcash_subtot.BackColor = CLR_VLTGRN
8170              .Parent.cost_subtot.BackColor = CLR_VLTGRN
8180            End If
8190          End If

8200          If .Parent.cnt_asset = 0 Then
8210            .Parent.shareface_subtot.Enabled = False
8220            .Parent.shareface_subtot.Locked = False
8230            .Parent.shareface_subtot.BorderColor = WIN_CLR_DISR
8240            .Parent.shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
8250            .Parent.shareface_subtot_lbl_dim_hi.Visible = True
8260          End If

              ' ** Save the number of transactions in the group.
8270          .Parent.unhidden_trans = lngRecsCur
8280          .Parent.frmAccountHideTrans2_Sub_Pick.Form.FocusHolder.SetFocus
8290          .Parent.frmAccountHideTrans2_Sub_Pick.Form.ledger_HIDDEN.Locked = True  ' ** As a group only!

8300          For lngX = 1& To lngRecsCur
8310            .Parent.PickSubExpand True  ' ** Form Procedure: frmAccountHideTrans2.
8320          Next

8330          .Parent.shareface_chk  ' ** Form Procedure: frmAccountHideTrans2.
8340          DoEvents

              ' ** They should all be Zeroes.
8350          If .Parent.shareface_subtot = 0 And .Parent.icash_subtot = 0 And .Parent.pcash_subtot = 0 And .Parent.cost_subtot = 0 Then
                ' ** Matching curr_id's won't be considered when unhiding.
8360            If .Parent.cnt_asset > 0 Then
8370              .Parent.shareface_subtot_box1.Visible = True
8380              .Parent.shareface_subtot_box2.Visible = True
8390            Else
8400              .Parent.shareface_subtot_box1.Visible = False
8410              .Parent.shareface_subtot_box2.Visible = False
8420            End If
8430            .Parent.icash_subtot_box1.Visible = True
8440            .Parent.icash_subtot_box2.Visible = True
8450            .Parent.pcash_subtot_box1.Visible = True
8460            .Parent.pcash_subtot_box2.Visible = True
8470            .Parent.cost_subtot_box1.Visible = True
8480            .Parent.cost_subtot_box2.Visible = True
8490            If .Parent.chkIncludeCurrency = True Then
8500              If blnCurrID = True Then
8510                .Parent.curr_id_box1.Visible = True
8520                .Parent.curr_id_box2.Visible = True
8530              End If
8540            End If
8550          End If

8560          DoCmd.SelectObject acForm, .Parent.Name, False
8570          .Parent.cmdUnHide.SetFocus

8580        Else
8590          Beep
8600          MsgBox "tblLedgerHidden EMPTY!"
8610        End If  ' ** blnContinue.
8620      End If
8630    End With

8640    DoCmd.Hourglass False

EXITP:
8650    Set rst = Nothing
8660    Set qdf = Nothing
8670    Set dbs = Nothing
8680    Exit Sub

ERRH:
8690    DoCmd.Hourglass False
8700    Select Case ERR.Number
        Case Else
8710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8720    End Select
8730    Resume EXITP

End Sub

Public Sub ShowJournalNo_Sub(blnShow As Boolean)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo_Sub"

        Dim blnCurrency As Boolean, blnSortHere As Boolean, blnResort As Boolean

8810    With Me

8820      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
8830        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
8840      End If

8850      blnCurrency = .Parent.chkIncludeCurrency
8860      blnSortHere = False

8870      Select Case blnShow
          Case True

8880        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
8890        .totdesc.Left = lngTotDesc_Left
8900        .totdesc_lbl.Left = lngTotDesc_Left
8910        .totdesc_lbl_line.Left = lngTotDesc_Left
8920        Select Case blnCurrency
            Case True
8930          .totdesc.Width = lngTotDesc_Width
8940          .totdesc_lbl.Width = .totdesc.Width
8950          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
8960        Case False
8970          .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)
8980          .totdesc_lbl.Width = .totdesc.Width
8990          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
9000        End Select
9010        If blnSortHere = True Then
9020          .Sort_line.Left = .totdesc_lbl.Left
9030          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
9040          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
9050        End If
9060        blnSortHere = False
9070        If .Sort_line.Left = .journaltype_lbl.Left Then blnSortHere = True
9080        .journaltype.Left = lngJournalType_Left
9090        .journaltype_lbl.Left = lngJournalType_Left
9100        .journaltype_lbl_line.Left = lngJournalType_Left
9110        If blnSortHere = True Then
9120          .Sort_line.Left = .journaltype_lbl.Left
9130          .Sort_lbl.Left = ((.journaltype_lbl.Left + .journaltype_lbl.Width) - lngSortLbl_Width)
9140        End If
9150        blnSortHere = False
9160        If .Sort_line.Left = .transdate_lbl.Left Then blnSortHere = True
9170        .transdate.Left = lngTransdate_Left
9180        .transdate_lbl.Left = lngTransdate_Left
9190        .transdate_lbl_line.Left = lngTransdate_Left
9200        If blnSortHere = True Then
9210          .Sort_line.Left = .transdate_lbl.Left
9220          .Sort_lbl.Left = ((.transdate_lbl.Left + .transdate_lbl.Width) - lngSortLbl_Width)
9230        End If
9240        blnSortHere = False
9250        .journalno.Visible = True
9260        .journalno_lbl.Visible = True
9270        .journalno_lbl2.Visible = True
9280        .journalno_lbl_line.Visible = True

9290      Case False

9300        blnResort = False
9310        If .journalno_Sort_line1.Visible = True Then blnSortHere = True
9320        .journalno.Visible = False
9330        .journalno_lbl.Visible = False
9340        .journalno_lbl2.Visible = False
9350        .journalno_lbl_line.Visible = False
9360        .journalno_Sort_line1.Visible = False
9370        .journalno_Sort_line2.Visible = False
9380        If blnSortHere = True Then
9390          blnResort = True
9400          .Sort_lbl.Visible = False
9410        End If
9420        blnSortHere = False
9430        If .Sort_line.Left = .transdate_lbl.Left Then blnSortHere = True
9440        .transdate.Left = (lngTransdate_Left - lngJournalNo_Width)
9450        .transdate_lbl.Left = .transdate.Left
9460        .transdate_lbl_line.Left = .transdate_lbl.Left
9470        If blnSortHere = True Then
9480          .Sort_line.Left = .transdate_lbl.Left
9490          .Sort_lbl.Left = ((.transdate_lbl.Left + .transdate_lbl.Width) - lngSortLbl_Width)
9500        End If
9510        blnSortHere = False
9520        If .Sort_line.Left = .journaltype_lbl.Left Then blnSortHere = True
9530        .journaltype.Left = (lngJournalType_Left - lngJournalNo_Width)
9540        .journaltype_lbl.Left = .journaltype.Left
9550        .journaltype_lbl_line.Left = .journaltype_lbl.Left
9560        If blnSortHere = True Then
9570          .Sort_line.Left = .journaltype_lbl.Left
9580          .Sort_lbl.Left = ((.journaltype_lbl.Left + .journaltype_lbl.Width) - lngSortLbl_Width)
9590        End If
9600        blnSortHere = False
9610        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
9620        .totdesc.Left = (lngTotDesc_Left - lngJournalNo_Width)
9630        .totdesc_lbl.Left = .totdesc.Left
9640        .totdesc_lbl_line.Left = .totdesc_lbl.Left
9650        Select Case blnCurrency
            Case True
9660          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)
9670          .totdesc_lbl.Width = .totdesc.Width
9680          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
9690        Case False
9700          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width + lngCurrID_Width)
9710          .totdesc_lbl.Width = .totdesc.Width
9720          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
9730        End Select
9740        If blnSortHere = True Then
9750          .Sort_line.Left = .totdesc_lbl.Left
9760          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
9770          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
9780        End If
9790        blnSortHere = False

9800        If blnResort = True Then
9810          SortNow "Form_Load"  ' ** Procedure: Below.
9820        End If

9830      End Select
9840    End With

EXITP:
9850    Exit Sub

ERRH:
9860    Select Case ERR.Number
        Case Else
9870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9880    End Select
9890    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnInclude As Boolean)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnJournalNo As Boolean, blnSortHere As Boolean, blnResort As Boolean

9910    With Me

9920      DoCmd.Hourglass True
9930      DoEvents

9940      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
9950        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
9960      End If

9970      blnJournalNo = .Parent.chkShowJournalNo
9980      blnSortHere = False

9990      Select Case blnInclude
          Case True

10000       If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
10010       Select Case blnJournalNo
            Case True
10020         .totdesc.Width = lngTotDesc_Width
10030         .totdesc_lbl.Width = .totdesc.Width
10040         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
10050       Case False
10060         .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)
10070         .totdesc_lbl.Width = .totdesc.Width
10080         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
10090       End Select
10100       If blnSortHere = True Then
10110         .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
10120         .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
10130       End If
10140       blnSortHere = False
10150       If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
10160       .shareface_0.Left = lngShareface_Left
10170       .shareface_1.Left = .shareface_0.Left
10180       .shareface_2.Left = .shareface_0.Left
10190       .shareface_3.Left = .shareface_0.Left
10200       .shareface_4.Left = .shareface_0.Left
10210       .shareface_box.Left = .shareface_0.Left
10220       .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
10230       .shareface_lbl_line.Left = .shareface_lbl.Left
10240       If blnSortHere = True Then
10250         .Sort_line.Left = .shareface_lbl.Left
10260         .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
10270       End If
10280       blnSortHere = False
10290       If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
10300       .ICash.Left = lngICash_Left
10310       .icash_str.Left = .ICash.Left
10320       .icash_forex.Left = (.icash_str.Left - lngTpp)
10330       .icash_box.Left = .ICash.Left
10340       .icash_str.Visible = True
10350       .ICash.Visible = False
10360       .icash_forex.Visible = True
10370       .icash_lbl.Left = (.ICash.Left - lngICashLbl_Offset)
10380       .icash_lbl_line.Left = .icash_lbl.Left
10390       If blnSortHere = True Then
10400         .Sort_line.Left = .icash_lbl.Left
10410         .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
10420       End If
10430       blnSortHere = False
10440       If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
10450       .PCash.Left = lngPCash_Left
10460       .pcash_str.Left = .PCash.Left
10470       .pcash_forex.Left = (.pcash_str.Left - lngTpp)
10480       .pcash_box.Left = .PCash.Left
10490       .pcash_str.Visible = True
10500       .PCash.Visible = False
10510       .pcash_forex.Visible = True
10520       .pcash_lbl.Left = (.PCash.Left - lngICashLbl_Offset)
10530       .pcash_lbl_line.Left = .pcash_lbl.Left
10540       If blnSortHere = True Then
10550         .Sort_line.Left = .pcash_lbl.Left
10560         .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
10570       End If
10580       blnSortHere = False
10590       If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
10600       .Cost.Left = lngCost_Left
10610       .cost_str.Left = .Cost.Left
10620       .cost_forex.Left = (.cost_str.Left - lngTpp)
10630       .cost_box.Left = .Cost.Left
10640       .cost_str.Visible = True
10650       .Cost.Visible = False
10660       .cost_forex.Visible = True
10670       .cost_lbl.Left = (.Cost.Left - lngICashLbl_Offset)
10680       .cost_lbl_line.Left = .cost_lbl.Left
10690       If blnSortHere = True Then
10700         .Sort_line.Left = .cost_lbl.Left
10710         .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
10720       End If
10730       blnSortHere = False
10740       .curr_id.Left = lngCurrID_Left
10750       .curr_id_forex.Left = (.curr_id.Left - lngTpp)
10760       .curr_id_box.Left = .curr_id.Left
10770       .curr_id_lbl.Left = (lngCurrID_Left - lngCurrIDLbl_Offset)
10780       .curr_id_lbl_line.Left = lngCurrID_Left
10790       .curr_id.Visible = True
10800       .curr_id_forex.Visible = True
10810       .curr_id_box.Visible = True
10820       .curr_id_lbl.Visible = True
10830       .curr_id_lbl_line.Visible = True

10840     Case False

10850       blnResort = False
10860       Select Case blnJournalNo
            Case True
10870         If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
10880         .curr_id.Visible = False
10890         .curr_id_forex.Visible = False
10900         .curr_id_box.Visible = False
10910         .curr_id_lbl.Visible = False
10920         .curr_id_lbl_line.Visible = False
10930         .curr_id.Left = 0&
10940         .curr_id_forex.Left = 0&
10950         .curr_id_box.Left = 0&
10960         .curr_id_lbl.Left = 0&
10970         .curr_id_lbl_line.Left = 0&
10980         If blnSortHere = True Then
10990           blnResort = True
11000           .Sort_line.Visible = False
11010           .Sort_lbl.Visible = False
11020         End If
11030         blnSortHere = False
11040         If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
11050         .Cost.Left = (lngCost_Left + lngCurrID_Width)
11060         .cost_str.Left = .Cost.Left
11070         .cost_forex.Left = (.cost_str.Left - lngTpp)
11080         .cost_box.Left = .Cost.Left
11090         Select Case gblnHasForExThis
              Case True
11100           .cost_str.Visible = True
11110           .Cost.Visible = False
11120           .cost_forex.Visible = True
11130         Case False
11140           Select Case .Parent.chkIncludeCurrency
                Case True
11150             .cost_str.Visible = True
11160             .Cost.Visible = False
11170             .cost_forex.Visible = True
11180           Case False
11190             .Cost.Visible = True
11200             .cost_str.Visible = False
11210             .cost_forex.Visible = False
11220           End Select
11230         End Select
11240         .cost_lbl.Left = (.Cost.Left - lngICashLbl_Offset)
11250         .cost_lbl_line.Left = .cost_lbl.Left
11260         If blnSortHere = True Then
11270           .Sort_line.Left = .cost_lbl.Left
11280           .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
11290         End If
11300         blnSortHere = False
11310         If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
11320         .PCash.Left = (lngPCash_Left + lngCurrID_Width)
11330         .pcash_str.Left = .PCash.Left
11340         .pcash_forex.Left = (.pcash_str.Left - lngTpp)
11350         .pcash_box.Left = .PCash.Left
11360         Select Case gblnHasForExThis
              Case True
11370           .pcash_str.Visible = True
11380           .PCash.Visible = False
11390           .pcash_forex.Visible = True
11400         Case False
11410           Select Case .Parent.chkIncludeCurrency
                Case True
11420             .pcash_str.Visible = True
11430             .PCash.Visible = False
11440             .pcash_forex.Visible = True
11450           Case False
11460             .PCash.Visible = True
11470             .pcash_str.Visible = False
11480             .pcash_forex.Visible = False
11490           End Select
11500         End Select
11510         .pcash_lbl.Left = (.PCash.Left - lngICashLbl_Offset)
11520         .pcash_lbl_line.Left = .pcash_lbl.Left
11530         If blnSortHere = True Then
11540           .Sort_line.Left = .pcash_lbl.Left
11550           .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
11560         End If
11570         blnSortHere = False
11580         If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
11590         .ICash.Left = (lngICash_Left + lngCurrID_Width)
11600         .icash_str.Left = .ICash.Left
11610         .icash_forex.Left = (.icash_str.Left - lngTpp)
11620         .icash_box.Left = .ICash.Left
11630         Select Case gblnHasForExThis
              Case True
11640           .icash_str.Visible = True
11650           .ICash.Visible = False
11660           .icash_forex.Visible = True
11670         Case False
11680           Select Case .Parent.chkIncludeCurrency
                Case True
11690             .icash_str.Visible = True
11700             .ICash.Visible = False
11710             .icash_forex.Visible = True
11720           Case False
11730             .ICash.Visible = True
11740             .icash_str.Visible = False
11750             .icash_forex.Visible = False
11760           End Select
11770         End Select
11780         .icash_lbl.Left = (.ICash.Left - lngICashLbl_Offset)
11790         .icash_lbl_line.Left = .icash_lbl.Left
11800         If blnSortHere = True Then
11810           .Sort_line.Left = .icash_lbl.Left
11820           .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
11830         End If
11840         blnSortHere = False
11850         If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
11860         .shareface_0.Left = (lngShareface_Left + lngCurrID_Width)
11870         .shareface_1.Left = .shareface_0.Left
11880         .shareface_2.Left = .shareface_0.Left
11890         .shareface_3.Left = .shareface_0.Left
11900         .shareface_4.Left = .shareface_0.Left
11910         .shareface_box.Left = .shareface_0.Left
11920         .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
11930         .shareface_lbl_line.Left = .shareface_lbl.Left
11940         If blnSortHere = True Then
11950           .Sort_line.Left = .shareface_lbl.Left
11960           .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
11970         End If
11980         blnSortHere = False
11990         If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
12000         .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)
12010         .totdesc_lbl.Width = .totdesc.Width
12020         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
12030         If blnSortHere = True Then
12040           .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
12050           .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
12060         End If
12070         blnSortHere = False
12080       Case False
12090         If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
12100         .curr_id.Visible = False
12110         .curr_id_forex.Visible = False
12120         .curr_id_box.Visible = False
12130         .curr_id_lbl.Visible = False
12140         .curr_id_lbl_line.Visible = False
12150         .curr_id.Left = 0&
12160         .curr_id_forex.Left = 0&
12170         .curr_id_box.Left = 0&
12180         .curr_id_lbl.Left = 0&
12190         .curr_id_lbl_line.Left = 0&
12200         If blnSortHere = True Then
12210           blnResort = True
12220           .Sort_line.Visible = False
12230           .Sort_lbl.Visible = False
12240         End If
12250         blnSortHere = False
12260         If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
12270         .Cost.Left = (lngCost_Left + lngCurrID_Width)
12280         .cost_str.Left = .Cost.Left
12290         .cost_forex.Left = (.cost_str.Left - lngTpp)
12300         .cost_box.Left = .Cost.Left
12310         Select Case gblnHasForExThis
              Case True
12320           .cost_str.Visible = True
12330           .Cost.Visible = False
12340           .cost_forex.Visible = True
12350         Case False
12360           Select Case .Parent.chkIncludeCurrency
                Case True
12370             .cost_str.Visible = True
12380             .Cost.Visible = False
12390             .cost_forex.Visible = True
12400           Case False
12410             .Cost.Visible = True
12420             .cost_str.Visible = False
12430             .cost_forex.Visible = False
12440           End Select
12450         End Select
12460         .cost_lbl.Left = (.Cost.Left - lngICashLbl_Offset)
12470         .cost_lbl_line.Left = .cost_lbl.Left
12480         If blnSortHere = True Then
12490           .Sort_line.Left = .cost_lbl.Left
12500           .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
12510         End If
12520         blnSortHere = False
12530         If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
12540         .PCash.Left = (lngPCash_Left + lngCurrID_Width)
12550         .pcash_str.Left = .PCash.Left
12560         .pcash_forex.Left = (.pcash_str.Left - lngTpp)
12570         .pcash_box.Left = .PCash.Left
12580         Select Case gblnHasForExThis
              Case True
12590           .pcash_str.Visible = True
12600           .PCash.Visible = False
12610           .pcash_forex.Visible = True
12620         Case False
12630           Select Case .Parent.chkIncludeCurrency
                Case True
12640             .pcash_str.Visible = True
12650             .PCash.Visible = False
12660             .pcash_forex.Visible = True
12670           Case False
12680             .PCash.Visible = True
12690             .pcash_str.Visible = False
12700             .pcash_forex.Visible = False
12710           End Select
12720         End Select
12730         .pcash_lbl.Left = (.PCash.Left - lngICashLbl_Offset)
12740         .pcash_lbl_line.Left = .pcash_lbl.Left
12750         If blnSortHere = True Then
12760           .Sort_line.Left = .pcash_lbl.Left
12770           .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
12780         End If
12790         blnSortHere = False
12800         If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
12810         .ICash.Left = (lngICash_Left + lngCurrID_Width)
12820         .icash_str.Left = .ICash.Left
12830         .icash_forex.Left = (.icash_str.Left - lngTpp)
12840         .icash_box.Left = .ICash.Left
12850         Select Case gblnHasForExThis
              Case True
12860           .icash_str.Visible = True
12870           .ICash.Visible = False
12880           .icash_forex.Visible = True
12890         Case False
12900           Select Case .Parent.chkIncludeCurrency
                Case True
12910             .icash_str.Visible = True
12920             .ICash.Visible = False
12930             .icash_forex.Visible = True
12940           Case False
12950             .ICash.Visible = True
12960             .icash_str.Visible = False
12970             .icash_forex.Visible = False
12980           End Select
12990         End Select
13000         .icash_lbl.Left = (.ICash.Left - lngICashLbl_Offset)
13010         .icash_lbl_line.Left = .icash_lbl.Left
13020         If blnSortHere = True Then
13030           .Sort_line.Left = .icash_lbl.Left
13040           .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
13050         End If
13060         blnSortHere = False
13070         If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
13080         .shareface_0.Left = (lngShareface_Left + lngCurrID_Width)
13090         .shareface_1.Left = .shareface_0.Left
13100         .shareface_2.Left = .shareface_0.Left
13110         .shareface_3.Left = .shareface_0.Left
13120         .shareface_4.Left = .shareface_0.Left
13130         .shareface_box.Left = .shareface_0.Left
13140         .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
13150         .shareface_lbl_line.Left = .shareface_lbl.Left
13160         If blnSortHere = True Then
13170           .Sort_line.Left = .shareface_lbl.Left
13180           .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
13190         End If
13200         blnSortHere = False
13210         If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
13220         .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width + lngCurrID_Width)
13230         .totdesc_lbl.Width = .totdesc.Width
13240         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
13250         If blnSortHere = True Then
13260           .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
13270           .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
13280         End If
13290         blnSortHere = False

13300       End Select

13310       If blnResort = True Then
13320         SortNow "Form_Load"  ' ** Procedure: Below.
13330       End If

13340     End Select

13350   End With

EXITP:
13360   Exit Sub

ERRH:
13370   DoCmd.Hourglass False
13380   Select Case ERR.Number
        Case Else
13390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13400   End Select
13410   Resume EXITP

End Sub

Public Function RecCnt() As Long

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

13510   lngRetVal = 0&

13520   With Me
13530     Set rst = .RecordsetClone
13540     With rst
13550       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
13560       Else
13570         .MoveLast
13580         lngRetVal = .RecordCount
13590       End If
13600       .Close
13610     End With
13620   End With

EXITP:
13630   Set rst = Nothing
13640   RecCnt = lngRetVal
13650   Exit Function

ERRH:
13660   lngRetVal = 0&
13670   Select Case ERR.Number
        Case Else
13680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13690   End Select
13700   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

13810   Select Case intWhere
        Case acCmdRecordsGoToFirst
13820     DoCmd.RunCommand acCmdRecordsGoToFirst
13830   Case acCmdRecordsGoToLast
13840     DoCmd.RunCommand acCmdRecordsGoToLast
13850   Case acCmdRecordsGoToPrevious
13860     DoCmd.RunCommand acCmdRecordsGoToPrevious
13870   Case acCmdRecordsGoToNext
13880     DoCmd.RunCommand acCmdRecordsGoToNext
13890   Case Else
13900     If IsMissing(varID) = False Then
13910       With Me
13920         Set rst = .RecordsetClone
13930         With rst
13940           .FindFirst "[journalno] = " & CStr(varID)
13950           If .NoMatch = False Then
13960             Me.Bookmark = .Bookmark
13970           End If
13980           .Close
13990         End With
14000       End With
14010     End If
14020   End Select

EXITP:
14030   Set rst = Nothing
14040   Exit Sub

ERRH:
14050   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
14060   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
14070   Case Else
14080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14090   End Select
14100   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [JournalType_Order], [journalno]"

14210   With Me
14220     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
14230       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
14240     End If
14250     .Controls(strSortLbl).Visible = False
14260     .Controls(strSortLbl).Top = lngSortLbl_Top
14270     .Controls(strSortLine).Visible = False
14280     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
14290     .journalno_Sort_line1.Visible = False
14300     .journalno_Sort_line2.Visible = False
14310     If .ledger_HIDDEN_lbl.Left <> (.ledger_HIDDEN.Left - lngLedgerHidLbl_Offset) Then
14320       .ledger_HIDDEN_lbl.Left = (.ledger_HIDDEN.Left - lngLedgerHidLbl_Offset)
14330     End If
14340     DoEvents
14350     If strProc = "Form_Load" Then
14360       strCalled = "transdate"
14370       strSortNow = strSortOrig
14380       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
14390       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
14400       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
14410       .Controls(strSortLbl).Left = lngSortLbl_Left
14420       .Controls(strSortLine).Left = lngSortLine_Left
14430       .Controls(strSortLine).Width = lngSortLine_Width
14440       .Controls(strSortLbl).Caption = strArwDn
14450       .Controls(strSortLbl).ForeColor = CLR_DKRED
14460       DoEvents
14470     Else
14480       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: revcode_ID_lbl_DblClick
14490       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
14500       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
14510       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
14520       .Controls(strSortLbl).Left = lngSortLbl_Left
14530       .Controls(strSortLine).Left = lngSortLine_Left
14540       .Controls(strSortLine).Width = lngSortLine_Width
14550       If strCalled = "transdate" Then
14560         strSortAsc = strSortOrig
14570         If strSortNow = strSortAsc Then
14580           strSortNow = "[transdate], [JournalType_Order], [journalno]"
14590           .Controls(strSortLbl).Caption = strArwUp
14600           .Controls(strSortLbl).ForeColor = CLR_DKBLU
14610         Else
14620           strSortNow = strSortAsc
14630           .Controls(strSortLbl).Caption = strArwDn
14640           .Controls(strSortLbl).ForeColor = CLR_DKRED
14650         End If
14660         DoEvents
14670       Else
14680         .Controls(strSortLbl).Caption = strArwUp
14690         .Controls(strSortLbl).ForeColor = CLR_DKBLU
14700         Select Case strCalled
              Case "journalno"
14710           strSortAsc = "[journalno]"
14720           lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
14730           .Controls(strSortLbl).Top = (lngSortLbl_Top - (2& * lngTpp))
14740           .Controls(strSortLbl).Left = lngSortLbl_Left
14750         Case "journaltype"
14760           strSortAsc = "[journaltype], [transdate], [journalno]"
14770         Case "shareface"
14780           strSortAsc = "[shareface]" & strStdAsc
14790         Case "totdesc"
14800           strSortAsc = "[totdesc]" & strStdAsc
14810         Case "icash"
14820           strSortAsc = "[icash]" & strStdAsc
14830         Case "pcash"
14840           strSortAsc = "[pcash]" & strStdAsc
14850         Case "cost"
14860           strSortAsc = "[cost]" & strStdAsc
14870         Case "curr_id"
14880           strSortAsc = "[curr_code]" & strStdAsc
14890           lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
14900           lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
14910           lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
14920           .Controls(strSortLbl).Left = lngSortLbl_Left
14930           .Controls(strSortLine).Left = lngSortLine_Left
14940           .Controls(strSortLine).Width = lngSortLine_Width
14950         Case "ledger_HIDDEN"
14960           strSortAsc = "[ledger_HIDDEN]" & strStdAsc
14970           lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
14980           lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
14990           lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
15000           .Controls(strSortLbl).Left = lngSortLbl_Left
15010           .Controls(strSortLine).Left = lngSortLine_Left
15020           .Controls(strSortLine).Width = lngSortLine_Width
15030         End Select
15040         If strSortNow = strSortAsc Then
15050           intCnt = CharCnt(strSortAsc, ",") + 1
15060           Select Case intCnt
                Case 1
15070             strTmp01 = strSortAsc & " DESC"
15080           Case 2
15090             intPos01 = InStr(strSortAsc, ",")
15100             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
15110             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
15120           Case 3
15130             intPos01 = InStr(strSortAsc, ",")
15140             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
15150             strTmp02 = Mid(strSortAsc, intPos01)
15160             intPos01 = InStr(2, strTmp02, ",")
15170             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
15180             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
15190           Case 4
15200             intPos01 = InStr(strSortAsc, ",")
15210             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
15220             strTmp02 = Mid(strSortAsc, intPos01)
15230             intPos01 = InStr(2, strTmp02, ",")
15240             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
15250             strTmp02 = Mid(strTmp02, intPos01)
15260             intPos01 = InStr(2, strTmp02, ",")
15270             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
15280             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
15290           Case 5
15300             intPos01 = InStr(strSortAsc, ",")
15310             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
15320             strTmp02 = Mid(strSortAsc, intPos01)
15330             intPos01 = InStr(2, strTmp02, ",")
15340             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
15350             strTmp02 = Mid(strTmp02, intPos01)
15360             intPos01 = InStr(2, strTmp02, ",")
15370             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
15380             strTmp02 = Mid(strTmp02, intPos01)
15390             intPos01 = InStr(2, strTmp02, ",")
15400             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
15410             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
15420           End Select
15430           strSortNow = strTmp01
15440           .Controls(strSortLbl).Caption = strArwDn
15450           .Controls(strSortLbl).ForeColor = CLR_DKRED
15460         Else
15470           strSortNow = strSortAsc
15480         End If
15490         DoEvents
15500       End If
15510     End If
15520     If strCalled = "journalno" Then
15530       .journalno_Sort_line1.Visible = True
15540       .journalno_Sort_line2.Visible = True
15550     Else
15560       .Controls(strSortLine).Visible = True
15570     End If
15580     .Controls(strSortLbl).Visible = True
15590     DoEvents
15600     .OrderBy = strSortNow
15610     .OrderByOn = True
15620   End With

        ' ** Potential Sorts:
        ' **   [transdate] DESC, [JournalType_Order], [journalno]
        ' **   [journalno]
        ' **   [journaltype], [transdate], [journalno]
        ' **   [shareface], [transdate], [JournalType_Order], [journalno]
        ' **   [totdesc], [transdate], [JournalType_Order], [journalno]
        ' **   [icash], [transdate], [JournalType_Order], [journalno]
        ' **   [pcash], [transdate], [JournalType_Order], [journalno]
        ' **   [cost], [transdate], [JournalType_Order], [journalno]
        ' **   [curr_code], [transdate], [JournalType_Order], [journalno]
        ' **   [ledger_HIDDEN], [transdate], [JournalType_Order], [journalno]

EXITP:
15630   Exit Sub

ERRH:
15640   Select Case ERR.Number
        Case Else
15650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15660   End Select
15670   Resume EXITP

End Sub

Public Function SortNow_Get() As String

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

15710   strRetVal = strSortNow

EXITP:
15720   SortNow_Get = strRetVal
15730   Exit Function

ERRH:
15740   Select Case ERR.Number
        Case Else
15750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15760   End Select
15770   Resume EXITP

End Function

Public Function SortOrig_Get() As String

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "SortOrig_Get"

        Dim strRetVal As String

15810   strRetVal = strSortOrig

EXITP:
15820   SortOrig_Get = strRetVal
15830   Exit Function

ERRH:
15840   Select Case ERR.Number
        Case Else
15850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15860   End Select
15870   Resume EXITP

End Function
