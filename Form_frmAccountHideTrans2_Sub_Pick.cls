VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans2_Sub_Pick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans2_Sub_Pick"

'VGC 04/15/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Show All:         A {chkShowAll on frmAccountHideTrans2}
' **   Close:            C {cmdClose on frmAccountHideTrans2}
' **   ledger_HIDDEN:    D {ledger_HIDDEN on frmAccountHideTrans2_Sub_List}
' **   Hidden Groups:    G {cmdHidden on frmAccountHideTrans2}
' **   Cancel:           E {cmdCancel on frmAccountHideTrans2}
' **   Include Archive:  H {chkIncludeArchive on frmAccountHideTrans2}
' **   UnHide:           I {cmdUnHide on frmAccountHideTrans2}
' **   Journal No:       J {chkShowJournalNo on frmAccountHideTrans2}
' **   ledger_HIDDEN:    N {ledger_HIDDEN}
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
Private blnCancelUpdate As Boolean
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
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

190         blnCancelUpdate = False

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

        Dim strFormat As String

610     With Me

620       SortNow THIS_PROC  ' ** Procedure: Below

630       .transdate.ForeColor = CLR_DISABLED_FG
640       .transdate.BackColor = CLR_DISABLED_BG
650       .journaltype.ForeColor = CLR_DISABLED_FG
660       .journaltype.BackColor = CLR_DISABLED_BG
670       .shareface_0.ForeColor = CLR_DISABLED_FG
680       .shareface_1.ForeColor = CLR_DISABLED_FG
690       .shareface_2.ForeColor = CLR_DISABLED_FG
700       .shareface_3.ForeColor = CLR_DISABLED_FG
710       .shareface_4.ForeColor = CLR_DISABLED_FG
720       .shareface_box.BackColor = CLR_DISABLED_BG
730       .totdesc.ForeColor = CLR_DISABLED_FG
740       .totdesc.BackColor = CLR_DISABLED_BG
750       .ICash.ForeColor = CLR_DISABLED_FG
760       .ICash.BackColor = CLR_DISABLED_BG
770       .PCash.ForeColor = CLR_DISABLED_FG
780       .PCash.BackColor = CLR_DISABLED_BG
790       .Cost.ForeColor = CLR_DISABLED_FG
800       .Cost.BackColor = CLR_DISABLED_BG

810       strFormat = "$#,##0.00;($#,##0.00);$0.00;$0.00"
820       .Parent.icash_subtot.Format = strFormat
830       .Parent.pcash_subtot.Format = strFormat
840       .Parent.cost_subtot.Format = strFormat

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
        ' **   ledger_HIDDEN:    D {ledger_HIDDEN on frmAccountHideTrans2_Sub_List}
        ' **   ledger_HIDDEN:    N {ledger_HIDDEN}

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
1180          DoCmd.SelectObject acForm, .Parent.Name
1190          .Parent.frmAccountHideTrans2_Sub_List.SetFocus
1200          .Parent.frmAccountHideTrans2_Sub_List.Form.ledger_HIDDEN.SetFocus  ' ** Don't invoke it!
1210        End With
1220      Case vbKeyN
1230        With Me
1240          intRetVal = 0
1250          .ledger_HIDDEN.SetFocus
1260        End With
1270      Case vbKeyX
1280        With Me
1290          intRetVal = 0
1300          DoCmd.SelectObject acForm, .Parent.Name, False
1310          .Parent.cmdClose.SetFocus
1320          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountHideTrans2
1330        End With
1340      End Select
1350    End If

        ' ** Ctrl keys.
1360    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1370      Select Case intRetVal
          Case vbKeyS
1380        intRetVal = 0
1390        cmdSave_Click  ' ** Procedure: Below.
1400      Case vbKeyTab, vbKeyReturn
1410        With Me
1420          intRetVal = 0
1430          DoCmd.SelectObject acForm, .Parent.Name, False
1440          If .Parent.cmdOK.Visible = True And .Parent.cmdOK.Enabled = True Then
1450            .Parent.cmdOK.SetFocus
1460          ElseIf .Parent.cmdUnHide.Visible = True And .Parent.cmdUnHide.Enabled = True Then
1470            .Parent.cmdUnHide.SetFocus
1480          ElseIf .Parent.cmdCancel.Enabled = True Then
1490            .Parent.cmdCancel.SetFocus
1500          ElseIf .Parent.chkIncludeArchive.Enabled = True Then
1510            .Parent.chkIncludeArchive.SetFocus
1520          ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
1530            .Parent.chkIncludeCurrency.SetFocus
1540          ElseIf .Parent.chkShowAll.Enabled = True Then
1550            .Parent.chkShowAll.SetFocus
1560          ElseIf .Parent.chkShowJournalNo.Enabled = True Then
1570            .Parent.chkShowJournalNo.SetFocus
1580          ElseIf .Parent.chkShowHiddenOnly.Enabled = True Then
1590            .Parent.chkShowHiddenOnly.SetFocus
1600          ElseIf .Parent.cmdPrintReport.Enabled = True Then
1610            .Parent.cmdPrintReport.SetFocus
1620          ElseIf .Parent.cmdHidden.Enabled = True Then
1630            .Parent.cmdHidden.SetFocus
1640          Else
1650            .Parent.cmdClose.SetFocus
1660          End If
1670        End With
1680      Case vbKeyUp, vbKeyPageUp, vbKeyHome
1690        intRetVal = 0
1700        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1710      Case vbKeyDown, vbKeyPageDown, vbKeyEnd
1720        intRetVal = 0
1730        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1740      End Select
1750    End If

        ' ** Ctrl-Shift keys.
1760    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1770      Select Case intRetVal
          Case vbKeyF
1780        With Me
1790          intRetVal = 0
1800          DoCmd.SelectObject acForm, .Parent.Name, False
1810          .Parent.FocusHolder.SetFocus
1820        End With
1830      Case vbKeyTab, vbKeyReturn
1840        With Me
1850          intRetVal = 0
1860          DoCmd.SelectObject acForm, .Parent.Name
1870          .Parent.frmAccountHideTrans2_Sub_List.SetFocus
1880        End With
1890      End Select
1900    End If

EXITP:
1910    KeyCode = intRetVal
1920    Exit Sub

ERRH:
1930    intRetVal = 0
1940    Select Case ERR.Number
        Case Else
1950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1960    End Select
1970    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2010    DoCmd.RunCommand acCmdSaveRecord
2020    DoEvents

EXITP:
2030    Exit Sub

ERRH:
2040    Select Case ERR.Number
        Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
2110    intShiftDown = (Shift And acShiftMask) > 0
2120    intAltDown = (Shift And acAltMask) > 0
2130    intCtrlDown = (Shift And acCtrlMask) > 0

2140    intRetVal = KeyCode

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2170        With Me
2180          intRetVal = 0
2190          lngRecsCur = RecCnt  ' ** Function: Below.
2200          If lngRecsCur > 0& Then
2210            .ledger_HIDDEN.SetFocus
2220          Else
2230            DoCmd.SelectObject acForm, .Parent.Name, False
2240            .Parent.cmdClose.SetFocus
2250          End If
2260        End With
2270      End Select
2280    End If

        ' ** Shift keys.
2290    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2300      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2310        With Me
2320          intRetVal = 0
2330          lngRecsCur = RecCnt  ' ** Function: Below.
2340          If lngRecsCur > 0& Then
2350            .ledger_HIDDEN.SetFocus
2360          Else
2370            DoCmd.SelectObject acForm, .Parent.Name, False
2380            .Parent.cmdClose.SetFocus
2390          End If
2400        End With
2410      End Select
2420    End If

EXITP:
2430    KeyCode = intRetVal
2440    Exit Sub

ERRH:
2450    intRetVal = 0
2460    Select Case ERR.Number
        Case Else
2470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2480    End Select
2490    Resume EXITP

End Sub

Private Sub journalno_lbl_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_lbl_DblClick"

2510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

2610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

2710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

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

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

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

Private Sub icash_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

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

Private Sub pcash_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

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

Private Sub cost_lbl_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

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

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

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

Private Sub ledger_HIDDEN_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_lbl_DblClick"

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

Private Sub ledger_HIDDEN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_MouseDown"

3510    With Me
3520      If .Parent.cmdUnHide.Visible = True Then
3530        Beep
3540        MsgBox "Transactions may only be unhidden as a group." & vbCrLf & vbCrLf & _
              "Use the UnHide button below.", vbInformation + vbOKOnly, "Unhide Transaction Group Only"
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

Private Sub ledger_HIDDEN_BeforeUpdate(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_BeforeUpdate"

3710    With Me
3720      If .Parent.cmdUnHide.Visible = True Then
            ' ** Unhide only as a group!
3730        .Undo
3740        DoEvents
3750        blnCancelUpdate = True
3760        Beep
3770      End If
3780    End With

EXITP:
3790    Exit Sub

ERRH:
3800    Select Case ERR.Number
        Case Else
3810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3820    End Select
3830    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_AfterUpdate()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngJournalNo As Long
        Dim strFormat As String

3910    With Me
3920      If blnCancelUpdate = False Then
3930        lngJournalNo = .journalno
3940        Select Case .ledger_HIDDEN
            Case True
3950          Beep
3960        Case False
3970          DoCmd.Hourglass True
3980          DoEvents
3990          cmdSave_Click  ' ** Procedure: Above.
4000          Set dbs = CurrentDb
4010          With dbs
                ' ** Delete tmpAccountHideTrans_03, by specified [jno].
4020            Set qdf = .QueryDefs("qryAccountHideTrans2_15")
4030            With qdf.Parameters
4040              ![jno] = lngJournalNo
4050            End With
4060            qdf.Execute dbFailOnError
4070            Set qdf = Nothing
                ' ** Update tmpAccountHideTrans_02, by specified [jno], [lhid], [chng].
4080            Set qdf = .QueryDefs("qryAccountHideTrans2_16")
4090            With qdf.Parameters
4100              ![jno] = lngJournalNo
4110              ![lhid] = False
4120              ![chng] = False
4130            End With
4140            qdf.Execute
4150            Set qdf = Nothing
4160            .Close
4170          End With
4180          .Requery
4190          DoEvents
4200          lngRecsCur = RecCnt  ' ** Function: Below.
4210          DoEvents
4220          If lngRecsCur = 0& Then
4230            .Parent.cmdOK.Enabled = False
4240            .Parent.cmdCancel.Enabled = False
4250            .Parent.TotalsEnable False  ' ** Form Procedure: frmAccountHideTrans2.
4260            .Parent.CheckBoxLock False  ' ** Form Procedure: frmAccountHideTrans2.
4270            .Parent.PickSubExpand False  ' ** Form Procedure: frmAccountHideTrans2.
4280            strFormat = "$#,##0.00;($#,##0.00);$0.00;$0.00"
4290            .Parent.icash_subtot.Format = strFormat
4300            .Parent.pcash_subtot.Format = strFormat
4310            .Parent.cost_subtot.Format = strFormat
4320          End If
4330          .Parent.running_pick_cnt = .Parent.running_pick_cnt - 1&
4340          .Parent.UpdateHiddenTrans  ' ** Form Procedure: frmAccountHideTrans2.
4350          .Parent.frmAccountHideTrans2_Sub_List.Form.Refresh
4360          DoEvents
4370          .Parent.UpdateTotals_Set True, 500&  ' ** Form Procedure: frmAccountHideTrans2.
4380          If .Parent.running_pick_cnt = 0& Then
4390            DoCmd.SelectObject acForm, .Parent.Name, False
4400            .Parent.frmAccountHideTrans2_Sub_List.SetFocus
4410          End If
4420        End Select
4430      Else
4440        blnCancelUpdate = False
4450      End If  ' ** blnCancelUpdate.
4460    End With

4470    DoCmd.Hourglass False

EXITP:
4480    Set qdf = Nothing
4490    Set dbs = Nothing
4500    Exit Sub

ERRH:
4510    DoCmd.Hourglass False
4520    Select Case ERR.Number
        Case Else
4530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4540    End Select
4550    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
4610    intShiftDown = (Shift And acShiftMask) > 0
4620    intAltDown = (Shift And acAltMask) > 0
4630    intCtrlDown = (Shift And acCtrlMask) > 0

4640    intRetVal = KeyCode

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4670        With Me
4680          intRetVal = 0
4690          lngRecsCur = RecCnt  ' ** Function: Below.
4700          If .CurrentRecord < lngRecsCur Then
4710            MoveRec acCmdRecordsGoToNext
4720            .ledger_HIDDEN.SetFocus
4730          Else
4740            If .Parent.cmdOK.Visible And .Parent.cmdOK.Enabled = True Then
4750              .Parent.cmdOK.SetFocus
4760            ElseIf .Parent.cmdUnHide.Visible = True And .Parent.cmdUnHide.Enabled = True Then
4770              .Parent.cmdUnHide.SetFocus
4780            ElseIf .Parent.cmdCancel.Enabled = True Then
4790              .Parent.cmdCancel.SetFocus
4800            ElseIf .Parent.chkIncludeArchive.Enabled = True Then
4810              .Parent.chkIncludeArchive.SetFocus
4820            ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
4830              .Parent.chkIncludeCurrency.SetFocus
4840            ElseIf .Parent.chkShowAll.Enabled = True Then
4850              .Parent.chkShowAll.SetFocus
4860            ElseIf .Parent.chkShowJournalNo.Enabled = True Then
4870              .Parent.chkShowJournalNo.SetFocus
4880            ElseIf .Parent.chkShowHiddenOnly.Enabled = True Then
4890              .Parent.chkShowHiddenOnly.SetFocus
4900            ElseIf .Parent.cmdPrintReport.Enabled = True Then
4910              .Parent.cmdPrintReport.SetFocus
4920            ElseIf .Parent.cmdHidden.Enabled = True Then
4930              .Parent.cmdHidden.SetFocus
4940            Else
4950              .Parent.cmdClose.SetFocus
4960            End If
4970          End If
4980        End With
4990      End Select
5000    End If

        ' ** Shift keys.
5010    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5020      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5030        With Me
5040          intRetVal = 0
5050          If .CurrentRecord > 1 Then
5060            MoveRec acCmdRecordsGoToPrevious
5070            .ledger_HIDDEN.SetFocus
5080          Else
5090            DoCmd.SelectObject acForm, .Parent.Name, False
5100            .Parent.frmAccountHideTrans2_Sub_List.SetFocus
5110            .Parent.frmAccountHideTrans2_Sub_List.Form.ledger_HIDDEN.SetFocus
5120          End If
5130        End With
5140      End Select
5150    End If

EXITP:
5160    KeyCode = intRetVal
5170    Exit Sub

ERRH:
5180    intRetVal = 0
5190    Select Case ERR.Number
        Case Else
5200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5210    End Select
5220    Resume EXITP

End Sub

Public Function IsPicked(lngJournalNo As Long) As Boolean

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "IsPicked"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

5310    blnRetVal = False

5320    With Me
5330      Set rst = .RecordsetClone
5340      With rst
5350        If .BOF = True And .EOF = True Then
              ' ** Nothing picked.
5360        Else
5370          .MoveLast
5380          lngRecs = .RecordCount
5390          .MoveFirst
5400          .FindFirst "[journalno] = " & CStr(lngJournalNo)
5410          If .NoMatch = False Then
5420            blnRetVal = True
5430          End If
5440        End If
5450        .Close
5460      End With
5470      If blnRetVal = False Then
            ' ** Something's causing it to give false results!
5480        varTmp00 = DCount("*", "tmpAccountHideTrans_03", "[journalno] = " & CStr(lngJournalNo))
5490        If IsNull(varTmp00) = False Then
5500          If varTmp00 > 0 Then
5510            blnRetVal = True
5520          End If
5530        End If
5540      End If
5550    End With

EXITP:
5560    IsPicked = blnRetVal
5570    Exit Function

ERRH:
5580    blnRetVal = False
5590    Select Case ERR.Number
        Case Else
5600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5610    End Select
5620    Resume EXITP

End Function

Public Sub SubEnable(blnEnable As Boolean)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "SubEnable"

        Dim ctl As Access.Control
        Dim blnJournalNo As Boolean, blnCurrency As Boolean
        Dim blnJNo As Boolean

5710    With Me
5720      blnJournalNo = .Parent.chkShowJournalNo
5730      blnCurrency = .Parent.chkIncludeCurrency
5740      For Each ctl In .FormHeader.Controls
5750        With ctl
5760          If (InStr(.Name, "journalno") > 0 And blnJournalNo = False) Or _
                  (InStr(.Name, "curr_id") > 0 And blnCurrency = False) Then
                ' ** Skip these
5770          Else
5780            If Left(.Name, 9) = "journalno" Then
5790              blnJNo = True
5800            Else
5810              blnJNo = False
5820            End If
5830            If InStr(.Name, "_lbl") > 0 Then
5840              If Right(.Name, 7) = "_dim_hi" Then
5850                .Visible = (Not blnEnable)
5860              Else
5870                Select Case blnEnable
                    Case True
5880                  Select Case .ControlType
                      Case acLabel
5890                    Select Case blnJNo
                        Case True
5900                      .ForeColor = CLR_GRY4
5910                    Case False
5920                      .ForeColor = CLR_DKGRY2
5930                    End Select
5940                  Case acLine
5950                    Select Case blnJNo
                        Case True
5960                      .BorderColor = CLR_GRY5
5970                    Case False
5980                      .BorderColor = CLR_DKGRY
5990                    End Select
6000                  End Select
6010                Case False
6020                  Select Case .ControlType
                      Case acLabel
6030                    .ForeColor = WIN_CLR_DISF
6040                  Case acLine
6050                    .BorderColor = WIN_CLR_DISR
6060                  End Select
6070                End Select
6080              End If
6090            End If
6100          End If
6110        End With  ' ** ctl.
6120      Next  ' ** ctl.
6130    End With

EXITP:
6140    Set ctl = Nothing
6150    Exit Sub

ERRH:
6160    Select Case ERR.Number
        Case Else
6170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6180    End Select
6190    Resume EXITP

End Sub

Public Sub ShowJournalNo_Sub(blnShow As Boolean)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo_Sub"

        Dim blnCurrency As Boolean, blnSortHere As Boolean, blnResort As Boolean

6210    With Me

6220      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
6230        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6240      End If

6250      blnCurrency = .Parent.chkIncludeCurrency
6260      blnSortHere = False

6270      Select Case blnShow
          Case True

6280        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
6290        .totdesc.Left = lngTotDesc_Left
6300        .totdesc_lbl.Left = lngTotDesc_Left
6310        .totdesc_lbl_dim_hi.Left = (.totdesc_lbl.Left + lngTpp)
6320        .totdesc_lbl_line.Left = lngTotDesc_Left
6330        .totdesc_lbl_line_dim_hi.Left = (.totdesc_lbl_line.Left + lngTpp)
6340        Select Case blnCurrency
            Case True
6350          .totdesc.Width = lngTotDesc_Width
6360          .totdesc_lbl.Width = .totdesc.Width
6370          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
6380          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6390          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
6400        Case False
6410          .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)
6420          .totdesc_lbl.Width = .totdesc.Width
6430          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
6440          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6450          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
6460        End Select
6470        If blnSortHere = True Then
6480          .Sort_line.Left = .totdesc_lbl.Left
6490          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
6500          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
6510        End If
6520        blnSortHere = False
            'If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
            '.shareface_0.Left = lngShareface_Left
            '.shareface_1.Left = lngShareface_Left
            '.shareface_2.Left = lngShareface_Left
            '.shareface_3.Left = lngShareface_Left
            '.shareface_4.Left = lngShareface_Left
            '.shareface_box.Left = lngShareface_Left
            '.shareface_lbl.Left = lngShareface_Left
            '.shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
            '.shareface_lbl_line.Left = lngShareface_Left
            '.shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
            'If blnSortHere = True Then
            '  .Sort_line.Left = .shareface_lbl.Left
            '  .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
            'End If
            'blnSortHere = False
6530        If .Sort_line.Left = .journaltype_lbl.Left Then blnSortHere = True
6540        .journaltype.Left = lngJournalType_Left
6550        .journaltype_lbl.Left = lngJournalType_Left
6560        .journaltype_lbl_dim_hi.Left = (.journaltype_lbl.Left + lngTpp)
6570        .journaltype_lbl_line.Left = lngJournalType_Left
6580        .journaltype_lbl_line_dim_hi.Left = (.journaltype_lbl_line.Left + lngTpp)
6590        If blnSortHere = True Then
6600          .Sort_line.Left = .journaltype_lbl.Left
6610          .Sort_lbl.Left = ((.journaltype_lbl.Left + .journaltype_lbl.Width) - lngSortLbl_Width)
6620        End If
6630        blnSortHere = False
6640        If .Sort_line.Left = .transdate_lbl.Left Then blnSortHere = True
6650        .transdate.Left = lngTransdate_Left
6660        .transdate_lbl.Left = lngTransdate_Left
6670        .transdate_lbl_dim_hi.Left = (.transdate_lbl.Left + lngTpp)
6680        .transdate_lbl_line.Left = lngTransdate_Left
6690        .transdate_lbl_line_dim_hi.Left = (.transdate_lbl_line.Left + lngTpp)
6700        If blnSortHere = True Then
6710          .Sort_line.Left = .transdate_lbl.Left
6720          .Sort_lbl.Left = ((.transdate_lbl.Left + .transdate_lbl.Width) - lngSortLbl_Width)
6730        End If
6740        blnSortHere = False

6750        lngRecsCur = RecCnt  ' ** Function: Below.
6760        If lngRecsCur = 0 Then
6770          .journalno.Visible = True
6780          .journalno_lbl.ForeColor = WIN_CLR_DISF
6790          .journalno_lbl.Visible = True
6800          .journalno_lbl_dim_hi.Visible = True
6810          .journalno_lbl2.ForeColor = WIN_CLR_DISF
6820          .journalno_lbl2.Visible = True
6830          .journalno_lbl2_dim_hi.Visible = True
6840          .journalno_lbl_line.BorderColor = WIN_CLR_DISR
6850          .journalno_lbl_line.Visible = True
6860          .journalno_lbl_line_dim_hi.Visible = True
6870        Else
6880          .journalno.Visible = True
6890          .journalno_lbl.ForeColor = CLR_GRY4
6900          .journalno_lbl.Visible = True
6910          .journalno_lbl2.ForeColor = CLR_GRY4
6920          .journalno_lbl2.Visible = True
6930          .journalno_lbl_line.BorderColor = CLR_GRY5
6940          .journalno_lbl_line.Visible = True
6950        End If

6960      Case False

6970        If .journalno_Sort_line1.Visible = True Then blnSortHere = True
6980        .journalno.Visible = False
6990        .journalno_lbl.Visible = False
7000        .journalno_lbl_dim_hi.Visible = False
7010        .journalno_lbl2.Visible = False
7020        .journalno_lbl2_dim_hi.Visible = False
7030        .journalno_lbl_line.Visible = False
7040        .journalno_lbl_line_dim_hi.Visible = False
7050        .journalno_Sort_line1.Visible = False
7060        .journalno_Sort_line2.Visible = False
7070        If blnSortHere = True Then
7080          blnResort = True
7090        End If
7100        blnSortHere = False
7110        If .Sort_line.Left = .transdate_lbl.Left Then blnSortHere = True
7120        .transdate.Left = (lngTransdate_Left - lngJournalNo_Width)
7130        .transdate_lbl.Left = .transdate.Left
7140        .transdate_lbl_dim_hi.Left = (.transdate_lbl.Left + lngTpp)
7150        .transdate_lbl_line.Left = .transdate_lbl.Left
7160        .transdate_lbl_line_dim_hi.Left = (.transdate_lbl_line.Left + lngTpp)
7170        If blnSortHere = True Then
7180          .Sort_line.Left = .transdate_lbl.Left
7190          .Sort_lbl.Left = ((.transdate_lbl.Left + .transdate_lbl.Width) - lngSortLbl_Width)
7200        End If
7210        blnSortHere = False
7220        If .Sort_line.Left = .journaltype_lbl.Left Then blnSortHere = True
7230        .journaltype.Left = (lngJournalType_Left - lngJournalNo_Width)
7240        .journaltype_lbl.Left = .journaltype.Left
7250        .journaltype_lbl_dim_hi.Left = (.journaltype_lbl.Left + lngTpp)
7260        .journaltype_lbl_line.Left = .journaltype_lbl.Left
7270        .journaltype_lbl_line_dim_hi.Left = (.journaltype_lbl_line.Left + lngTpp)
7280        If blnSortHere = True Then
7290          .Sort_line.Left = .journaltype_lbl.Left
7300          .Sort_lbl.Left = ((.journaltype_lbl.Left + .journaltype_lbl.Width) - lngSortLbl_Width)
7310        End If
7320        blnSortHere = False
            'If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
            '.shareface_0.Left = (lngShareface_Left - lngJournalNo_Width)
            '.shareface_1.Left = .shareface_0.Left
            '.shareface_2.Left = .shareface_0.Left
            '.shareface_3.Left = .shareface_0.Left
            '.shareface_4.Left = .shareface_0.Left
            '.shareface_box.Left = .shareface_0.Left
            '.shareface_lbl.Left = .shareface_0.Left
            '.shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
            '.shareface_lbl_line.Left = .shareface_0.Left
            '.shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
            'If blnSortHere = True Then
            '  .Sort_line.Left = .shareface_lbl.Left
            '  .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
            'End If
            'blnSortHere = False
7330        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
7340        .totdesc.Left = (lngTotDesc_Left - lngJournalNo_Width)
7350        .totdesc_lbl.Left = .totdesc.Left
7360        .totdesc_lbl_dim_hi.Left = (.totdesc_lbl.Left + lngTpp)
7370        .totdesc_lbl_line.Left = .totdesc_lbl.Left
7380        .totdesc_lbl_line_dim_hi.Left = (.totdesc_lbl_line.Left + lngTpp)
7390        Select Case blnCurrency
            Case True
7400          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)
7410          .totdesc_lbl.Width = .totdesc.Width
7420          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
7430          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7440          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
7450        Case False
7460          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width + lngCurrID_Width)
7470          .totdesc_lbl.Width = .totdesc.Width
7480          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
7490          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7500          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
7510        End Select
7520        If blnSortHere = True Then
7530          .Sort_line.Left = .totdesc_lbl.Left
7540          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
7550          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
7560        End If
7570        blnSortHere = False

7580        If blnResort = True Then
7590          SortNow "Form_Load"  ' ** Procedure: Below.
7600        End If

7610      End Select

7620    End With

EXITP:
7630    Exit Sub

ERRH:
7640    Select Case ERR.Number
        Case Else
7650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7660    End Select
7670    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnInclude As Boolean)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnJournalNo As Boolean, blnSortHere As Boolean, blnResort As Boolean

7710    With Me

7720      DoCmd.Hourglass True
7730      DoEvents

7740      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
7750        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7760      End If

7770      blnJournalNo = .Parent.chkShowJournalNo
7780      blnSortHere = False

7790      Select Case blnInclude
          Case True

7800        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
7810        Select Case blnJournalNo
            Case True
7820          .totdesc.Width = lngTotDesc_Width
7830          .totdesc_lbl.Width = .totdesc.Width
7840          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
7850          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7860          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
7870        Case False
7880          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)
7890          .totdesc_lbl.Width = .totdesc.Width
7900          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
7910          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7920          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
7930        End Select
7940        If blnSortHere = True Then
7950          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
7960          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
7970        End If
7980        blnSortHere = False
7990        If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
8000        .shareface_0.Left = lngShareface_Left
8010        .shareface_1.Left = .shareface_0.Left
8020        .shareface_2.Left = .shareface_0.Left
8030        .shareface_3.Left = .shareface_0.Left
8040        .shareface_4.Left = .shareface_0.Left
8050        .shareface_box.Left = .shareface_0.Left
8060        .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
8070        .shareface_lbl_line.Left = .shareface_lbl.Left
8080        .shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
8090        .shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
8100        If blnSortHere = True Then
8110          .Sort_line.Left = .shareface_lbl_line.Left
8120          .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
8130        End If
8140        blnSortHere = False
8150        If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
8160        .ICash.Left = lngICash_Left
8170        .icash_str.Left = .ICash.Left
8180        .icash_forex.Left = (.icash_str.Left - lngTpp)
8190        .icash_box.Left = .ICash.Left
8200        .icash_str.Visible = True
8210        .ICash.Visible = False
8220        .icash_forex.Visible = True
8230        .icash_lbl.Left = (.ICash.Left - lngICashLbl_Offset)
8240        .icash_lbl_dim_hi.Left = (.icash_lbl.Left + lngTpp)
8250        .icash_lbl_line.Left = .icash_lbl.Left
8260        .icash_lbl_line_dim_hi.Left = (.icash_lbl_line.Left + lngTpp)
8270        If blnSortHere = True Then
8280          .Sort_line.Left = .icash_lbl.Left
8290          .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
8300        End If
8310        blnSortHere = False
8320        If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
8330        .PCash.Left = lngPCash_Left
8340        .pcash_str.Left = .PCash.Left
8350        .pcash_forex.Left = (.pcash_str.Left - lngTpp)
8360        .pcash_box.Left = .PCash.Left
8370        .pcash_str.Visible = True
8380        .PCash.Visible = False
8390        .pcash_forex.Visible = True
8400        .pcash_lbl.Left = (.PCash.Left - lngICashLbl_Offset)
8410        .pcash_lbl_dim_hi.Left = (.pcash_lbl.Left + lngTpp)
8420        .pcash_lbl_line.Left = .pcash_lbl.Left
8430        .pcash_lbl_line_dim_hi.Left = (.pcash_lbl_line.Left + lngTpp)
8440        If blnSortHere = True Then
8450          .Sort_line.Left = .pcash_lbl.Left
8460          .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
8470        End If
8480        blnSortHere = False
8490        If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
8500        .Cost.Left = lngCost_Left
8510        .cost_str.Left = .Cost.Left
8520        .cost_forex.Left = (.cost_str.Left - lngTpp)
8530        .cost_box.Left = .Cost.Left
8540        .cost_str.Visible = True
8550        .Cost.Visible = False
8560        .cost_forex.Visible = True
8570        .cost_lbl.Left = (.Cost.Left - lngICashLbl_Offset)
8580        .cost_lbl_dim_hi.Left = (.cost_lbl.Left + lngTpp)
8590        .cost_lbl_line.Left = .cost_lbl.Left
8600        .cost_lbl_line_dim_hi.Left = (.cost_lbl_line.Left + lngTpp)
8610        If blnSortHere = True Then
8620          .Sort_line.Left = .cost_lbl.Left
8630          .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
8640        End If
8650        blnSortHere = False
8660        .curr_id.Left = lngCurrID_Left
8670        .curr_id_forex.Left = (.curr_id.Left - lngTpp)
8680        .curr_id_box.Left = .curr_id.Left
8690        .curr_id_lbl.Left = (lngCurrID_Left - lngCurrIDLbl_Offset)
8700        .curr_id_lbl_dim_hi.Left = (.curr_id_lbl.Left + lngTpp)
8710        .curr_id_lbl_line.Left = lngCurrID_Left
8720        .curr_id_lbl_line_dim_hi.Left = (.curr_id_lbl_line.Left + lngTpp)
8730        lngRecsCur = RecCnt  ' ** Function: Below.
8740        If lngRecsCur = 0 Then
8750          .curr_id_lbl.ForeColor = WIN_CLR_DISF
8760          .curr_id_lbl.Visible = True
8770          .curr_id_lbl_dim_hi.Visible = True
8780          .curr_id_lbl_line.BorderColor = WIN_CLR_DISR
8790          .curr_id_lbl_line.Visible = True
8800          .curr_id_lbl_line_dim_hi.Visible = True
8810        Else
8820          .curr_id_lbl.ForeColor = CLR_DKGRY2
8830          .curr_id_lbl.Visible = True
8840          .curr_id_lbl_line.BorderColor = CLR_DKGRY
8850          .curr_id_lbl_line.Visible = True
8860        End If
8870        .curr_id.Visible = True
8880        .curr_id_forex.Visible = True
8890        .curr_id_box.Visible = True

8900      Case False

8910        blnResort = False
8920        Select Case blnJournalNo
            Case True
8930          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
8940          .curr_id.Visible = False
8950          .curr_id_forex.Visible = False
8960          .curr_id_box.Visible = False
8970          .curr_id_lbl.Visible = False
8980          .curr_id_lbl_dim_hi.Visible = False
8990          .curr_id_lbl_line.Visible = False
9000          .curr_id_lbl_line_dim_hi.Visible = False
9010          .curr_id.Left = 0&
9020          .curr_id_forex.Left = 0&
9030          .curr_id_box.Left = 0&
9040          .curr_id_lbl.Left = 0&
9050          .curr_id_lbl_dim_hi.Left = 0&
9060          .curr_id_lbl_line.Left = 0&
9070          .curr_id_lbl_line_dim_hi.Left = 0&
9080          If blnSortHere = True Then
9090            blnResort = True
9100            .Sort_line.Visible = False
9110            .Sort_lbl.Visible = False
9120          End If
9130          blnSortHere = False
9140          If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
9150          .Cost.Left = (lngCost_Left + lngCurrID_Width)
9160          .cost_str.Left = .Cost.Left
9170          .cost_forex.Left = (.cost_str.Left - lngTpp)
9180          .cost_box.Left = .Cost.Left
9190          Select Case gblnHasForExThis
              Case True
9200            .Cost.Visible = False
9210            .cost_str.Visible = True
9220            .cost_forex.Visible = True
9230          Case False
9240            Select Case .Parent.chkIncludeCurrency
                Case True
9250              .Cost.Visible = False
9260              .cost_str.Visible = True
9270              .cost_forex.Visible = True
9280            Case False
9290              .Cost.Visible = True
9300              .cost_str.Visible = False
9310              .cost_forex.Visible = False
9320            End Select
9330          End Select
9340          .cost_lbl.Left = (.Cost.Left - lngICashLbl_Offset)
9350          .cost_lbl_dim_hi.Left = (.cost_lbl.Left + lngTpp)
9360          .cost_lbl_line.Left = .cost_lbl.Left
9370          .cost_lbl_line_dim_hi.Left = (.cost_lbl_line.Left + lngTpp)
9380          If blnSortHere = True Then
9390            .Sort_line.Left = .cost_lbl.Left
9400            .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
9410          End If
9420          blnSortHere = False
9430          If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
9440          .PCash.Left = (lngPCash_Left + lngCurrID_Width)
9450          .pcash_str.Left = .PCash.Left
9460          .pcash_forex.Left = (.pcash_str.Left - lngTpp)
9470          .pcash_box.Left = .PCash.Left
9480          Select Case gblnHasForExThis
              Case True
9490            .pcash_str.Visible = True
9500            .PCash.Visible = False
9510            .pcash_forex.Visible = True
9520          Case False
9530            Select Case .Parent.chkIncludeCurrency
                Case True
9540              .pcash_str.Visible = True
9550              .PCash.Visible = False
9560              .pcash_forex.Visible = True
9570            Case False
9580              .PCash.Visible = True
9590              .pcash_str.Visible = False
9600              .pcash_forex.Visible = False
9610            End Select
9620          End Select
9630          .pcash_lbl.Left = (.PCash.Left - lngICashLbl_Offset)
9640          .pcash_lbl_dim_hi.Left = (.pcash_lbl.Left + lngTpp)
9650          .pcash_lbl_line.Left = .pcash_lbl.Left
9660          .pcash_lbl_line_dim_hi.Left = (.pcash_lbl_line.Left + lngTpp)
9670          If blnSortHere = True Then
9680            .Sort_line.Left = .pcash_lbl.Left
9690            .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
9700          End If
9710          blnSortHere = False
9720          If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
9730          .ICash.Left = (lngICash_Left + lngCurrID_Width)
9740          .icash_str.Left = .ICash.Left
9750          .icash_forex.Left = (.icash_str.Left - lngTpp)
9760          .icash_box.Left = .ICash.Left
9770          Select Case gblnHasForExThis
              Case True
9780            .icash_str.Visible = True
9790            .ICash.Visible = False
9800            .icash_forex.Visible = True
9810          Case False
9820            Select Case .Parent.chkIncludeCurrency
                Case True
9830              .icash_str.Visible = True
9840              .ICash.Visible = False
9850              .icash_forex.Visible = True
9860            Case False
9870              .ICash.Visible = True
9880              .icash_str.Visible = False
9890              .icash_forex.Visible = False
9900            End Select
9910          End Select
9920          .icash_lbl.Left = (.ICash.Left - lngICashLbl_Offset)
9930          .icash_lbl_dim_hi.Left = (.icash_lbl.Left + lngTpp)
9940          .icash_lbl_line.Left = .icash_lbl.Left
9950          .icash_lbl_line_dim_hi.Left = (.icash_lbl_line.Left + lngTpp)
9960          If blnSortHere = True Then
9970            .Sort_line.Left = .icash_lbl.Left
9980            .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
9990          End If
10000         blnSortHere = False
10010         If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
10020         .shareface_0.Left = (lngShareface_Left + lngCurrID_Width)
10030         .shareface_1.Left = .shareface_0.Left
10040         .shareface_2.Left = .shareface_0.Left
10050         .shareface_3.Left = .shareface_0.Left
10060         .shareface_4.Left = .shareface_0.Left
10070         .shareface_box.Left = .shareface_0.Left
10080         .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
10090         .shareface_lbl_line.Left = .shareface_lbl.Left
10100         .shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
10110         .shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
10120         If blnSortHere = True Then
10130           .Sort_line.Left = .shareface_lbl_line.Left
10140           .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
10150         End If
10160         blnSortHere = False
10170         If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
10180         .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)
10190         .totdesc_lbl.Width = .totdesc.Width
10200         .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
10210         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
10220         .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
10230         If blnSortHere = True Then
10240           .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
10250           .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
10260         End If
10270         blnSortHere = False
10280       Case False
10290         If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
10300         .curr_id.Visible = False
10310         .curr_id_forex.Visible = False
10320         .curr_id_box.Visible = False
10330         .curr_id_lbl.Visible = False
10340         .curr_id_lbl_dim_hi.Visible = False
10350         .curr_id_lbl_line.Visible = False
10360         .curr_id_lbl_line_dim_hi.Visible = False
10370         .curr_id.Left = 0&
10380         .curr_id_forex.Left = 0&
10390         .curr_id_box.Left = 0&
10400         .curr_id_lbl.Left = 0&
10410         .curr_id_lbl_dim_hi.Left = 0&
10420         .curr_id_lbl_line.Left = 0&
10430         .curr_id_lbl_line_dim_hi.Left = 0&
10440         If blnSortHere = True Then
10450           blnResort = True
10460           .Sort_line.Visible = False
10470           .Sort_lbl.Visible = False
10480         End If
10490         blnSortHere = False
10500         If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
10510         .Cost.Left = (lngCost_Left + lngCurrID_Width)
10520         .cost_str.Left = .Cost.Left
10530         .cost_forex.Left = (.cost_str.Left - lngTpp)
10540         .cost_box.Left = .Cost.Left
10550         Select Case gblnHasForExThis
              Case True
10560           .cost_str.Visible = True
10570           .Cost.Visible = False
10580           .cost_forex.Visible = True
10590         Case False
10600           Select Case .Parent.chkIncludeCurrency
                Case True
10610             .cost_str.Visible = True
10620             .Cost.Visible = False
10630             .cost_forex.Visible = True
10640           Case False
10650             .Cost.Visible = True
10660             .cost_str.Visible = False
10670             .cost_forex.Visible = False
10680           End Select
10690         End Select
10700         .cost_lbl.Left = (.Cost.Left - lngICashLbl_Offset)
10710         .cost_lbl_dim_hi.Left = (.cost_lbl.Left + lngTpp)
10720         .cost_lbl_line.Left = .cost_lbl.Left
10730         .cost_lbl_line_dim_hi.Left = (.cost_lbl_line.Left + lngTpp)
10740         If blnSortHere = True Then
10750           .Sort_line.Left = .cost_lbl.Left
10760           .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
10770         End If
10780         blnSortHere = False
10790         If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
10800         .PCash.Left = (lngPCash_Left + lngCurrID_Width)
10810         .pcash_str.Left = .PCash.Left
10820         .pcash_forex.Left = (.pcash_str.Left - lngTpp)
10830         .pcash_box.Left = .PCash.Left
10840         Select Case gblnHasForExThis
              Case True
10850           .pcash_str.Visible = True
10860           .PCash.Visible = False
10870           .pcash_forex.Visible = True
10880         Case False
10890           Select Case .Parent.chkIncludeCurrency
                Case True
10900             .pcash_str.Visible = True
10910             .PCash.Visible = False
10920             .pcash_forex.Visible = True
10930           Case False
10940             .PCash.Visible = True
10950             .pcash_str.Visible = False
10960             .pcash_forex.Visible = False
10970           End Select
10980         End Select
10990         .pcash_lbl.Left = (.PCash.Left - lngICashLbl_Offset)
11000         .pcash_lbl_dim_hi.Left = (.pcash_lbl.Left + lngTpp)
11010         .pcash_lbl_line.Left = .pcash_lbl.Left
11020         .pcash_lbl_line_dim_hi.Left = (.pcash_lbl_line.Left + lngTpp)
11030         If blnSortHere = True Then
11040           .Sort_line.Left = .pcash_lbl.Left
11050           .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
11060         End If
11070         blnSortHere = False
11080         If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
11090         .ICash.Left = (lngICash_Left + lngCurrID_Width)
11100         .icash_str.Left = .ICash.Left
11110         .icash_forex.Left = (.icash_str.Left - lngTpp)
11120         .icash_box.Left = .ICash.Left
11130         Select Case gblnHasForExThis
              Case True
11140           .icash_str.Visible = True
11150           .ICash.Visible = False
11160           .icash_forex.Visible = True
11170         Case False
11180           Select Case .Parent.chkIncludeCurrency
                Case True
11190             .icash_str.Visible = True
11200             .ICash.Visible = False
11210             .icash_forex.Visible = True
11220           Case False
11230             .ICash.Visible = True
11240             .icash_str.Visible = False
11250             .icash_forex.Visible = False
11260           End Select
11270         End Select
11280         .icash_lbl.Left = (.ICash.Left - lngICashLbl_Offset)
11290         .icash_lbl_dim_hi.Left = (.icash_lbl.Left + lngTpp)
11300         .icash_lbl_line.Left = .icash_lbl.Left
11310         .icash_lbl_line_dim_hi.Left = (.icash_lbl_line.Left + lngTpp)
11320         If blnSortHere = True Then
11330           .Sort_line.Left = .icash_lbl.Left
11340           .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
11350         End If
11360         blnSortHere = False
11370         If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
11380         .shareface_0.Left = (lngShareface_Left + lngCurrID_Width)
11390         .shareface_1.Left = .shareface_0.Left
11400         .shareface_2.Left = .shareface_0.Left
11410         .shareface_3.Left = .shareface_0.Left
11420         .shareface_4.Left = .shareface_0.Left
11430         .shareface_box.Left = .shareface_0.Left
11440         .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
11450         .shareface_lbl_line.Left = .shareface_lbl.Left
11460         .shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
11470         .shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
11480         If blnSortHere = True Then
11490           .Sort_line.Left = .shareface_lbl_line.Left
11500           .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
11510         End If
11520         blnSortHere = False
11530         If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
11540         .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width + lngCurrID_Width)
11550         .totdesc_lbl.Width = .totdesc.Width
11560         .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
11570         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
11580         .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
11590         If blnSortHere = True Then
11600           .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
11610           .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
11620         End If
11630         blnSortHere = False

11640       End Select

11650       If blnResort = True Then
11660         SortNow "Form_Load"  ' ** Procedure: Below.
11670       End If

11680     End Select

11690   End With

EXITP:
11700   Exit Sub

ERRH:
11710   DoCmd.Hourglass False
11720   Select Case ERR.Number
        Case Else
11730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11740   End Select
11750   Resume EXITP

End Sub

Public Function RecCnt() As Long

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim varTmp00 As Variant
        Dim lngRetVal As Long

11810   lngRetVal = 0&

11820   With Me
11830     Set rst = .RecordsetClone
11840     With rst
11850       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
11860       Else
11870         .MoveLast
11880         lngRetVal = .RecordCount
11890       End If
11900       .Close
11910     End With
11920     Set rst = Nothing
11930     If lngRetVal = 0& Then
            ' ** For some reason this seems to give an erroneous count!
11940       varTmp00 = DCount("*", "tmpAccountHideTrans_03")
11950       lngRetVal = Nz(varTmp00, 0)
11960     End If
11970   End With

EXITP:
11980   Set rst = Nothing
11990   RecCnt = lngRetVal
12000   Exit Function

ERRH:
12010   lngRetVal = 0&
12020   Select Case ERR.Number
        Case Else
12030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12040   End Select
12050   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

12110   Select Case intWhere
        Case acCmdRecordsGoToFirst
12120     DoCmd.RunCommand acCmdRecordsGoToFirst
12130   Case acCmdRecordsGoToLast
12140     DoCmd.RunCommand acCmdRecordsGoToLast
12150   Case acCmdRecordsGoToPrevious
12160     DoCmd.RunCommand acCmdRecordsGoToPrevious
12170   Case acCmdRecordsGoToNext
12180     DoCmd.RunCommand acCmdRecordsGoToNext
12190   Case Else
12200     If IsMissing(varID) = False Then
12210       With Me
12220         Set rst = .RecordsetClone
12230         With rst
12240           .FindFirst "[journalno] = " & CStr(varID)
12250           If .NoMatch = False Then
12260             Me.Bookmark = .Bookmark
12270           End If
12280           .Close
12290         End With
12300       End With
12310     End If
12320   End Select

EXITP:
12330   Set rst = Nothing
12340   Exit Sub

ERRH:
12350   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
12360   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
12370   Case Else
12380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12390   End Select
12400   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [JournalType_Order], [journalno]"

12510   With Me
12520     lngRecsCur = RecCnt  ' ** Function: Above.
12530     If lngRecsCur > 0& Then
12540       If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
12550         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
12560       End If
12570       .Controls(strSortLbl).Visible = False
12580       .Controls(strSortLbl).Top = lngSortLbl_Top
12590       .Controls(strSortLine).Visible = False
12600       .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
12610       .journalno_Sort_line1.Visible = False
12620       .journalno_Sort_line2.Visible = False
12630       If .ledger_HIDDEN_lbl.Left <> (.ledger_HIDDEN.Left - lngLedgerHidLbl_Offset) Then
12640         .ledger_HIDDEN_lbl.Left = (.ledger_HIDDEN.Left - lngLedgerHidLbl_Offset)
12650       End If
12660       DoEvents
12670       If strProc = "Form_Load" Then
12680         strCalled = "transdate"
12690         strSortNow = strSortOrig
12700         lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
12710         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
12720         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
12730         .Controls(strSortLbl).Left = lngSortLbl_Left
12740         .Controls(strSortLine).Left = lngSortLine_Left
12750         .Controls(strSortLine).Width = lngSortLine_Width
12760         .Controls(strSortLbl).Caption = strArwDn
12770         .Controls(strSortLbl).ForeColor = CLR_DKRED
12780         DoEvents
12790       Else
12800         strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: revcode_ID_lbl_DblClick
12810         lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
12820         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
12830         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
12840         .Controls(strSortLbl).Left = lngSortLbl_Left
12850         .Controls(strSortLine).Left = lngSortLine_Left
12860         .Controls(strSortLine).Width = lngSortLine_Width
12870         If strCalled = "transdate" Then
12880           strSortAsc = strSortOrig
12890           If strSortNow = strSortAsc Then
12900             strSortNow = "[transdate], [JournalType_Order], [journalno]"
12910             .Controls(strSortLbl).Caption = strArwUp
12920             .Controls(strSortLbl).ForeColor = CLR_DKBLU
12930           Else
12940             strSortNow = strSortAsc
12950             .Controls(strSortLbl).Caption = strArwDn
12960             .Controls(strSortLbl).ForeColor = CLR_DKRED
12970           End If
12980           DoEvents
12990         Else
13000           .Controls(strSortLbl).Caption = strArwUp
13010           .Controls(strSortLbl).ForeColor = CLR_DKBLU
13020           Select Case strCalled
                Case "journalno"
13030             strSortAsc = "[journalno]"
13040             lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
13050             .Controls(strSortLbl).Top = (lngSortLbl_Top - (2& * lngTpp))
13060             .Controls(strSortLbl).Left = lngSortLbl_Left
13070           Case "journaltype"
13080             strSortAsc = "[journaltype], [transdate], [journalno]"
13090           Case "shareface"
13100             strSortAsc = "[shareface]" & strStdAsc
13110           Case "totdesc"
13120             strSortAsc = "[totdesc]" & strStdAsc
13130           Case "icash"
13140             strSortAsc = "[icash]" & strStdAsc
13150           Case "pcash"
13160             strSortAsc = "[pcash]" & strStdAsc
13170           Case "cost"
13180             strSortAsc = "[cost]" & strStdAsc
13190           Case "curr_id"
13200             strSortAsc = "[curr_code]" & strStdAsc
13210             lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
13220             lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
13230             lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
13240             .Controls(strSortLbl).Left = lngSortLbl_Left
13250             .Controls(strSortLine).Left = lngSortLine_Left
13260             .Controls(strSortLine).Width = lngSortLine_Width
13270           Case "ledger_HIDDEN"
13280             strSortAsc = "[ledger_HIDDEN]" & strStdAsc
13290             lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
13300             lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
13310             lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
13320             .Controls(strSortLbl).Left = lngSortLbl_Left
13330             .Controls(strSortLine).Left = lngSortLine_Left
13340             .Controls(strSortLine).Width = lngSortLine_Width
13350           End Select
13360           If strSortNow = strSortAsc Then
13370             intCnt = CharCnt(strSortAsc, ",") + 1
13380             Select Case intCnt
                  Case 1
13390               strTmp01 = strSortAsc & " DESC"
13400             Case 2
13410               intPos01 = InStr(strSortAsc, ",")
13420               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
13430               strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
13440             Case 3
13450               intPos01 = InStr(strSortAsc, ",")
13460               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
13470               strTmp02 = Mid(strSortAsc, intPos01)
13480               intPos01 = InStr(2, strTmp02, ",")
13490               strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
13500               strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
13510             Case 4
13520               intPos01 = InStr(strSortAsc, ",")
13530               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
13540               strTmp02 = Mid(strSortAsc, intPos01)
13550               intPos01 = InStr(2, strTmp02, ",")
13560               strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
13570               strTmp02 = Mid(strTmp02, intPos01)
13580               intPos01 = InStr(2, strTmp02, ",")
13590               strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
13600               strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
13610             Case 5
13620               intPos01 = InStr(strSortAsc, ",")
13630               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
13640               strTmp02 = Mid(strSortAsc, intPos01)
13650               intPos01 = InStr(2, strTmp02, ",")
13660               strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
13670               strTmp02 = Mid(strTmp02, intPos01)
13680               intPos01 = InStr(2, strTmp02, ",")
13690               strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
13700               strTmp02 = Mid(strTmp02, intPos01)
13710               intPos01 = InStr(2, strTmp02, ",")
13720               strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
13730               strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
13740             End Select
13750             strSortNow = strTmp01
13760             .Controls(strSortLbl).Caption = strArwDn
13770             .Controls(strSortLbl).ForeColor = CLR_DKRED
13780           Else
13790             strSortNow = strSortAsc
13800           End If
13810           DoEvents
13820         End If
13830       End If
13840       If strCalled = "journalno" Then
13850         .journalno_Sort_line1.Visible = True
13860         .journalno_Sort_line2.Visible = True
13870       Else
13880         .Controls(strSortLine).Visible = True
13890       End If
13900       .Controls(strSortLbl).Visible = True
13910       DoEvents
13920       .OrderBy = strSortNow
13930       .OrderByOn = True
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

Public Function SortNow_Get() As String

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

14110   strRetVal = strSortNow

EXITP:
14120   SortNow_Get = strRetVal
14130   Exit Function

ERRH:
14140   Select Case ERR.Number
        Case Else
14150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14160   End Select
14170   Resume EXITP

End Function
