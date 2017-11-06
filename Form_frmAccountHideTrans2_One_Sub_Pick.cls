VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans2_One_Sub_Pick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans2_One_Sub_Pick"

'VGC 07/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Show All:         A {chkShowAll on frmAccountHideTrans2_One}
' **   Close:            C {cmdClose on frmAccountHideTrans2_One}
' **   List Hidden:      D {ledger_HIDDEN on frmAccountHideTrans2_One_Sub_List}
' **   Cancel:           E {cmdCancel on frmAccountHideTrans2_One}
' **   From:             F {DateStart on frmAccountHideTrans2_One}
' **   Include Archive:  H {chkIncludeArchive on frmAccountHideTrans2_One}
' **   CUSIP:            I {opgAssetSource_optCusip on frmAccountHideTrans2_One}
' **   Show JournalNo:   J {chkShowJournalNo on frmAccountHideTrans2_One}
' **   Journal Type:     L {cmbJournalTypes on frmAccountHideTrans2_One}
' **   Name:             M {opgAssetSource_optName on frmAccountHideTrans2_One}
' **   Pick Hidden:      N {ledger_HIDDEN}
' **   OK:               O {cmdOK on frmAccountHideTrans2_One}
' **   Clear All:        R {cmdClearAll on frmAccountHideTrans2_One}
' **   Asset:            S {cmbAssets on frmAccountHideTrans2_One}
' **   To:               T {DateEnd on frmAccountHideTrans2_One}
' **   Include Currency: U {chkIncludeCurrency on frmAccountHideTrans2_One}
' **   Exit:             X {cmdClose on frmAccountHideTrans2_One}
' **   Type:             Y {opgAssetSource_optType on frmAccountHideTrans2_One}

' ** Shortcut Ctrl keys responsive from this form:
' **   Help:             H {cmdHelp on frmAccountHideTrans2_One}
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private lngForm_Width As Long, lngJournalNo_Width As Long, lngCurrID_Width As Long
Private lngJournalNo_Left As Long, lngTransdate_Left As Long, lngShareface_Left As Long, lngTotDesc_Left As Long, lngTotDesc_Width As Long
Private lngICash_Left As Long, lngICashLbl_Offset As Long, lngPCash_Left As Long, lngCost_Left As Long
Private lngCurrID_Left As Long, lngCurrIDLbl_Offset As Long, lngJournalType_Left As Long
Private lngLedgerHid_Left As Long, lngLedgerHidLbl_Offset As Long, lngLedgerHidLblLine_Offset As Long

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
150       If strCallingForm = "frmAccountHideTrans2_One" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         .form_width_line.Left = 0&  ' ** Just get it out of the way.

180         lngForm_Width = .Width
190         lngJournalNo_Left = .journalno.Left
200         lngTransdate_Left = .transdate.Left
210         lngJournalType_Left = .journaltype.Left
220         lngShareface_Left = .shareface_0.Left
230         lngTotDesc_Left = .totdesc.Left
240         lngTotDesc_Width = .totdesc.Width
250         lngICash_Left = .icash_str.Left
260         lngICashLbl_Offset = (lngICash_Left - .icash_str_lbl.Left)
270         lngPCash_Left = .pcash_str.Left
280         lngCost_Left = .cost_str.Left
290         lngCurrID_Left = .curr_id.Left
300         lngCurrIDLbl_Offset = (lngCurrID_Left - .curr_id_lbl.Left)
310         lngLedgerHid_Left = .ledger_HIDDEN.Left
320         lngLedgerHidLbl_Offset = (lngLedgerHid_Left - .ledger_HIDDEN_lbl.Left)
330         lngLedgerHidLblLine_Offset = (lngLedgerHid_Left - .ledger_HIDDEN_lbl_line.Left)
340         lngJournalNo_Width = (lngTransdate_Left - lngJournalNo_Left)
350         lngCurrID_Width = ((lngLedgerHid_Left - lngLedgerHidLblLine_Offset) - lngCurrID_Left)

360         lngRecsCur = RecCnt  ' ** Function: Below.
370         If lngRecsCur > 0& Then
380           .FocusHolder2.SetFocus
390         End If

400       Else
410         Cancel = -1
420       End If
430     End With

440     If Cancel = -1 Then
450       Beep
460       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
470     End If

EXITP:
480     Exit Sub

ERRH:
490     Select Case ERR.Number
        Case Else
500       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
510     End Select
520     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me

620       CLR_DISABLED_FG = CLR_DKGRY
630       CLR_DISABLED_BG = CLR_LTTEAL

640       .journalno.ForeColor = CLR_DISABLED_FG
650       .journalno.BackColor = CLR_DISABLED_BG
660       .transdate.ForeColor = CLR_DISABLED_FG
670       .transdate.BackColor = CLR_DISABLED_BG
680       .journaltype.ForeColor = CLR_DISABLED_FG
690       .journaltype.BackColor = CLR_DISABLED_BG
700       .shareface_box.BackColor = CLR_DISABLED_BG
710       .shareface_0.ForeColor = CLR_DISABLED_FG
720       .shareface_1.ForeColor = CLR_DISABLED_FG
730       .shareface_2.ForeColor = CLR_DISABLED_FG
740       .shareface_3.ForeColor = CLR_DISABLED_FG
750       .shareface_4.ForeColor = CLR_DISABLED_FG
760       .totdesc.ForeColor = CLR_DISABLED_FG
770       .totdesc.BackColor = CLR_DISABLED_BG
780       .icash_str.ForeColor = CLR_DISABLED_FG
790       .icash_str.BackColor = CLR_DISABLED_BG
800       .pcash_str.ForeColor = CLR_DISABLED_FG
810       .pcash_str.BackColor = CLR_DISABLED_BG
820       .cost_str.ForeColor = CLR_DISABLED_FG
830       .cost_str.BackColor = CLR_DISABLED_BG
840       .curr_id.ForeColor = CLR_DISABLED_FG
850       .curr_id.BackColor = CLR_DISABLED_BG

860       LabelEnable False  ' ** Procedure: Below

870       .Parent.cmdOK.Enabled = False
880       .Parent.cmdCancel.Enabled = False

890       .Parent.shareface_subtot_box1.Visible = False
900       .Parent.shareface_subtot_box2.Visible = False
910       .Parent.icash_subtot_box1.Visible = False
920       .Parent.icash_subtot_box2.Visible = False
930       .Parent.pcash_subtot_box1.Visible = False
940       .Parent.pcash_subtot_box2.Visible = False
950       .Parent.cost_subtot_box1.Visible = False
960       .Parent.cost_subtot_box2.Visible = False
970       .Parent.curr_id_box1.Visible = False
980       .Parent.curr_id_box2.Visible = False

990       .Parent.shareface_subtot.Enabled = False
1000      .Parent.shareface_subtot.Locked = False
1010      .Parent.shareface_subtot = 0#
1020      .Parent.shareface_subtot.BorderColor = WIN_CLR_DISR
1030      .Parent.shareface_subtot.BackStyle = acBackStyleTransparent
1040      .Parent.shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
1050      .Parent.shareface_subtot_lbl_dim_hi.Visible = True
1060      .Parent.icash_subtot.Enabled = False
1070      .Parent.icash_subtot.Locked = False
1080      .Parent.icash_subtot = 0@
1090      .Parent.icash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
1100      .Parent.icash_subtot.BorderColor = WIN_CLR_DISR
1110      .Parent.icash_subtot.BackStyle = acBackStyleTransparent
1120      .Parent.icash_subtot_lbl.ForeColor = WIN_CLR_DISF
1130      .Parent.icash_subtot_lbl_dim_hi.Visible = True
1140      .Parent.pcash_subtot.Enabled = False
1150      .Parent.pcash_subtot.Locked = False
1160      .Parent.pcash_subtot = 0@
1170      .Parent.pcash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
1180      .Parent.pcash_subtot.BorderColor = WIN_CLR_DISR
1190      .Parent.pcash_subtot.BackStyle = acBackStyleTransparent
1200      .Parent.pcash_subtot_lbl.ForeColor = WIN_CLR_DISF
1210      .Parent.pcash_subtot_lbl_dim_hi.Visible = True
1220      .Parent.cost_subtot.Enabled = False
1230      .Parent.cost_subtot.Locked = False
1240      .Parent.cost_subtot = 0@
1250      .Parent.cost_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
1260      .Parent.cost_subtot.BorderColor = WIN_CLR_DISR
1270      .Parent.cost_subtot.BackStyle = acBackStyleTransparent
1280      .Parent.cost_subtot_lbl.ForeColor = WIN_CLR_DISF
1290      .Parent.cost_subtot_lbl_dim_hi.Visible = True

1300      If gblnHasForExThis = True Then
1310        .Parent.curr_id = Null
1320        .Parent.curr_id.BorderColor = WIN_CLR_DISR
1330        .Parent.curr_id.BackStyle = acBackStyleTransparent
1340        .Parent.curr_id.Visible = True
1350        .Parent.curr_id_lbl.ForeColor = WIN_CLR_DISR
1360        .Parent.curr_id_lbl.Visible = True
1370        .Parent.curr_id_lbl_dim_hi.Visible = True
1380      Else
1390        Select Case .Parent.chkIncludeCurrency
            Case True
1400          .Parent.curr_id = Null
1410          .Parent.curr_id.BorderColor = WIN_CLR_DISR
1420          .Parent.curr_id.BackStyle = acBackStyleTransparent
1430          .Parent.curr_id.Visible = True
1440          .Parent.curr_id_lbl.ForeColor = WIN_CLR_DISR
1450          .Parent.curr_id_lbl.Visible = True
1460          .Parent.curr_id_lbl_dim_hi.Visible = True
1470        Case False
1480          .Parent.curr_id.Visible = False
1490          .Parent.curr_id_lbl.Visible = False
1500          .Parent.curr_id_lbl_dim_hi.Visible = False
1510        End Select
1520      End If

1530      .Parent.subtot_lbl_dim.Visible = True
1540      .Parent.subtot_lbl_dim_hi.Visible = True
1550      .Parent.subtot_lbl.Visible = False

1560      .Parent.pending_trans.Enabled = False
1570      .Parent.pending_trans.Locked = False
1580      .Parent.pending_trans.BorderColor = WIN_CLR_DISR
1590      .Parent.pending_trans.BackStyle = acBackStyleTransparent

1600    End With

EXITP:
1610    Exit Sub

ERRH:
1620    Select Case ERR.Number
        Case Else
1630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1640    End Select
1650    Resume EXITP

End Sub

Private Sub Form_Current()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

1710    With Me
1720      Select Case .sort
          Case 1
1730        .ledger_HIDDEN.Locked = False
1740      Case 2
1750        .ledger_HIDDEN.Locked = True
1760      End Select
1770    End With

EXITP:
1780    Exit Sub

ERRH:
1790    Select Case ERR.Number
        Case Else
1800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1810    End Select
1820    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1910    intShiftDown = (Shift And acShiftMask) > 0
1920    intAltDown = (Shift And acAltMask) > 0
1930    intCtrlDown = (Shift And acCtrlMask) > 0

1940    intRetVal = KeyCode

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Hidden:           D {ledger_HIDDEN on frmAccountHideTrans2_One_Sub_List}
        ' **   Pick Hidden:      N {ledger_HIDDEN}
        ' **   Exit:             X {cmdClose on frmAccountHideTrans2_One}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Help:             H {cmdHelp on frmAccountHideTrans2_One}
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyEscape
1970        With Me
1980          intRetVal = 0
1990          DoCmd.SelectObject acForm, .Parent.Name, False
2000          .Parent.cmdClose.SetFocus
2010          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountHideTrans2_One.
2020        End With
2030      Case vbKeyUp
2040        intRetVal = 0
2050        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2060      Case vbKeyDown
2070        intRetVal = 0
2080        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2090      End Select
2100    End If

        ' ** Alt keys.
2110    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2120      Select Case intRetVal
          Case vbKeyD
2130        With Me
2140          intRetVal = 0
2150          DoCmd.SelectObject acForm, .Parent.Name, False
2160          .Parent.frmAccountHideTrans2_One_Sub_List.SetFocus
2170          .Parent.frmAccountHideTrans2_One_Sub_List.Form.ledger_HIDDEN.SetFocus  ' ** Don't invoke.
2180        End With
2190      Case vbKeyN
2200        With Me
2210          intRetVal = 0
2220          .ledger_HIDDEN.SetFocus  ' ** Don't invoke.
2230        End With
2240      Case vbKeyX
2250        With Me
2260          intRetVal = 0
2270          DoCmd.SelectObject acForm, .Parent.Name, False
2280          .Parent.cmdClose.SetFocus
2290          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountHideTrans2_One.
2300        End With
2310      End Select
2320    End If

        ' ** Ctrl keys.
2330    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2340      Select Case intRetVal
          Case vbKeyH
2350        With Me
2360          intRetVal = 0
2370          DoCmd.SelectObject acForm, .Parent.Name, False
2380          .Parent.cmdHelp.SetFocus
2390          .Parent.cmdHelp_Click  ' ** Form Procedure: frmAccountHideTrans2_One.
2400        End With
2410      Case vbKeyS
2420        intRetVal = 0
2430        cmdSave_Click  ' ** Procedure: Below.
2440      Case vbKeyTab
2450        With Me
2460          intRetVal = 0
2470          DoCmd.SelectObject acForm, .Parent.Name, False
2480          If .Parent.cmdOK.Enabled = True Then
2490            .Parent.cmdOK.SetFocus
2500          ElseIf Parent.cmdCancel.Enabled = True Then
2510            .Parent.cmdCancel.SetFocus
2520          ElseIf .Parent.chkIncludeArchive.Enabled = True Then
2530            .Parent.chkIncludeArchive.SetFocus
2540          ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
2550            .Parent.chkIncludeCurrency.SetFocus
2560          Else
2570            .Parent.chkShowJournalNo.SetFocus
2580          End If
2590        End With
2600      Case vbKeyPageUp, vbKeyHome
2610        intRetVal = 0
2620        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2630      Case vbKeyPageDown, vbKeyEnd
2640        intRetVal = 0
2650        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2660      End Select
2670    End If

        ' ** Ctrl-Shift keys.
2680    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2690      Select Case intRetVal
          Case vbKeyF
2700        With Me
2710          intRetVal = 0
2720          DoCmd.SelectObject acForm, .Parent.Name, False
2730          .Parent.FocusHolder.SetFocus
2740        End With
2750      Case vbKeyTab
2760        With Me
2770          intRetVal = 0
2780          DoCmd.SelectObject acForm, .Parent.Name, False
2790          lngRecsCur = .Parent.frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
2800          If lngRecsCur > 0& Then
2810            .Parent.frmAccountHideTrans2_One_Sub_List.SetFocus
2820            .Parent.frmAccountHideTrans2_One_Sub_List.Form.FocusHolder.SetFocus
2830          Else
2840            .Parent.cmdClose.SetFocus
2850          End If
2860        End With
2870      End Select
2880    End If

EXITP:
2890    KeyCode = intRetVal
2900    Exit Sub

ERRH:
2910    intRetVal = 0
2920    Select Case ERR.Number
        Case Else
2930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2940    End Select
2950    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3010    DoCmd.RunCommand acCmdSaveRecord
3020    DoEvents

EXITP:
3030    Exit Sub

ERRH:
3040    Select Case ERR.Number
        Case Else
3050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3060    End Select
3070    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
3110    intShiftDown = (Shift And acShiftMask) > 0
3120    intAltDown = (Shift And acAltMask) > 0
3130    intCtrlDown = (Shift And acCtrlMask) > 0

3140    intRetVal = KeyCode

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3170        With Me
3180          intRetVal = 0
3190          .ledger_HIDDEN.SetFocus
3200        End With
3210      End Select
3220    End If

        ' ** Shift keys.
3230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3250        With Me
3260          intRetVal = 0
3270          If .CurrentRecord > 1 Then
3280            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3290            .ledger_HIDDEN.SetFocus
3300          Else
3310            DoCmd.SelectObject acForm, .Parent.Name, False
3320            lngRecsCur = .Parent.frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_List.
3330            If lngRecsCur > 0& Then
3340              .Parent.frmAccountHideTrans2_One_Sub_List.SetFocus
3350              .Parent.frmAccountHideTrans2_One_Sub_List.Form.FocusHolder.SetFocus
3360            Else
3370              .Parent.cmdClose.SetFocus
3380            End If
3390          End If
3400        End With
3410      End Select
3420    End If

EXITP:
3430    KeyCode = intRetVal
3440    Exit Sub

ERRH:
3450    intRetVal = 0
3460    Select Case ERR.Number
        Case Else
3470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3480    End Select
3490    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_BeforeUpdate(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_BeforeUpdate"

3510    With Me
3520      If .sort = 2 Then
            ' ** They can't change the status of the reversal transaction.
3530        Cancel = -1
3540      End If
3550    End With

EXITP:
3560    Exit Sub

ERRH:
3570    Select Case ERR.Number
        Case Else
3580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3590    End Select
3600    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_AfterUpdate()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngJournalNo As Long

3710    With Me
3720      If .sort = 1 Then  ' ** The original transaction to be reversed.

3730        Select Case .ledger_HIDDEN
            Case True
              ' ** It shouldn't ever be switching from False to True.
3740        Case False
              ' ** Cancel the selection.

3750          DoCmd.Hourglass True
3760          DoEvents

3770          lngJournalNo = .journalno

3780          .FocusHolder2.SetFocus

3790          DoCmd.SelectObject acForm, .Parent.Name, False
3800          .Parent.frmAccountHideTrans2_One_Sub_List.SetFocus
3810          DoEvents

3820          Set dbs = CurrentDb
3830          With dbs
                ' ** Empty tblLedgerHidden_Staging2.
3840            Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_03")
3850            qdf.Execute
3860            Set qdf = Nothing
                ' ** Update tblLedgerHidden_Staging1, for ledger_HIDDEN = False.
3870            Set qdf = .QueryDefs("qryAccountHideTrans2_One_03")
3880            qdf.Execute
3890            Set qdf = Nothing
3900            .Close
3910          End With
3920          Set dbs = Nothing

3930          .Parent.frmAccountHideTrans2_One_Sub_List.Form.Requery
3940          .Parent.frmAccountHideTrans2_One_Sub_List.Form.MoveRec 0, lngJournalNo  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_List.
3950          .Parent.frmAccountHideTrans2_One_Sub_List.Form.ledger_HIDDEN.SetFocus
3960          DoEvents

3970          .Requery
3980          LabelEnable False  ' ** Procedure: Below.

3990          .Parent.cmdOK.Enabled = False
4000          .Parent.cmdCancel.Enabled = False

4010          .Parent.shareface_subtot_box1.Visible = False
4020          .Parent.shareface_subtot_box2.Visible = False
4030          .Parent.icash_subtot_box1.Visible = False
4040          .Parent.icash_subtot_box2.Visible = False
4050          .Parent.pcash_subtot_box1.Visible = False
4060          .Parent.pcash_subtot_box2.Visible = False
4070          .Parent.cost_subtot_box1.Visible = False
4080          .Parent.cost_subtot_box2.Visible = False
4090          .Parent.curr_id_box1.Visible = False
4100          .Parent.curr_id_box2.Visible = False

4110          .Parent.shareface_subtot.Visible = True
4120          .Parent.shareface_divint_subtot.Visible = False
4130          .Parent.shareface_subtot.Enabled = False
4140          .Parent.shareface_subtot.Locked = False
4150          .Parent.shareface_subtot = 0#
4160          .Parent.shareface_subtot.BorderColor = WIN_CLR_DISR
4170          .Parent.shareface_subtot.BackStyle = acBackStyleTransparent
4180          .Parent.shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
4190          .Parent.shareface_subtot_lbl_dim_hi.Visible = True
4200          .Parent.icash_subtot.Enabled = False
4210          .Parent.icash_subtot.Locked = False
4220          .Parent.icash_subtot = 0@
4230          .Parent.icash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
4240          .Parent.icash_subtot.BorderColor = WIN_CLR_DISR
4250          .Parent.icash_subtot.BackStyle = acBackStyleTransparent
4260          .Parent.icash_subtot_lbl.ForeColor = WIN_CLR_DISF
4270          .Parent.icash_subtot_lbl_dim_hi.Visible = True
4280          .Parent.pcash_subtot.Enabled = False
4290          .Parent.pcash_subtot.Locked = False
4300          .Parent.pcash_subtot = 0@
4310          .Parent.pcash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
4320          .Parent.pcash_subtot.BorderColor = WIN_CLR_DISR
4330          .Parent.pcash_subtot.BackStyle = acBackStyleTransparent
4340          .Parent.pcash_subtot_lbl.ForeColor = WIN_CLR_DISF
4350          .Parent.pcash_subtot_lbl_dim_hi.Visible = True
4360          .Parent.cost_subtot.Enabled = False
4370          .Parent.cost_subtot.Locked = False
4380          .Parent.cost_subtot = 0@
4390          .Parent.cost_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
4400          .Parent.cost_subtot.BorderColor = WIN_CLR_DISR
4410          .Parent.cost_subtot.BackStyle = acBackStyleTransparent
4420          .Parent.cost_subtot_lbl.ForeColor = WIN_CLR_DISF
4430          .Parent.cost_subtot_lbl_dim_hi.Visible = True

4440          If gblnHasForExThis = True Then
4450            .Parent.curr_id = Null
4460            .Parent.curr_id.BorderColor = WIN_CLR_DISR
4470            .Parent.curr_id.BackStyle = acBackStyleTransparent
4480            .Parent.curr_id.Visible = True
4490            .Parent.curr_id_lbl.ForeColor = WIN_CLR_DISR
4500            .Parent.curr_id_lbl.Visible = True
4510            .Parent.curr_id_lbl_dim_hi.Visible = True
4520          Else
4530            Select Case .Parent.chkIncludeCurrency
                Case True
4540              .Parent.curr_id = Null
4550              .Parent.curr_id.BorderColor = WIN_CLR_DISR
4560              .Parent.curr_id.BackStyle = acBackStyleTransparent
4570              .Parent.curr_id.Visible = True
4580              .Parent.curr_id_lbl.ForeColor = WIN_CLR_DISR
4590              .Parent.curr_id_lbl.Visible = True
4600              .Parent.curr_id_lbl_dim_hi.Visible = True
4610            Case False
4620              .Parent.curr_id.Visible = False
4630              .Parent.curr_id_lbl.Visible = False
4640              .Parent.curr_id_lbl_dim_hi.Visible = False
4650            End Select
4660          End If

4670          .Parent.subtot_lbl_dim.Visible = True
4680          .Parent.subtot_lbl_dim_hi.Visible = True
4690          .Parent.subtot_lbl.Visible = False

4700          .Parent.pending_trans.Enabled = False
4710          .Parent.pending_trans.Locked = False
4720          .Parent.pending_trans.BorderColor = WIN_CLR_DISR
4730          .Parent.pending_trans.BackStyle = acBackStyleTransparent
4740          .Parent.pending_trans = 0

4750          DoCmd.Hourglass False

4760        End Select
4770      End If
4780    End With

EXITP:
4790    Set qdf = Nothing
4800    Set dbs = Nothing
4810    Exit Sub

ERRH:
4820    DoCmd.Hourglass False
4830    Select Case ERR.Number
        Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
4910    intShiftDown = (Shift And acShiftMask) > 0
4920    intAltDown = (Shift And acAltMask) > 0
4930    intCtrlDown = (Shift And acCtrlMask) > 0

4940    intRetVal = KeyCode

        ' ** Plain keys.
4950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4970        With Me
4980          intRetVal = 0
4990          lngRecsCur = RecCnt  ' ** Function: Below.
5000          If .CurrentRecord < lngRecsCur Then
5010            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5020            .ledger_HIDDEN.SetFocus
5030          Else
5040            DoCmd.SelectObject acForm, .Parent.Name, False
5050            If .Parent.cmdOK.Enabled = True Then
5060              .Parent.cmdOK.SetFocus
5070            ElseIf Parent.cmdCancel.Enabled = True Then
5080              .Parent.cmdCancel.SetFocus
5090            ElseIf .Parent.chkIncludeArchive.Enabled = True Then
5100              .Parent.chkIncludeArchive.SetFocus
5110            ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
5120              .Parent.chkIncludeCurrency.SetFocus
5130            Else
5140              .Parent.chkShowAll.SetFocus
5150            End If
5160          End If
5170        End With
5180      End Select
5190    End If

        ' ** Shift keys.
5200    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5210      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5220        With Me
5230          intRetVal = 0
5240          If .CurrentRecord > 1 Then
5250            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5260            .ledger_HIDDEN.SetFocus
5270          Else
5280            DoCmd.SelectObject acForm, .Parent.Name, False
5290            lngRecsCur = .Parent.frmAccountHideTrans2_One_Sub_List.Form.RecCnt  ' ** Form Funcstion: frmAccountHideTrans2_One_Sub_List.
5300            If lngRecsCur > 0& Then
5310              .Parent.frmAccountHideTrans2_One_Sub_List.SetFocus
5320              .Parent.frmAccountHideTrans2_One_Sub_List.Form.FocusHolder.SetFocus
5330            Else
5340              .Parent.cmdClose.SetFocus
5350            End If
5360          End If
5370        End With
5380      End Select
5390    End If

EXITP:
5400    KeyCode = intRetVal
5410    Exit Sub

ERRH:
5420    intRetVal = 0
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Sub

Public Sub LabelEnable(blnEnable As Boolean)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "LabelEnable"

5510    With Me
5520      Select Case blnEnable
          Case True
5530        .journalno_lbl.ForeColor = CLR_GRY4
5540        .journalno_lbl_dim_hi.Visible = False
5550        .journalno_lbl2.ForeColor = CLR_GRY4
5560        .journalno_lbl2_dim_hi.Visible = False
5570        .journalno_lbl_line.BorderColor = CLR_GRY5
5580        .journalno_lbl_line_dim_hi.Visible = False
5590        .transdate_lbl.ForeColor = CLR_DKGRY2
5600        .transdate_lbl_dim_hi.Visible = False
5610        .transdate_lbl_line.BorderColor = CLR_DKGRY
5620        .transdate_lbl_line_dim_hi.Visible = False
5630        .journaltype_lbl.ForeColor = CLR_DKGRY2
5640        .journaltype_lbl_dim_hi.Visible = False
5650        .journaltype_lbl_line.BorderColor = CLR_DKGRY
5660        .journaltype_lbl_line_dim_hi.Visible = False
5670        .totdesc_lbl.ForeColor = CLR_DKGRY2
5680        .totdesc_lbl_dim_hi.Visible = False
5690        .totdesc_lbl_line.BorderColor = CLR_DKGRY
5700        .totdesc_lbl_line_dim_hi.Visible = False
5710        .shareface_lbl.ForeColor = CLR_DKGRY2
5720        .shareface_lbl_dim_hi.Visible = False
5730        .shareface_lbl_line.BorderColor = CLR_DKGRY
5740        .shareface_lbl_line_dim_hi.Visible = False
5750        .icash_str_lbl.ForeColor = CLR_DKGRY2
5760        .icash_str_lbl_dim_hi.Visible = False
5770        .icash_str_lbl_line.BorderColor = CLR_DKGRY
5780        .icash_str_lbl_line_dim_hi.Visible = False
5790        .pcash_str_lbl.ForeColor = CLR_DKGRY2
5800        .pcash_str_lbl_dim_hi.Visible = False
5810        .pcash_str_lbl_line.BorderColor = CLR_DKGRY
5820        .pcash_str_lbl_line_dim_hi.Visible = False
5830        .cost_str_lbl.ForeColor = CLR_DKGRY2
5840        .cost_str_lbl_dim_hi.Visible = False
5850        .cost_str_lbl_line.BorderColor = CLR_DKGRY
5860        .cost_str_lbl_line_dim_hi.Visible = False
5870        .curr_id_lbl.ForeColor = CLR_DKGRY2
5880        .curr_id_lbl_dim_hi.Visible = False
5890        .curr_id_lbl_line.BorderColor = CLR_DKGRY
5900        .curr_id_lbl_line_dim_hi.Visible = False
5910        .ledger_HIDDEN_lbl.ForeColor = CLR_DKGRY2
5920        .ledger_HIDDEN_lbl_dim_hi.Visible = False
5930        .ledger_HIDDEN_lbl_line.BorderColor = CLR_DKGRY
5940        .ledger_HIDDEN_lbl_line_dim_hi.Visible = False
5950      Case False
5960        .journalno_lbl.ForeColor = WIN_CLR_DISF
5970        .journalno_lbl2.ForeColor = WIN_CLR_DISF
5980        .journalno_lbl_line.BorderColor = WIN_CLR_DISR
5990        Select Case .Parent.chkShowJournalNo
            Case True
6000          .journalno_lbl_dim_hi.Visible = True
6010          .journalno_lbl2_dim_hi.Visible = True
6020          .journalno_lbl_line_dim_hi.Visible = True
6030        Case False
6040          .journalno_lbl_dim_hi.Visible = False
6050          .journalno_lbl2_dim_hi.Visible = False
6060          .journalno_lbl_line_dim_hi.Visible = False
6070        End Select
6080        .transdate_lbl.ForeColor = WIN_CLR_DISF
6090        .transdate_lbl_dim_hi.Visible = True
6100        .transdate_lbl_line.BorderColor = WIN_CLR_DISR
6110        .transdate_lbl_line_dim_hi.Visible = True
6120        .journaltype_lbl.ForeColor = WIN_CLR_DISF
6130        .journaltype_lbl_dim_hi.Visible = True
6140        .journaltype_lbl_line.BorderColor = WIN_CLR_DISR
6150        .journaltype_lbl_line_dim_hi.Visible = True
6160        .totdesc_lbl.ForeColor = WIN_CLR_DISF
6170        .totdesc_lbl_dim_hi.Visible = True
6180        .totdesc_lbl_line.BorderColor = WIN_CLR_DISR
6190        .totdesc_lbl_line_dim_hi.Visible = True
6200        .shareface_lbl.ForeColor = WIN_CLR_DISF
6210        .shareface_lbl_dim_hi.Visible = True
6220        .shareface_lbl_line.BorderColor = WIN_CLR_DISR
6230        .shareface_lbl_line_dim_hi.Visible = True
6240        .icash_str_lbl.ForeColor = WIN_CLR_DISF
6250        .icash_str_lbl_dim_hi.Visible = True
6260        .icash_str_lbl_line.BorderColor = WIN_CLR_DISR
6270        .icash_str_lbl_line_dim_hi.Visible = True
6280        .pcash_str_lbl.ForeColor = WIN_CLR_DISF
6290        .pcash_str_lbl_dim_hi.Visible = True
6300        .pcash_str_lbl_line.BorderColor = WIN_CLR_DISR
6310        .pcash_str_lbl_line_dim_hi.Visible = True
6320        .cost_str_lbl.ForeColor = WIN_CLR_DISF
6330        .cost_str_lbl_dim_hi.Visible = True
6340        .cost_str_lbl_line.BorderColor = WIN_CLR_DISR
6350        .cost_str_lbl_line_dim_hi.Visible = True
6360        .curr_id_lbl.ForeColor = WIN_CLR_DISF
6370        .curr_id_lbl_line.BorderColor = WIN_CLR_DISR
6380        Select Case .Parent.chkIncludeCurrency
            Case True
6390          .curr_id_lbl_dim_hi.Visible = True
6400          .curr_id_lbl_line_dim_hi.Visible = True
6410        Case False
6420          .curr_id_lbl_dim_hi.Visible = False
6430          .curr_id_lbl_line_dim_hi.Visible = False
6440        End Select
6450        .ledger_HIDDEN_lbl.ForeColor = WIN_CLR_DISF
6460        .ledger_HIDDEN_lbl_dim_hi.Visible = True
6470        .ledger_HIDDEN_lbl_line.BorderColor = WIN_CLR_DISR
6480        .ledger_HIDDEN_lbl_line_dim_hi.Visible = True
6490      End Select
6500    End With

EXITP:
6510    Exit Sub

ERRH:
6520    Select Case ERR.Number
        Case Else
6530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6540    End Select
6550    Resume EXITP

End Sub

Public Sub ShowJournalNo_Sub(blnShow As Boolean)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo_Sub"

        Dim blnCurrency As Boolean

6610    With Me

6620      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
6630        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6640      End If

6650      blnCurrency = .Parent.chkIncludeCurrency

6660      Select Case blnShow
          Case True

6670        .totdesc.Left = lngTotDesc_Left
6680        .totdesc_lbl.Left = lngTotDesc_Left
6690        .totdesc_lbl_dim_hi.Left = (.totdesc_lbl.Left + lngTpp)
6700        .totdesc_lbl_line.Left = lngTotDesc_Left
6710        .totdesc_lbl_line_dim_hi.Left = (.totdesc_lbl_line.Left + lngTpp)
6720        Select Case blnCurrency
            Case True
6730          .totdesc.Width = lngTotDesc_Width
6740          .totdesc_lbl.Width = .totdesc.Width
6750          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
6760          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6770          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
6780        Case False
6790          .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)
6800          .totdesc_lbl.Width = .totdesc.Width
6810          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
6820          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6830          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
6840        End Select
6850        .journaltype.Left = lngJournalType_Left
6860        .journaltype_lbl.Left = lngJournalType_Left
6870        .journaltype_lbl_dim_hi.Left = (.journaltype_lbl.Left + lngTpp)
6880        .journaltype_lbl_line.Left = lngJournalType_Left
6890        .journaltype_lbl_line_dim_hi.Left = (.journaltype_lbl_line.Left + lngTpp)
6900        .transdate.Left = lngTransdate_Left
6910        .transdate_lbl.Left = lngTransdate_Left
6920        .transdate_lbl_dim_hi.Left = (.transdate_lbl.Left + lngTpp)
6930        .transdate_lbl_line.Left = lngTransdate_Left
6940        .transdate_lbl_line_dim_hi.Left = (.transdate_lbl_line.Left + lngTpp)
6950        .journalno.Visible = True
6960        .journalno_lbl.Visible = True
6970        .journalno_lbl2.Visible = True
6980        .journalno_lbl_line.Visible = True
6990        lngRecsCur = RecCnt  ' ** Function: Below.
7000        If lngRecsCur = 0& Then
7010          .journalno_lbl_dim_hi.Visible = True
7020          .journalno_lbl2_dim_hi.Visible = True
7030          .journalno_lbl_line_dim_hi.Visible = True
7040        End If

7050      Case False

7060        .journalno.Visible = False
7070        .journalno_lbl.Visible = False
7080        .journalno_lbl_dim_hi.Visible = False
7090        .journalno_lbl2.Visible = False
7100        .journalno_lbl2_dim_hi.Visible = False
7110        .journalno_lbl_line.Visible = False
7120        .journalno_lbl_line_dim_hi.Visible = False
7130        .journalno_Sort_line1.Visible = False
7140        .journalno_Sort_line2.Visible = False
7150        .transdate.Left = (lngTransdate_Left - lngJournalNo_Width)
7160        .transdate_lbl.Left = .transdate.Left
7170        .transdate_lbl_dim_hi.Left = (.transdate_lbl.Left + lngTpp)
7180        .transdate_lbl_line.Left = .transdate_lbl.Left
7190        .transdate_lbl_line_dim_hi.Left = (.transdate_lbl_line.Left + lngTpp)
7200        .journaltype.Left = (lngJournalType_Left - lngJournalNo_Width)
7210        .journaltype_lbl.Left = .journaltype.Left
7220        .journaltype_lbl_dim_hi.Left = (.journaltype_lbl.Left + lngTpp)
7230        .journaltype_lbl_line.Left = .journaltype_lbl.Left
7240        .journaltype_lbl_line_dim_hi.Left = (.journaltype_lbl_line.Left + lngTpp)
7250        .totdesc.Left = (lngTotDesc_Left - lngJournalNo_Width)
7260        .totdesc_lbl.Left = .totdesc.Left
7270        .totdesc_lbl_dim_hi.Left = (.totdesc_lbl.Left + lngTpp)
7280        .totdesc_lbl_line.Left = .totdesc_lbl.Left
7290        .totdesc_lbl_line_dim_hi.Left = (.totdesc_lbl_line.Left + lngTpp)
7300        Select Case blnCurrency
            Case True
7310          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)
7320          .totdesc_lbl.Width = .totdesc.Width
7330          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
7340          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7350          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
7360        Case False
7370          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width + lngCurrID_Width)
7380          .totdesc_lbl.Width = .totdesc.Width
7390          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
7400          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7410          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
7420        End Select

7430      End Select

7440    End With

EXITP:
7450    Exit Sub

ERRH:
7460    Select Case ERR.Number
        Case Else
7470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7480    End Select
7490    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnInclude As Boolean)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnJournalNo As Boolean

7510    With Me

7520      DoCmd.Hourglass True
7530      DoEvents

7540      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
7550        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7560      End If

7570      blnJournalNo = .Parent.chkShowJournalNo

7580      Select Case blnInclude
          Case True

7590        Select Case blnJournalNo
            Case True
7600          .totdesc.Width = lngTotDesc_Width
7610          .totdesc_lbl.Width = .totdesc.Width
7620          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
7630          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7640          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
7650        Case False
7660          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)
7670          .totdesc_lbl.Width = .totdesc.Width
7680          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
7690          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7700          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
7710        End Select
7720        .shareface_0.Left = lngShareface_Left
7730        .shareface_1.Left = .shareface_0.Left
7740        .shareface_2.Left = .shareface_0.Left
7750        .shareface_3.Left = .shareface_0.Left
7760        .shareface_4.Left = .shareface_0.Left
7770        .shareface_box.Left = .shareface_0.Left
7780        .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
7790        .shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
7800        .shareface_lbl_line.Left = .shareface_lbl.Left
7810        .shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
7820        .icash_str.Left = lngICash_Left
7830        .icash_forex.Left = (.icash_str.Left - lngTpp)
7840        .icash_box.Left = .icash_str.Left
7850        .icash_str_lbl.Left = (.icash_str.Left - lngICashLbl_Offset)
7860        .icash_str_lbl_dim_hi.Left = (.icash_str_lbl.Left + lngTpp)
7870        .icash_str_lbl_line.Left = .icash_str_lbl.Left
7880        .icash_str_lbl_line_dim_hi.Left = (.icash_str_lbl_line.Left + lngTpp)
7890        .pcash_str.Left = lngPCash_Left
7900        .pcash_forex.Left = (.pcash_str.Left - lngTpp)
7910        .pcash_box.Left = .pcash_str.Left
7920        .pcash_str_lbl.Left = (.pcash_str.Left - lngICashLbl_Offset)
7930        .pcash_str_lbl_dim_hi.Left = (.pcash_str_lbl.Left + lngTpp)
7940        .pcash_str_lbl_line.Left = .pcash_str_lbl.Left
7950        .pcash_str_lbl_line_dim_hi.Left = (.pcash_str_lbl_line.Left + lngTpp)
7960        .cost_str.Left = lngCost_Left
7970        .cost_forex.Left = (.cost_str.Left - lngTpp)
7980        .cost_box.Left = .cost_str.Left
7990        .cost_str_lbl.Left = (.cost_str.Left - lngICashLbl_Offset)
8000        .cost_str_lbl_dim_hi.Left = (.cost_str_lbl.Left + lngTpp)
8010        .cost_str_lbl_line.Left = .cost_str_lbl.Left
8020        .cost_str_lbl_line_dim_hi.Left = (.cost_str_lbl_line.Left + lngTpp)
8030        .curr_id.Left = lngCurrID_Left
8040        .curr_id_forex.Left = (.curr_id.Left - lngTpp)
8050        .curr_id_box.Left = .curr_id.Left
8060        .curr_id_lbl.Left = (lngCurrID_Left - lngCurrIDLbl_Offset)
8070        .curr_id_lbl_dim_hi.Left = (.curr_id_lbl.Left + lngTpp)
8080        .curr_id_lbl_line.Left = lngCurrID_Left
8090        .curr_id_lbl_line_dim_hi.Left = (.curr_id_lbl_line.Left + lngTpp)
8100        .curr_id.Visible = True
8110        .curr_id_forex.Visible = True
8120        .curr_id_box.Visible = True
8130        .curr_id_lbl.Visible = True
8140        .curr_id_lbl_line.Visible = True
8150        lngRecsCur = RecCnt  ' ** Function: Below.
8160        If lngRecsCur = 0& Then
8170          .curr_id_lbl_dim_hi.Visible = True
8180          .curr_id_lbl_line_dim_hi.Visible = True
8190        End If

8200      Case False

8210        Select Case blnJournalNo
            Case True
8220          .curr_id.Visible = False
8230          .curr_id_forex.Visible = False
8240          .curr_id_box.Visible = False
8250          .curr_id_lbl.Visible = False
8260          .curr_id_lbl_dim_hi.Visible = False
8270          .curr_id_lbl_line.Visible = False
8280          .curr_id_lbl_line_dim_hi.Visible = False
8290          .curr_id.Left = 0&
8300          .curr_id_forex.Left = 0&
8310          .curr_id_box.Left = 0&
8320          .curr_id_lbl.Left = 0&
8330          .curr_id_lbl_dim_hi.Left = 0&
8340          .curr_id_lbl_line.Left = 0&
8350          .curr_id_lbl_line_dim_hi.Left = 0&
8360          .cost_str.Left = (lngCost_Left + lngCurrID_Width)
8370          .cost_forex.Left = (.cost_str.Left - lngTpp)
8380          .cost_box.Left = .cost_str.Left
8390          .cost_str_lbl.Left = (.cost_str.Left - lngICashLbl_Offset)
8400          .cost_str_lbl_dim_hi.Left = (.cost_str_lbl.Left + lngTpp)
8410          .cost_str_lbl_line.Left = .cost_str_lbl.Left
8420          .cost_str_lbl_line_dim_hi.Left = (.cost_str_lbl_line.Left + lngTpp)
8430          .pcash_str.Left = (lngPCash_Left + lngCurrID_Width)
8440          .pcash_forex.Left = (.pcash_str.Left - lngTpp)
8450          .pcash_box.Left = .pcash_str.Left
8460          .pcash_str_lbl.Left = (.pcash_str.Left - lngICashLbl_Offset)
8470          .pcash_str_lbl_dim_hi.Left = (.pcash_str_lbl.Left + lngTpp)
8480          .pcash_str_lbl_line.Left = .pcash_str_lbl.Left
8490          .pcash_str_lbl_line_dim_hi.Left = (.pcash_str_lbl_line.Left + lngTpp)
8500          .icash_str.Left = (lngICash_Left + lngCurrID_Width)
8510          .icash_forex.Left = (.icash_str.Left - lngTpp)
8520          .icash_box.Left = .icash_str.Left
8530          .icash_str_lbl.Left = (.icash_str.Left - lngICashLbl_Offset)
8540          .icash_str_lbl_dim_hi.Left = (icash_str_lbl.Left + lngTpp)
8550          .icash_str_lbl_line.Left = .icash_str_lbl.Left
8560          .icash_str_lbl_line_dim_hi.Left = (.icash_str_lbl_line.Left + lngTpp)
8570          .shareface_0.Left = (lngShareface_Left + lngCurrID_Width)
8580          .shareface_1.Left = .shareface_0.Left
8590          .shareface_2.Left = .shareface_0.Left
8600          .shareface_3.Left = .shareface_0.Left
8610          .shareface_4.Left = .shareface_0.Left
8620          .shareface_box.Left = .shareface_0.Left
8630          .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
8640          .shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
8650          .shareface_lbl_line.Left = .shareface_lbl.Left
8660          .shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
8670          .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)
8680          .totdesc_lbl.Width = .totdesc.Width
8690          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
8700          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
8710          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
8720        Case False
8730          .curr_id.Visible = False
8740          .curr_id_forex.Visible = False
8750          .curr_id_box.Visible = False
8760          .curr_id_lbl.Visible = False
8770          .curr_id_lbl_dim_hi.Visible = False
8780          .curr_id_lbl_line.Visible = False
8790          .curr_id_lbl_line_dim_hi.Visible = False
8800          .curr_id.Left = 0&
8810          .curr_id_forex.Left = 0&
8820          .curr_id_box.Left = 0&
8830          .curr_id_lbl.Left = 0&
8840          .curr_id_lbl_dim_hi.Left = 0&
8850          .curr_id_lbl_line.Left = 0&
8860          .curr_id_lbl_line_dim_hi.Left = 0&
8870          .cost_str.Left = (lngCost_Left + lngCurrID_Width)
8880          .cost_forex.Left = (.cost_str.Left - lngTpp)
8890          .cost_box.Left = .cost_str.Left
8900          .cost_str_lbl.Left = (.cost_str.Left - lngICashLbl_Offset)
8910          .cost_str_lbl_dim_hi.Left = (.cost_str_lbl.Left + lngTpp)
8920          .cost_str_lbl_line.Left = .cost_str_lbl.Left
8930          .cost_str_lbl_line_dim_hi.Left = (.cost_str_lbl_line.Left + lngTpp)
8940          .pcash_str.Left = (lngPCash_Left + lngCurrID_Width)
8950          .pcash_forex.Left = (.pcash_str.Left - lngTpp)
8960          .pcash_box.Left = .pcash_str.Left
8970          .pcash_str_lbl.Left = (.pcash_str.Left - lngICashLbl_Offset)
8980          .pcash_str_lbl_dim_hi.Left = (.pcash_str_lbl.Left + lngTpp)
8990          .pcash_str_lbl_line.Left = .pcash_str_lbl.Left
9000          .pcash_str_lbl_line_dim_hi.Left = (.pcash_str_lbl_line.Left + lngTpp)
9010          .icash_str.Left = (lngICash_Left + lngCurrID_Width)
9020          .icash_forex.Left = (.icash_str.Left - lngTpp)
9030          .icash_box.Left = .icash_str.Left
9040          .icash_str_lbl.Left = (.icash_str.Left - lngICashLbl_Offset)
9050          .icash_str_lbl_dim_hi.Left = (.icash_str_lbl.Left + lngTpp)
9060          .icash_str_lbl_line.Left = .icash_str_lbl.Left
9070          .icash_str_lbl_line_dim_hi.Left = (.icash_str_lbl_line.Left + lngTpp)
9080          .shareface_0.Left = (lngShareface_Left + lngCurrID_Width)
9090          .shareface_1.Left = .shareface_0.Left
9100          .shareface_2.Left = .shareface_0.Left
9110          .shareface_3.Left = .shareface_0.Left
9120          .shareface_4.Left = .shareface_0.Left
9130          .shareface_box.Left = .shareface_0.Left
9140          .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
9150          .shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
9160          .shareface_lbl_line.Left = .shareface_lbl.Left
9170          .shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
9180          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width + lngCurrID_Width)
9190          .totdesc_lbl.Width = .totdesc.Width
9200          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
9210          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
9220          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
9230        End Select

9240      End Select

9250    End With

EXITP:
9260    Exit Sub

ERRH:
9270    DoCmd.Hourglass False
9280    Select Case ERR.Number
        Case Else
9290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9300    End Select
9310    Resume EXITP

End Sub

Public Function RecCnt() As Long

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

9410    lngRetVal = 0&

9420    With Me
9430      .Requery
9440      DoEvents
9450      Set rst = .RecordsetClone
9460      With rst
9470        If .BOF = True And .EOF = True Then
              ' ** This is giving Zero when there are recs!
9480  On Error Resume Next
9490          .MoveLast
9500          lngRetVal = .RecordCount
9510  On Error GoTo ERRH
9520        Else
9530          .MoveLast
9540          lngRetVal = .RecordCount
9550        End If
9560        .Close
9570      End With
9580    End With

EXITP:
9590    Set rst = Nothing
9600    RecCnt = lngRetVal
9610    Exit Function

ERRH:
9620    lngRetVal = 0&
9630    Select Case ERR.Number
        Case Else
9640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9650    End Select
9660    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

9710    Select Case intWhere
        Case acCmdRecordsGoToFirst
9720      DoCmd.RunCommand acCmdRecordsGoToFirst
9730    Case acCmdRecordsGoToLast
9740      DoCmd.RunCommand acCmdRecordsGoToLast
9750    Case acCmdRecordsGoToPrevious
9760      DoCmd.RunCommand acCmdRecordsGoToPrevious
9770    Case acCmdRecordsGoToNext
9780      DoCmd.RunCommand acCmdRecordsGoToNext
9790    Case Else
9800      If IsMissing(varID) = False Then
9810        With Me
9820          Set rst = .RecordsetClone
9830          With rst
9840            .FindFirst "[journalno] = " & CStr(varID)
9850            If .NoMatch = False Then
9860              Me.Bookmark = .Bookmark
9870            End If
9880            .Close
9890          End With
9900        End With
9910      End If
9920    End Select

EXITP:
9930    Set rst = Nothing
9940    Exit Sub

ERRH:
9950    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
9960    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
9970    Case Else
9980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9990    End Select
10000   Resume EXITP

End Sub
