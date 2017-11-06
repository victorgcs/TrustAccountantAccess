VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans2_One_Sub_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans2_One_Sub_List"

'VGC 07/26/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Show All:         A {chkShowAll on frmAccountHideTrans2_One}
' **   Close:            C {cmdClose on frmAccountHideTrans2_One}
' **   List Hidden:      D {ledger_HIDDEN}
' **   Cancel:           E {cmdCancel on frmAccountHideTrans2_One}
' **   From:             F {DateStart on frmAccountHideTrans2_One}
' **   Include Archive:  H {chkIncludeArchive on frmAccountHideTrans2_One}
' **   CUSIP:            I {opgAssetSource_optCusip on frmAccountHideTrans2_One}
' **   Show JournalNo:   J {chkShowJournalNo on frmAccountHideTrans2_One}
' **   Journal Type:     L {cmbJournalTypes on frmAccountHideTrans2_One}
' **   Name:             M {opgAssetSource_optName on frmAccountHideTrans2_One}
' **   Pick Hidden:      N {ledger_HIDDEN on frmAccountHideTrans2_One_Sub_Pick}
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

Private Const strSortOrig As String = "[transdate] DESC, [JournalType_Order], [journalno]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Top As Long, lngSortLbl_Width As Long
Private lngSortLine_Left As Long, lngSortLine_Width As Long

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

160         gstrFormQuerySpec = .Parent.Name

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
170         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180         lngSortLbl_Top = .Sort_lbl.Top
190         lngSortLbl_Width = .Sort_lbl.Width

200         .form_width_line.Left = 0&  ' ** Just get it out of the way.

210         lngForm_Width = .Width
220         lngJournalNo_Left = .journalno.Left
230         lngTransdate_Left = .transdate.Left
240         lngJournalType_Left = .journaltype.Left
250         lngShareface_Left = .shareface_0.Left
260         lngTotDesc_Left = .totdesc.Left
270         lngTotDesc_Width = .totdesc.Width
280         lngICash_Left = .icash_str.Left
290         lngICashLbl_Offset = (lngICash_Left - .icash_str_lbl.Left)
300         lngPCash_Left = .pcash_str.Left
310         lngCost_Left = .cost_str.Left
320         lngCurrID_Left = .curr_id.Left
330         lngCurrIDLbl_Offset = (lngCurrID_Left - .curr_id_lbl.Left)
340         lngLedgerHid_Left = .ledger_HIDDEN.Left
350         lngLedgerHidLbl_Offset = (lngLedgerHid_Left - .ledger_HIDDEN_lbl.Left)
360         lngLedgerHidLblLine_Offset = (lngLedgerHid_Left - .ledger_HIDDEN_lbl_line.Left)
370         lngJournalNo_Width = (lngTransdate_Left - lngJournalNo_Left)
380         lngCurrID_Width = ((lngLedgerHid_Left - lngLedgerHidLblLine_Offset) - lngCurrID_Left)

390         lngRecsCur = RecCnt  ' ** Function: Below.
400         If lngRecsCur > 0& Then
410           .FocusHolder.SetFocus
420         End If

430       Else
440         Cancel = -1
450       End If
460     End With

470     If Cancel = -1 Then
480       Beep
490       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
500     End If

EXITP:
510     Exit Sub

ERRH:
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

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

860       SortNow THIS_PROC  ' ** Procedure: Below.

870     End With

EXITP:
880     Exit Sub

ERRH:
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub

Private Sub Form_Current()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim lngHids As Long

1010    With Me
1020      lngHids = .hidden_tot
1030      If lngHids = 0& Then
1040        .ledger_HIDDEN.Locked = False
1050      Else
1060        Select Case .ledger_HIDDEN
            Case True
1070          .ledger_HIDDEN.Locked = False  ' ** They can unhide what they just hid.
1080        Case False
1090          .ledger_HIDDEN.Locked = True  ' ** There can only be 1 hidden entry at-a-time in this list.
1100        End Select
1110      End If
1120    End With

EXITP:
1130    Exit Sub

ERRH:
1140    Select Case ERR.Number
        Case Else
1150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1160    End Select
1170    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1210    intShiftDown = (Shift And acShiftMask) > 0
1220    intAltDown = (Shift And acAltMask) > 0
1230    intCtrlDown = (Shift And acCtrlMask) > 0

1240    intRetVal = KeyCode

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Hidden:           D {ledger_HIDDEN}
        ' **   Pick Hidden:      N {ledger_HIDDEN on frmAccountHideTrans2_One_Sub_Pick}
        ' **   Exit:             X {cmdClose on frmAccountHideTrans2_One}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Help:             H {cmdHelp on frmAccountHideTrans2_One}
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyEscape
1270        With Me
1280          intRetVal = 0
1290          DoCmd.SelectObject acForm, .Parent.Name, False
1300          .Parent.cmdClose.SetFocus
1310          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountHideTrans2_One.
1320        End With
1330      Case vbKeyUp
1340        intRetVal = 0
1350        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1360      Case vbKeyDown
1370        intRetVal = 0
1380        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1390      End Select
1400    End If

        ' ** Alt keys.
1410    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1420      Select Case intRetVal
          Case vbKeyD
1430        With Me
1440          intRetVal = 0
1450          .ledger_HIDDEN.SetFocus  ' ** Don't invoke.
1460        End With
1470      Case vbKeyN
1480        With Me
1490          intRetVal = 0
1500          lngRecsCur = .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_Pick.
1510          If lngRecsCur > 0& Then
1520            DoCmd.SelectObject acForm, .Parent.Name, False
1530            .Parent.frmAccountHideTrans2_One_Sub_Pick.SetFocus
1540            .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.ledger_HIDDEN.SetFocus  ' ** Don't invoke.
1550          Else
1560            Beep
1570          End If
1580        End With
1590      Case vbKeyX
1600        With Me
1610          intRetVal = 0
1620          DoCmd.SelectObject acForm, .Parent.Name, False
1630          .Parent.cmdClose.SetFocus
1640          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountHideTrans2_One.
1650        End With
1660      End Select
1670    End If

        ' ** Ctrl keys.
1680    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1690      Select Case intRetVal
          Case vbKeyH
1700        With Me
1710          intRetVal = 0
1720          DoCmd.SelectObject acForm, .Parent.Name, False
1730          .Parent.cmdHelp.SetFocus
1740          .Parent.cmdHelp_Click  ' ** Form Procedure: frmAccountHideTrans2_One.
1750        End With
1760      Case vbKeyS
1770        intRetVal = 0
1780        cmdSave_Click  ' ** Procedure: Below.
1790      Case vbKeyTab, vbKeyReturn
1800        With Me
1810          intRetVal = 0
1820          DoCmd.SelectObject acForm, .Parent.Name, False
1830          lngRecsCur = .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_Pick.
1840          If lngRecsCur > 0& Then
1850            .Parent.frmAccountHideTrans2_One_Sub_Pick.SetFocus
1860            .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.FocusHolder.SetFocus
1870          Else
1880            If .Parent.chkIncludeArchive.Enabled = True Then
1890              .Parent.chkIncludeArchive.SetFocus
1900            ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
1910              .Parent.chkIncludeCurrency.SetFocus
1920            Else
1930              .Parent.chkShowJournalNo.SetFocus
1940            End If
1950          End If
1960        End With
1970      Case vbKeyPageUp, vbKeyHome
1980        intRetVal = 0
1990        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2000      Case vbKeyPageDown, vbKeyEnd
2010        intRetVal = 0
2020        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2030      End Select
2040    End If

        ' ** Ctrl-Shift keys.
2050    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2060      Select Case intRetVal
          Case vbKeyF
2070        With Me
2080          intRetVal = 0
2090          DoCmd.SelectObject acForm, .Parent.Name, False
2100          .Parent.FocusHolder.SetFocus
2110        End With
2120      Case vbKeyTab, vbKeyReturn
2130        With Me
2140          intRetVal = 0
2150          DoCmd.SelectObject acForm, .Parent.Name, False
2160          .Parent.DateEnd.SetFocus
2170        End With
2180      End Select
2190    End If

EXITP:
2200    KeyCode = intRetVal
2210    Exit Sub

ERRH:
2220    intRetVal = 0
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

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
2490          .ledger_HIDDEN.SetFocus
2500        End With
2510      End Select
2520    End If

        ' ** Shift keys.
2530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2550        With Me
2560          intRetVal = 0
2570          If .CurrentRecord > 1 Then
2580            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2590            .ledger_HIDDEN.SetFocus
2600          Else
2610            DoCmd.SelectObject acForm, .Parent.Name, False
2620            .Parent.cmdClose.SetFocus
2630          End If
2640        End With
2650      End Select
2660    End If

EXITP:
2670    KeyCode = intRetVal
2680    Exit Sub

ERRH:
2690    intRetVal = 0
2700    Select Case ERR.Number
        Case Else
2710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2720    End Select
2730    Resume EXITP

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

Private Sub shareface_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

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

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

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

Private Sub icash_str_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_str_lbl_DblClick"

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

Private Sub pcash_str_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_str_lbl_DblClick"

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

Private Sub cost_str_lbl_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_str_lbl_DblClick"

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

Private Sub ledger_HIDDEN_AfterUpdate()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngJournalNo As Long
        Dim strSource As String

3810    With Me
3820      DoCmd.Hourglass True
3830      DoEvents
3840      lngJournalNo = .journalno
3850      strSource = .src

3860      Select Case .ledger_HIDDEN
          Case True
            ' ** There can be only 1 hidden entry at-a-time in this list.
3870        cmdSave_Click  ' ** Procedure: Above.
3880        DoEvents
3890        .Recalc
3900        Select Case strSource
            Case "L"
3910          Set dbs = CurrentDb
3920          With dbs
                ' ** Empty tblLedgerHidden_Staging2.
3930            Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_03")
3940            qdf.Execute
3950            Set qdf = Nothing
                ' ** Append qryAccountHideTrans2_One_Sub_Pick_01_01 (Ledger,
                ' ** by specified [jno]) to tblLedgerHidden_Staging2.
3960            Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_02_01")
3970            With qdf.Parameters
3980              ![jno] = lngJournalNo
3990            End With
4000            qdf.Execute
4010            Set qdf = Nothing
4020            .Close
4030          End With
4040          Set dbs = Nothing
4050        Case "A"
4060          Set dbs = CurrentDb
4070          With dbs
                ' ** Empty tblLedgerHidden_Staging2.
4080            Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_03")
4090            qdf.Execute
4100            Set qdf = Nothing
                ' ** Append qryAccountHideTrans2_One_Sub_Pick_01_02 (LedgerArchive,
                ' ** by specified [jno]) to tblLedgerHidden_Staging2.
4110            Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_02_02")
4120            With qdf.Parameters
4130              ![jno] = lngJournalNo
4140            End With
4150            qdf.Execute
4160            Set qdf = Nothing
4170            .Close
4180          End With
4190          Set dbs = Nothing
4200        End Select

4210        .Parent.pending_trans = 1
4220        DoEvents

4230        .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.Requery
4240        .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.LabelEnable True  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.
4250        .Parent.CreateReversingEntry lngJournalNo, strSource  ' ** Form Function: frmAccountHideTrans2_One.

4260      Case False
4270        cmdSave_Click  ' ** Procedure: Above.
4280        DoEvents
4290        .Recalc
4300        Set dbs = CurrentDb
4310        With dbs
              ' ** Empty tblLedgerHidden_Staging2.
4320          Set qdf = .QueryDefs("qryAccountHideTrans2_One_Sub_Pick_03")
4330          qdf.Execute
4340          Set qdf = Nothing
4350          .Close
4360        End With
4370        Set dbs = Nothing
4380        .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.Requery
4390        .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.LabelEnable False  ' ** Form Procedure: frmAccountHideTrans2_One_Sub_Pick.

4400        .Parent.cmdOK.Enabled = False
4410        .Parent.cmdCancel.Enabled = False

4420        .Parent.shareface_subtot_box1.Visible = False
4430        .Parent.shareface_subtot_box2.Visible = False
4440        .Parent.icash_subtot_box1.Visible = False
4450        .Parent.icash_subtot_box2.Visible = False
4460        .Parent.pcash_subtot_box1.Visible = False
4470        .Parent.pcash_subtot_box2.Visible = False
4480        .Parent.cost_subtot_box1.Visible = False
4490        .Parent.cost_subtot_box2.Visible = False
4500        .Parent.curr_id_box1.Visible = False
4510        .Parent.curr_id_box2.Visible = False

4520        .Parent.shareface_subtot.Visible = True
4530        .Parent.shareface_divint_subtot.Visible = False
4540        .Parent.shareface_subtot.Enabled = False
4550        .Parent.shareface_subtot.Locked = False
4560        .Parent.shareface_subtot = 0#
4570        .Parent.shareface_subtot.BorderColor = WIN_CLR_DISR
4580        .Parent.shareface_subtot.BackStyle = acBackStyleTransparent
4590        .Parent.shareface_subtot_lbl.ForeColor = WIN_CLR_DISF
4600        .Parent.shareface_subtot_lbl_dim_hi.Visible = True

4610        .Parent.icash_subtot.Enabled = False
4620        .Parent.icash_subtot.Locked = False
4630        .Parent.icash_subtot = 0@
4640        .Parent.icash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
4650        .Parent.icash_subtot.BorderColor = WIN_CLR_DISR
4660        .Parent.icash_subtot.BackStyle = acBackStyleTransparent
4670        .Parent.icash_subtot_lbl.ForeColor = WIN_CLR_DISF
4680        .Parent.icash_subtot_lbl_dim_hi.Visible = True

4690        .Parent.pcash_subtot.Enabled = False
4700        .Parent.pcash_subtot.Locked = False
4710        .Parent.pcash_subtot = 0@
4720        .Parent.pcash_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
4730        .Parent.pcash_subtot.BorderColor = WIN_CLR_DISR
4740        .Parent.pcash_subtot.BackStyle = acBackStyleTransparent
4750        .Parent.pcash_subtot_lbl.ForeColor = WIN_CLR_DISF
4760        .Parent.pcash_subtot_lbl_dim_hi.Visible = True

4770        .Parent.cost_subtot.Enabled = False
4780        .Parent.cost_subtot.Locked = False
4790        .Parent.cost_subtot = 0@
4800        .Parent.cost_subtot.Format = "$#,##0.00;($#,##0.00);$0.00;$0.00"
4810        .Parent.cost_subtot.BorderColor = WIN_CLR_DISR
4820        .Parent.cost_subtot.BackStyle = acBackStyleTransparent
4830        .Parent.cost_subtot_lbl.ForeColor = WIN_CLR_DISF
4840        .Parent.cost_subtot_lbl_dim_hi.Visible = True

4850        .Parent.subtot_lbl_dim.Visible = True
4860        .Parent.subtot_lbl_dim_hi.Visible = True
4870        .Parent.subtot_lbl.Visible = False

4880        If gblnHasForExThis = True Then
4890          .Parent.curr_id = Null
4900          .Parent.curr_id.BorderColor = WIN_CLR_DISR
4910          .Parent.curr_id.BackStyle = acBackStyleTransparent
4920          .Parent.curr_id.Visible = True
4930          .Parent.curr_id_lbl.ForeColor = WIN_CLR_DISR
4940          .Parent.curr_id_lbl.Visible = True
4950          .Parent.curr_id_lbl_dim_hi.Visible = True
4960        Else
4970          Select Case .Parent.chkIncludeCurrency
              Case True
4980            .Parent.curr_id = Null
4990            .Parent.curr_id.BorderColor = WIN_CLR_DISR
5000            .Parent.curr_id.BackStyle = acBackStyleTransparent
5010            .Parent.curr_id.Visible = True
5020            .Parent.curr_id_lbl.ForeColor = WIN_CLR_DISR
5030            .Parent.curr_id_lbl.Visible = True
5040            .Parent.curr_id_lbl_dim_hi.Visible = True
5050          Case False
5060            .Parent.curr_id.Visible = False
5070            .Parent.curr_id_lbl.Visible = False
5080            .Parent.curr_id_lbl_dim_hi.Visible = False
5090          End Select
5100        End If

5110        .Parent.pending_trans.Enabled = False
5120        .Parent.pending_trans.Locked = False
5130        .Parent.pending_trans.BorderColor = WIN_CLR_DISR
5140        .Parent.pending_trans.BackStyle = acBackStyleTransparent

5150        .Parent.pending_trans = 0

5160      End Select
5170      DoCmd.Hourglass False
5180    End With

EXITP:
5190    Set qdf = Nothing
5200    Set dbs = Nothing
5210    Exit Sub

ERRH:
5220    DoCmd.Hourglass False
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
5310    intShiftDown = (Shift And acShiftMask) > 0
5320    intAltDown = (Shift And acAltMask) > 0
5330    intCtrlDown = (Shift And acCtrlMask) > 0

5340    intRetVal = KeyCode

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5370        With Me
5380          intRetVal = 0
5390          lngRecsCur = RecCnt  ' ** Function: Below.
5400          If .CurrentRecord < lngRecsCur Then
5410            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5420            .ledger_HIDDEN.SetFocus
5430          Else
5440            DoCmd.SelectObject acForm, .Parent.Name, False
5450            lngRecsCur = .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_One_Sub_Pick.
5460            If lngRecsCur > 0& Then
5470              .Parent.frmAccountHideTrans2_One_Sub_Pick.SetFocus
5480              .Parent.frmAccountHideTrans2_One_Sub_Pick.Form.FocusHolder.SetFocus
5490            Else
5500              .Parent.DateStart.SetFocus
5510            End If
5520          End If
5530        End With
5540      End Select
5550    End If

        ' ** Shift keys.
5560    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5570      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5580        With Me
5590          intRetVal = 0
5600          If .CurrentRecord > 1 Then
5610            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5620            .ledger_HIDDEN.SetFocus
5630          Else
5640            DoCmd.SelectObject acForm, .Parent.Name, False
5650            .Parent.cmdClose.SetFocus
5660          End If
5670        End With
5680      End Select
5690    End If

EXITP:
5700    KeyCode = intRetVal
5710    Exit Sub

ERRH:
5720    intRetVal = 0
5730    Select Case ERR.Number
        Case Else
5740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5750    End Select
5760    Resume EXITP

End Sub

Public Sub ShowJournalNo_Sub(blnShow As Boolean)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo_Sub"

        Dim blnCurrency As Boolean, blnSortHere As Boolean, blnResort As Boolean

5810    With Me

5820      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
5830        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5840      End If

5850      blnCurrency = .Parent.chkIncludeCurrency
5860      blnSortHere = False

5870      Select Case blnShow
          Case True

5880        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
5890        .totdesc.Left = lngTotDesc_Left
5900        .totdesc_lbl_dim_hi.Left = (.totdesc.Left + lngTpp)
5910        .totdesc_lbl.Left = lngTotDesc_Left
5920        .totdesc_lbl_line.Left = lngTotDesc_Left
5930        .totdesc_lbl_line_dim_hi.Left = (.totdesc_lbl_line.Left + lngTpp)
5940        Select Case blnCurrency
            Case True
5950          .totdesc.Width = lngTotDesc_Width
5960          .totdesc_lbl.Width = .totdesc.Width
5970          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
5980        Case False
5990          .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)
6000          .totdesc_lbl.Width = .totdesc.Width
6010          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6020        End Select
6030        .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
6040        .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
6050        If blnSortHere = True Then
6060          .Sort_line.Left = .totdesc_lbl.Left
6070          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
6080          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
6090        End If
6100        blnSortHere = False
6110        If .Sort_line.Left = .journaltype_lbl.Left Then blnSortHere = True
6120        .journaltype.Left = lngJournalType_Left
6130        .journaltype_lbl.Left = lngJournalType_Left
6140        .journaltype_lbl_dim_hi.Left = (.journaltype_lbl.Left + lngTpp)
6150        .journaltype_lbl_line.Left = lngJournalType_Left
6160        .journaltype_lbl_line_dim_hi.Left = (.journaltype_lbl_line.Left + lngTpp)
6170        If blnSortHere = True Then
6180          .Sort_line.Left = .journaltype_lbl.Left
6190          .Sort_lbl.Left = ((.journaltype_lbl.Left + .journaltype_lbl.Width) - lngSortLbl_Width)
6200        End If
6210        blnSortHere = False
6220        If .Sort_line.Left = .transdate_lbl.Left Then blnSortHere = True
6230        .transdate.Left = lngTransdate_Left
6240        .transdate_lbl.Left = lngTransdate_Left
6250        .transdate_lbl_dim_hi.Left = (.transdate_lbl.Left + lngTpp)
6260        .transdate_lbl_line.Left = lngTransdate_Left
6270        .transdate_lbl_line_dim_hi.Left = (.transdate_lbl_line.Left + lngTpp)
6280        If blnSortHere = True Then
6290          .Sort_line.Left = .transdate_lbl.Left
6300          .Sort_lbl.Left = ((.transdate_lbl.Left + .transdate_lbl.Width) - lngSortLbl_Width)
6310        End If
6320        blnSortHere = False
6330        .journalno.Visible = True
6340        .journalno_lbl.Visible = True
6350        .journalno_lbl2.Visible = True
6360        .journalno_lbl_line.Visible = True

6370      Case False

6380        blnResort = False
6390        If .journalno_Sort_line1.Visible = True Then blnSortHere = True
6400        .journalno.Visible = False
6410        .journalno_lbl.Visible = False
6420        .journalno_lbl2.Visible = False
6430        .journalno_lbl_line.Visible = False
6440        .journalno_Sort_line1.Visible = False
6450        .journalno_Sort_line2.Visible = False
6460        .journalno_lbl_dim_hi.Visible = False
6470        .journalno_lbl2_dim_hi.Visible = False
6480        .journalno_lbl_line_dim_hi.Visible = False
6490        If blnSortHere = True Then
6500          blnResort = True
6510          .Sort_lbl.Visible = False
6520        End If
6530        blnSortHere = False
6540        If .Sort_line.Left = .transdate_lbl.Left Then blnSortHere = True
6550        .transdate.Left = (lngTransdate_Left - lngJournalNo_Width)
6560        .transdate_lbl.Left = .transdate.Left
6570        .transdate_lbl_dim_hi.Left = (.transdate_lbl.Left + lngTpp)
6580        .transdate_lbl_line.Left = .transdate_lbl.Left
6590        .transdate_lbl_line_dim_hi.Left = (.transdate_lbl_line.Left + lngTpp)
6600        If blnSortHere = True Then
6610          .Sort_line.Left = .transdate_lbl.Left
6620          .Sort_lbl.Left = ((.transdate_lbl.Left + .transdate_lbl.Width) - lngSortLbl_Width)
6630        End If
6640        blnSortHere = False
6650        If .Sort_line.Left = .journaltype_lbl.Left Then blnSortHere = True
6660        .journaltype.Left = (lngJournalType_Left - lngJournalNo_Width)
6670        .journaltype_lbl.Left = .journaltype.Left
6680        .journaltype_lbl_dim_hi.Left = (.journaltype_lbl.Left + lngTpp)
6690        .journaltype_lbl_line.Left = .journaltype_lbl.Left
6700        .journaltype_lbl_line_dim_hi.Left = (.journaltype_lbl_line.Left + lngTpp)
6710        If blnSortHere = True Then
6720          .Sort_line.Left = .journaltype_lbl.Left
6730          .Sort_lbl.Left = ((.journaltype_lbl.Left + .journaltype_lbl.Width) - lngSortLbl_Width)
6740        End If
6750        blnSortHere = False
6760        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
6770        .totdesc.Left = (lngTotDesc_Left - lngJournalNo_Width)
6780        .totdesc_lbl.Left = .totdesc.Left
6790        .totdesc_lbl_dim_hi.Left = (.totdesc_lbl.Left + lngTpp)
6800        .totdesc_lbl_line.Left = .totdesc_lbl.Left
6810        .totdesc_lbl_line_dim_hi.Left = (.totdesc_lbl_line.Left + lngTpp)
6820        Select Case blnCurrency
            Case True
6830          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)
6840          .totdesc_lbl.Width = .totdesc.Width
6850          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6860        Case False
6870          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width + lngCurrID_Width)
6880          .totdesc_lbl.Width = .totdesc.Width
6890          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6900        End Select
6910        .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
6920        .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
6930        If blnSortHere = True Then
6940          .Sort_line.Left = .totdesc_lbl.Left
6950          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
6960          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
6970        End If
6980        blnSortHere = False

6990        If blnResort = True Then
7000          SortNow "Form_Load"  ' ** Procedure: Below.
7010        End If

7020      End Select

7030    End With

EXITP:
7040    Exit Sub

ERRH:
7050    Select Case ERR.Number
        Case Else
7060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7070    End Select
7080    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnInclude As Boolean)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnJournalNo As Boolean, blnSortHere As Boolean, blnResort As Boolean

7110    With Me

7120      DoCmd.Hourglass True
7130      DoEvents

7140      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
7150        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7160      End If

7170      blnJournalNo = .Parent.chkShowJournalNo
7180      blnSortHere = False

7190      Select Case blnInclude
          Case True

7200        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
7210        Select Case blnJournalNo
            Case True
7220          .totdesc.Width = lngTotDesc_Width
7230          .totdesc_lbl.Width = .totdesc.Width
7240          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7250        Case False
7260          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)
7270          .totdesc_lbl.Width = .totdesc.Width
7280          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7290        End Select
7300        .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
7310        .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
7320        If blnSortHere = True Then
7330          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
7340          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
7350        End If
7360        blnSortHere = False
7370        If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
7380        .shareface_0.Left = lngShareface_Left
7390        .shareface_1.Left = .shareface_0.Left
7400        .shareface_2.Left = .shareface_0.Left
7410        .shareface_3.Left = .shareface_0.Left
7420        .shareface_4.Left = .shareface_0.Left
7430        .shareface_box.Left = .shareface_0.Left
7440        .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
7450        .shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
7460        .shareface_lbl_line.Left = .shareface_lbl.Left
7470        .shareface_lbl_line_dim_hi.Left = .shareface_lbl_line.Left
7480        If blnSortHere = True Then
7490          .Sort_line.Left = .shareface_lbl.Left
7500          .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
7510        End If
7520        blnSortHere = False
7530        If .Sort_line.Left = .icash_str_lbl.Left Then blnSortHere = True
7540        .ICash.Left = lngICash_Left
7550        .icash_str.Left = .ICash.Left
7560        .icash_forex.Left = (.icash_str.Left - lngTpp)
7570        .icash_box.Left = .ICash.Left
7580        .icash_str_lbl.Left = (.icash_str.Left - lngICashLbl_Offset)
7590        .icash_str_lbl_dim_hi.Left = (.icash_str_lbl.Left + lngTpp)
7600        .icash_str_lbl_line.Left = .icash_str_lbl.Left
7610        .icash_str_lbl_line_dim_hi.Left = (.icash_str_lbl_line.Left + lngTpp)
7620        If blnSortHere = True Then
7630          .Sort_line.Left = .icash_str_lbl.Left
7640          .Sort_lbl.Left = ((.icash_str_lbl.Left + .icash_str_lbl.Width) - lngSortLbl_Width)
7650        End If
7660        blnSortHere = False
7670        If .Sort_line.Left = .pcash_str_lbl.Left Then blnSortHere = True
7680        .PCash.Left = lngPCash_Left
7690        .pcash_str.Left = .PCash.Left
7700        .pcash_forex.Left = (.pcash_str.Left - lngTpp)
7710        .pcash_box.Left = .PCash.Left
7720        .pcash_str_lbl.Left = (.pcash_str.Left - lngICashLbl_Offset)
7730        .pcash_str_lbl_dim_hi.Left = (.pcash_str_lbl.Left + lngTpp)
7740        .pcash_str_lbl_line.Left = .pcash_str_lbl.Left
7750        .pcash_str_lbl_line_dim_hi.Left = (.pcash_str_lbl_line.Left + lngTpp)
7760        If blnSortHere = True Then
7770          .Sort_line.Left = .pcash_str_lbl.Left
7780          .Sort_lbl.Left = ((.pcash_str_lbl.Left + .pcash_str_lbl.Width) - lngSortLbl_Width)
7790        End If
7800        blnSortHere = False
7810        If .Sort_line.Left = .cost_str_lbl.Left Then blnSortHere = True
7820        .Cost.Left = lngCost_Left
7830        .cost_str.Left = lngCost_Left
7840        .cost_forex.Left = (.cost_str.Left - lngTpp)
7850        .cost_box.Left = .Cost.Left
7860        .cost_str_lbl.Left = (.cost_str.Left - lngICashLbl_Offset)
7870        .cost_str_lbl_dim_hi.Left = (.cost_str_lbl.Left + lngTpp)
7880        .cost_str_lbl_line.Left = .cost_str_lbl.Left
7890        .cost_str_lbl_line_dim_hi.Left = (.cost_str_lbl_line.Left + lngTpp)
7900        If blnSortHere = True Then
7910          .Sort_line.Left = .cost_str_lbl.Left
7920          .Sort_lbl.Left = ((.cost_str_lbl.Left + .cost_str_lbl.Width) - lngSortLbl_Width)
7930        End If
7940        blnSortHere = False
7950        .curr_id.Left = lngCurrID_Left
7960        .curr_id_forex.Left = (.curr_id.Left - lngTpp)
7970        .curr_id_box.Left = .curr_id.Left
7980        .curr_id_lbl.Left = (lngCurrID_Left - lngCurrIDLbl_Offset)
7990        .curr_id_lbl_dim_hi.Left = (.curr_id_lbl.Left + lngTpp)
8000        .curr_id_lbl_line.Left = lngCurrID_Left
8010        .curr_id_lbl_line_dim_hi.Left = (.curr_id_lbl_line.Left + lngTpp)
8020        .curr_id.Visible = True
8030        .curr_id_forex.Visible = True
8040        .curr_id_box.Visible = True
8050        .curr_id_lbl.Visible = True
8060        .curr_id_lbl_line.Visible = True

8070      Case False

8080        blnResort = False
8090        Select Case blnJournalNo
            Case True
8100          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
8110          .curr_id.Visible = False
8120          .curr_id_forex.Visible = False
8130          .curr_id_box.Visible = False
8140          .curr_id_lbl.Visible = False
8150          .curr_id_lbl_line.Visible = False
8160          .curr_id.Left = 0&
8170          .curr_id_forex.Left = 0&
8180          .curr_id_box.Left = 0&
8190          .curr_id_lbl.Left = 0&
8200          .curr_id_lbl_line.Left = 0&
8210          .curr_id_lbl_dim_hi.Visible = False
8220          .curr_id_lbl_line_dim_hi.Visible = False
8230          .curr_id_lbl_dim_hi.Left = 0&
8240          .curr_id_lbl_line_dim_hi.Left = 0&
8250          If blnSortHere = True Then
8260            blnResort = True
8270            .Sort_line.Visible = False
8280            .Sort_lbl.Visible = False
8290          End If
8300          blnSortHere = False
8310          If .Sort_line.Left = .cost_str_lbl.Left Then blnSortHere = True
8320          .Cost.Left = (lngCost_Left + lngCurrID_Width)
8330          .cost_str.Left = .Cost.Left
8340          .cost_forex.Left = (.cost_str.Left - lngTpp)
8350          .cost_box.Left = .Cost.Left
8360          .cost_str_lbl.Left = (.cost_str.Left - lngICashLbl_Offset)
8370          .cost_str_lbl_dim_hi.Left = (.cost_str_lbl.Left + lngTpp)
8380          .cost_str_lbl_line.Left = .cost_str_lbl.Left
8390          .cost_str_lbl_line_dim_hi.Left = (.cost_str_lbl_line.Left + lngTpp)
8400          If blnSortHere = True Then
8410            .Sort_line.Left = .cost_str_lbl.Left
8420            .Sort_lbl.Left = ((.cost_str_lbl.Left + .cost_str_lbl.Width) - lngSortLbl_Width)
8430          End If
8440          blnSortHere = False
8450          If .Sort_line.Left = .pcash_str_lbl.Left Then blnSortHere = True
8460          .PCash.Left = (lngPCash_Left + lngCurrID_Width)
8470          .pcash_str.Left = .PCash.Left
8480          .pcash_forex.Left = (.pcash_str.Left - lngTpp)
8490          .pcash_box.Left = .PCash.Left
8500          .pcash_str_lbl.Left = (.pcash_str.Left - lngICashLbl_Offset)
8510          .pcash_str_lbl_dim_hi.Left = (.pcash_str_lbl.Left + lngTpp)
8520          .pcash_str_lbl_line.Left = .pcash_str_lbl.Left
8530          .pcash_str_lbl_line_dim_hi.Left = (.pcash_str_lbl_line.Left + lngTpp)
8540          If blnSortHere = True Then
8550            .Sort_line.Left = .pcash_str_lbl.Left
8560            .Sort_lbl.Left = ((.pcash_str_lbl.Left + .pcash_str_lbl.Width) - lngSortLbl_Width)
8570          End If
8580          blnSortHere = False
8590          If .Sort_line.Left = .icash_str_lbl.Left Then blnSortHere = True
8600          .ICash.Left = (lngICash_Left + lngCurrID_Width)
8610          .icash_str.Left = .ICash.Left
8620          .icash_forex.Left = (.icash_str.Left - lngTpp)
8630          .icash_box.Left = .ICash.Left
8640          .icash_str_lbl.Left = (.icash_str.Left - lngICashLbl_Offset)
8650          .icash_str_lbl_dim_hi.Left = (.icash_str_lbl.Left + lngTpp)
8660          .icash_str_lbl_line.Left = .icash_str_lbl.Left
8670          .icash_str_lbl_line_dim_hi.Left = (.icash_str_lbl_line.Left + lngTpp)
8680          If blnSortHere = True Then
8690            .Sort_line.Left = .icash_str_lbl.Left
8700            .Sort_lbl.Left = ((.icash_str_lbl.Left + .icash_str_lbl.Width) - lngSortLbl_Width)
8710          End If
8720          blnSortHere = False
8730          If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
8740          .shareface_0.Left = (lngShareface_Left + lngCurrID_Width)
8750          .shareface_1.Left = .shareface_0.Left
8760          .shareface_2.Left = .shareface_0.Left
8770          .shareface_3.Left = .shareface_0.Left
8780          .shareface_4.Left = .shareface_0.Left
8790          .shareface_box.Left = .shareface_0.Left
8800          .shareface_lbl.Left = (.shareface_0.Left - lngICashLbl_Offset)
8810          .shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
8820          .shareface_lbl_line.Left = .shareface_lbl.Left
8830          .shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
8840          If blnSortHere = True Then
8850            .Sort_line.Left = .shareface_lbl.Left
8860            .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
8870          End If
8880          blnSortHere = False
8890          If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
8900          .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)
8910          .totdesc_lbl.Width = .totdesc.Width
8920          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
8930          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
8940          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
8950          If blnSortHere = True Then
8960            .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
8970            .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
8980          End If
8990          blnSortHere = False
9000        Case False
9010          If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
9020          .curr_id.Visible = False
9030          .curr_id_forex.Visible = False
9040          .curr_id_box.Visible = False
9050          .curr_id_lbl.Visible = False
9060          .curr_id_lbl_line.Visible = False
9070          .curr_id.Left = 0&
9080          .curr_id_forex.Left = 0&
9090          .curr_id_box.Left = 0&
9100          .curr_id_lbl.Left = 0&
9110          .curr_id_lbl_line.Left = 0&
9120          .curr_id_lbl_dim_hi.Visible = False
9130          .curr_id_lbl_line_dim_hi.Visible = False
9140          .curr_id_lbl_dim_hi.Left = 0&
9150          .curr_id_lbl_line_dim_hi.Left = 0&
9160          If blnSortHere = True Then
9170            blnResort = True
9180            .Sort_line.Visible = False
9190            .Sort_lbl.Visible = False
9200          End If
9210          blnSortHere = False
9220          If .Sort_line.Left = .cost_str_lbl.Left Then blnSortHere = True
9230          .Cost.Left = (lngCost_Left + lngCurrID_Width)
9240          .cost_str.Left = .Cost.Left
9250          .cost_forex.Left = (.cost_str.Left - lngTpp)
9260          .cost_box.Left = .Cost.Left
9270          .cost_str_lbl.Left = (.cost_str.Left - lngICashLbl_Offset)
9280          .cost_str_lbl_dim_hi.Left = (.cost_str_lbl.Left + lngTpp)
9290          .cost_str_lbl_line.Left = .cost_str_lbl.Left
9300          .cost_str_lbl_line_dim_hi.Left = (.cost_str_lbl_line.Left + lngTpp)
9310          If blnSortHere = True Then
9320            .Sort_line.Left = .cost_str_lbl.Left
9330            .Sort_lbl.Left = ((.cost_str_lbl.Left + .cost_str_lbl.Width) - lngSortLbl_Width)
9340          End If
9350          blnSortHere = False
9360          If .Sort_line.Left = .pcash_str_lbl.Left Then blnSortHere = True
9370          .PCash.Left = (lngPCash_Left + lngCurrID_Width)
9380          .pcash_str.Left = .PCash.Left
9390          .pcash_forex.Left = (.pcash_str.Left - lngTpp)
9400          .pcash_box.Left = .PCash.Left
9410          .pcash_str_lbl.Left = (.pcash_str.Left - lngICashLbl_Offset)
9420          .pcash_str_lbl_dim_hi.Left = (.pcash_str_lbl.Left + lngTpp)
9430          .pcash_str_lbl_line.Left = .pcash_str_lbl.Left
9440          .pcash_str_lbl_line_dim_hi.Left = (.pcash_str_lbl_line.Left + lngTpp)
9450          If blnSortHere = True Then
9460            .Sort_line.Left = .pcash_str_lbl.Left
9470            .Sort_lbl.Left = ((.pcash_str_lbl.Left + .pcash_str_lbl.Width) - lngSortLbl_Width)
9480          End If
9490          blnSortHere = False
9500          If .Sort_line.Left = .icash_str_lbl.Left Then blnSortHere = True
9510          .ICash.Left = (lngICash_Left + lngCurrID_Width)
9520          .icash_str.Left = .ICash.Left
9530          .icash_forex.Left = (.icash_str.Left - lngTpp)
9540          .icash_box.Left = .ICash.Left
9550          .icash_str_lbl.Left = (.icash_str.Left - lngICashLbl_Offset)
9560          .icash_str_lbl_dim_hi.Left = (.icash_str_lbl.Left + lngTpp)
9570          .icash_str_lbl_line.Left = .icash_str_lbl.Left
9580          .icash_str_lbl_line_dim_hi.Left = (.icash_str_lbl_line.Left + lngTpp)
9590          If blnSortHere = True Then
9600            .Sort_line.Left = .icash_str_lbl.Left
9610            .Sort_lbl.Left = ((.icash_str_lbl.Left + .icash_str_lbl.Width) - lngSortLbl_Width)
9620          End If
9630          blnSortHere = False
9640          If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
9650          .shareface_0.Left = (lngShareface_Left + lngCurrID_Width)
9660          .shareface_1.Left = .shareface_0.Left
9670          .shareface_2.Left = .shareface_0.Left
9680          .shareface_3.Left = .shareface_0.Left
9690          .shareface_4.Left = .shareface_0.Left
9700          .shareface_box.Left = .shareface_0.Left
9710          .shareface_lbl.Left = (shareface_0.Left - lngICashLbl_Offset)
9720          .shareface_lbl_dim_hi.Left = (.shareface_lbl.Left + lngTpp)
9730          .shareface_lbl_line.Left = .shareface_lbl.Left
9740          .shareface_lbl_line_dim_hi.Left = (.shareface_lbl_line.Left + lngTpp)
9750          If blnSortHere = True Then
9760            .Sort_line.Left = .shareface_lbl.Left
9770            .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
9780          End If
9790          blnSortHere = False
9800          If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
9810          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width + lngCurrID_Width)
9820          .totdesc_lbl.Width = .totdesc.Width
9830          .totdesc_lbl_dim_hi.Width = .totdesc_lbl.Width
9840          .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
9850          .totdesc_lbl_line_dim_hi.Width = .totdesc_lbl_line.Width
9860          If blnSortHere = True Then
9870            .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
9880            .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
9890          End If
9900          blnSortHere = False
9910        End Select

9920        If blnResort = True Then
9930          SortNow "Form_Load"  ' ** Procedure: Below.
9940        End If

9950      End Select

9960    End With

EXITP:
9970    Exit Sub

ERRH:
9980    DoCmd.Hourglass False
9990    Select Case ERR.Number
        Case Else
10000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10010   End Select
10020   Resume EXITP

End Sub

Public Function RecCnt() As Long

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

10110   lngRetVal = 0&

10120   With Me
10130     Set rst = .RecordsetClone
10140     With rst
10150       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
10160       Else
10170         .MoveLast
10180         lngRetVal = .RecordCount
10190       End If
10200       .Close
10210     End With
10220   End With

EXITP:
10230   Set rst = Nothing
10240   RecCnt = lngRetVal
10250   Exit Function

ERRH:
10260   lngRetVal = 0&
10270   Select Case ERR.Number
        Case Else
10280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10290   End Select
10300   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

10410   Select Case intWhere
        Case acCmdRecordsGoToFirst
10420     DoCmd.RunCommand acCmdRecordsGoToFirst
10430   Case acCmdRecordsGoToLast
10440     DoCmd.RunCommand acCmdRecordsGoToLast
10450   Case acCmdRecordsGoToPrevious
10460     DoCmd.RunCommand acCmdRecordsGoToPrevious
10470   Case acCmdRecordsGoToNext
10480     DoCmd.RunCommand acCmdRecordsGoToNext
10490   Case Else
10500     If IsMissing(varID) = False Then
10510       With Me
10520         Set rst = .RecordsetClone
10530         With rst
10540           .FindFirst "[journalno] = " & CStr(varID)
10550           If .NoMatch = False Then
10560             Me.Bookmark = .Bookmark
10570           End If
10580           .Close
10590         End With
10600       End With
10610     End If
10620   End Select

EXITP:
10630   Set rst = Nothing
10640   Exit Sub

ERRH:
10650   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
10660   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
10670   Case Else
10680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10690   End Select
10700   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [JournalType_Order], [journalno]"

10810   With Me
10820     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
10830       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
10840     End If
10850     .Controls(strSortLbl).Visible = False
10860     .Controls(strSortLbl).Top = lngSortLbl_Top
10870     .Controls(strSortLine).Visible = False
10880     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
10890     .journalno_Sort_line1.Visible = False
10900     .journalno_Sort_line2.Visible = False
10910     DoEvents
10920     If strProc = "Form_Load" Then
10930       strCalled = "transdate"
10940       strSortNow = strSortOrig
10950       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
10960       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
10970       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
10980       .Controls(strSortLbl).Left = lngSortLbl_Left
10990       .Controls(strSortLine).Left = lngSortLine_Left
11000       .Controls(strSortLine).Width = lngSortLine_Width
11010       .Controls(strSortLbl).Caption = strArwDn
11020       .Controls(strSortLbl).ForeColor = CLR_DKRED
11030       DoEvents
11040     Else
11050       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: revcode_ID_lbl_DblClick
11060       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
11070       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
11080       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
11090       .Controls(strSortLbl).Left = lngSortLbl_Left
11100       .Controls(strSortLine).Left = lngSortLine_Left
11110       .Controls(strSortLine).Width = lngSortLine_Width
11120       If strCalled = "transdate" Then
11130         strSortAsc = strSortOrig
11140         If strSortNow = strSortAsc Then
11150           strSortNow = "[transdate], [JournalType_Order], [journalno]"
11160           .Controls(strSortLbl).Caption = strArwUp
11170           .Controls(strSortLbl).ForeColor = CLR_DKBLU
11180         Else
11190           strSortNow = strSortAsc
11200           .Controls(strSortLbl).Caption = strArwDn
11210           .Controls(strSortLbl).ForeColor = CLR_DKRED
11220         End If
11230         DoEvents
11240       Else
11250         .Controls(strSortLbl).Caption = strArwUp
11260         .Controls(strSortLbl).ForeColor = CLR_DKBLU
11270         Select Case strCalled
              Case "journalno"
11280           strSortAsc = "[journalno]"
11290           lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
11300           .Controls(strSortLbl).Top = (lngSortLbl_Top - (2& * lngTpp))
11310           .Controls(strSortLbl).Left = lngSortLbl_Left
11320         Case "journaltype"
11330           strSortAsc = "[journaltype], [transdate], [journalno]"
11340         Case "shareface"
11350           strSortAsc = "[shareface]" & strStdAsc
11360         Case "totdesc"
11370           strSortAsc = "[totdesc]" & strStdAsc
11380         Case "icash", "icash_str"
11390           strSortAsc = "[icash]" & strStdAsc
11400         Case "pcash", "pcash_str"
11410           strSortAsc = "[pcash]" & strStdAsc
11420         Case "cost", "cost_str"
11430           strSortAsc = "[cost]" & strStdAsc
11440         Case "curr_id"
11450           strSortAsc = "[curr_code]" & strStdAsc
11460           lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
11470           lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
11480           lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
11490           .Controls(strSortLbl).Left = lngSortLbl_Left
11500           .Controls(strSortLine).Left = lngSortLine_Left
11510           .Controls(strSortLine).Width = lngSortLine_Width
11520         Case "ledger_HIDDEN"
11530           strSortAsc = "[ledger_HIDDEN]" & strStdAsc
11540         End Select
11550         If strSortNow = strSortAsc Then
11560           intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
11570           Select Case intCnt
                Case 1
11580             strTmp01 = strSortAsc & " DESC"
11590           Case 2
11600             intPos01 = InStr(strSortAsc, ",")
11610             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11620             strTmp01 = strTmp01 & Mid(strTmp01, intPos01) & " DESC"
11630           Case 3
11640             intPos01 = InStr(strSortAsc, ",")
11650             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11660             strTmp02 = Mid(strSortAsc, intPos01)
11670             intPos01 = InStr(2, strTmp02, ",")
11680             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11690             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
11700           Case 4
11710             intPos01 = InStr(strSortAsc, ",")
11720             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11730             strTmp02 = Mid(strSortAsc, intPos01)
11740             intPos01 = InStr(2, strTmp02, ",")
11750             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11760             strTmp02 = Mid(strTmp02, intPos01)
11770             intPos01 = InStr(2, strTmp02, ",")
11780             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11790             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
11800           End Select
11810           strSortNow = strTmp01
11820           .Controls(strSortLbl).Caption = strArwDn
11830           .Controls(strSortLbl).ForeColor = CLR_DKRED
11840         Else
11850           strSortNow = strSortAsc
11860         End If
11870         DoEvents
11880       End If
11890     End If
11900     If strCalled = "journalno" Then
11910       .journalno_Sort_line1.Visible = True
11920       .journalno_Sort_line2.Visible = True
11930     Else
11940       .Controls(strSortLine).Visible = True
11950     End If
11960     .Controls(strSortLbl).Visible = True
11970     DoEvents
11980     .OrderBy = strSortNow
11990     .OrderByOn = True
12000   End With

EXITP:
12010   Exit Sub

ERRH:
12020   Select Case ERR.Number
        Case Else
12030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12040   End Select
12050   Resume EXITP

End Sub
