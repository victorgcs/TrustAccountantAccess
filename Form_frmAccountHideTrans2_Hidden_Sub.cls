VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans2_Hidden_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans2_Hidden_Sub"

'VGC 09/07/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Check Again:      A {cmdCheckAgain on frmAccountHideTrans2_Hidden}
' **   Close:            C {cmdClose on frmAccountHideTrans2_Hidden}
' **   Show JournalNo:   J {chkShowJournalNo on frmAccountHideTrans2_Hidden}
' **   Print Report:     P {cmdPrintReport on frmAccountHideTrans2_Hidden}
' **   Exit:             X {cmdClose on frmAccountHideTrans2_Hidden}

' ***********************************************
' ** Easter Egg (Secret Key):
' ** ledger_HIDDEN_lbl_MouseUp
' ***********************************************

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[ledghid_grpnum], [ledghid_ord]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngTotDesc_Left As Long, lngTotDesc_Width As Long, lngShareface_Left As Long, lngJournalType_Left As Long
Private lngTransdate_Left As Long, lngJournalNo_Width As Long, lngCurrID_Width As Long
Private lngICash_Left As Long, lngPCash_Left As Long, lngCost_Left As Long, lngCurrID_Left As Long

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private blnRequery As Boolean, blnEmptyHidden As Boolean
Private lngRecsCur As Long, lngTpp As Long, lngHiddenLbl_Left As Long
Private strCallingForm As String ', strCallersCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountHideTrans2_Hidden" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

190         blnRequery = False: blnEmptyHidden = False

200         .Sort_lbl.Visible = False
210         .Sort_line.Visible = False
220         lngSortLbl_Width = .Sort_lbl.Width

230         .curr_id.Visible = False
240         .curr_id_lbl.Visible = False
250         .curr_id_lbl_line.Visible = False

260         lngSortLine_Top = .Sort_line.Top
270         lngHiddenLbl_Left = .ledger_HIDDEN_lbl.Left
280         lngTotDesc_Left = .totdesc.Left
290         lngTotDesc_Width = .totdesc.Width
300         lngShareface_Left = .shareface_0.Left
310         lngJournalType_Left = .journaltype.Left
320         lngTransdate_Left = .transdate.Left
330         lngJournalNo_Width = (.transdate.Left - .journalno.Left)
340         lngICash_Left = .ICash.Left
350         lngPCash_Left = .PCash.Left
360         lngCost_Left = .Cost.Left
370         lngCurrID_Left = .curr_id.Left
380         lngCurrID_Width = (.curr_id.Width + (4& * lngTpp))

390       Else
400         Cancel = -1
410       End If
420     End With

430     If Cancel = -1 Then
440       Beep
450       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
460     End If

EXITP:
470     Exit Sub

ERRH:
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me

620       .transdate.ForeColor = CLR_DISABLED_FG
630       .transdate.BackColor = CLR_DISABLED_BG
640       .journaltype.ForeColor = CLR_DISABLED_FG
650       .journaltype.BackColor = CLR_DISABLED_BG
660       .shareface_0.ForeColor = CLR_DISABLED_FG
670       .shareface_1.ForeColor = CLR_DISABLED_FG
680       .shareface_2.ForeColor = CLR_DISABLED_FG
690       .shareface_3.ForeColor = CLR_DISABLED_FG
700       .shareface_4.ForeColor = CLR_DISABLED_FG
710       .shareface_box.BackColor = CLR_DISABLED_BG
720       .totdesc.ForeColor = CLR_DISABLED_FG
730       .totdesc.BackColor = CLR_DISABLED_BG
740       .ICash.ForeColor = CLR_DISABLED_FG
750       .ICash.BackColor = CLR_DISABLED_BG
760       .PCash.ForeColor = CLR_DISABLED_FG
770       .PCash.BackColor = CLR_DISABLED_BG
780       .Cost.ForeColor = CLR_DISABLED_FG
790       .Cost.BackColor = CLR_DISABLED_BG

800   On Error Resume Next
810       lngRecsCur = RecCnt  ' ** Function: Below.
820       If ERR.Number <> 0 Then
830   On Error GoTo ERRH
840         blnRequery = True
850         .TimerInterval = 1000&
860       ElseIf lngRecsCur = 0& Then
870   On Error GoTo ERRH
880         blnRequery = True
890         .TimerInterval = 1000&
900       Else
910   On Error GoTo ERRH
920         SortNow THIS_PROC  ' ** Procedure: Below
930       End If

940     End With

EXITP:
950     Exit Sub

ERRH:
960     Select Case ERR.Number
        Case Else
970       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
980     End Select
990     Resume EXITP

End Sub

Private Sub Form_Current()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim lngCurrID As Long
        Dim varTmp00 As Variant, strTmp01 As String

1010    With Me
1020  On Error Resume Next
1030      varTmp00 = .ledghid_grpnum
1040      If ERR.Number = 0 Then
1050  On Error GoTo ERRH
1060        lngCurrID = .curr_id
1070        If IsNull(varTmp00) = False Then
1080          .Parent.ledghid_grpnum = .ledghid_grpnum
1090          .Parent.shareface_subtot = .shareface_subtot
1100          .Parent.icash_subtot = .icash_subtot
1110          .Parent.pcash_subtot = .pcash_subtot
1120          .Parent.cost_subtot = .cost_subtot
1130          .Parent.curr_id = lngCurrID
1140          If (.ledghid_grpnum Mod 2) = 0 Then
1150            .Parent.ledghid_grpnum.BackColor = CLR_LTBLU
1160            Select Case .IsInvalid
                Case True
1170              .Parent.shareface_subtot.BackColor = CLR_VLTRED
1180              .Parent.IsInvalid_box.Visible = True
1190              .Parent.IsInvalid_lbl_box.Visible = True
1200              .Parent.IsInvalid_lbl1.Visible = True
1210              .Parent.IsInvalid_lbl2.Visible = True
1220            Case False
1230              .Parent.shareface_subtot.BackColor = CLR_LTBLU
1240              .Parent.IsInvalid_box.Visible = False
1250              .Parent.IsInvalid_lbl_box.Visible = False
1260              .Parent.IsInvalid_lbl1.Visible = False
1270              .Parent.IsInvalid_lbl2.Visible = False
1280            End Select
1290            .Parent.icash_subtot.BackColor = CLR_LTBLU
1300            .Parent.pcash_subtot.BackColor = CLR_LTBLU
1310            .Parent.cost_subtot.BackColor = CLR_LTBLU
1320            .Parent.curr_id.BackColor = CLR_LTBLU
1330          Else
1340            .Parent.ledghid_grpnum.BackColor = CLR_LTGRN
1350            Select Case .IsInvalid
                Case True
1360              .Parent.shareface_subtot.BackColor = CLR_VLTRED
1370              .Parent.IsInvalid_box.Visible = True
1380              .Parent.IsInvalid_lbl_box.Visible = True
1390              .Parent.IsInvalid_lbl1.Visible = True
1400              .Parent.IsInvalid_lbl2.Visible = True
1410            Case False
1420              .Parent.shareface_subtot.BackColor = CLR_LTGRN
1430              .Parent.IsInvalid_box.Visible = False
1440              .Parent.IsInvalid_lbl_box.Visible = False
1450              .Parent.IsInvalid_lbl1.Visible = False
1460              .Parent.IsInvalid_lbl2.Visible = False
1470            End Select
1480            .Parent.icash_subtot.BackColor = CLR_LTGRN
1490            .Parent.pcash_subtot.BackColor = CLR_LTGRN
1500            .Parent.cost_subtot.BackColor = CLR_LTGRN
1510            .Parent.curr_id.BackColor = CLR_LTGRN
1520          End If
1530          If lngCurrID = 150& Then
1540            strTmp01 = "Currency"
1550          Else
1560            If .currsym_symbol = "CHF" Then
1570              strTmp01 = "\C\H\F #,##.00;(\C\H\F #,##.00);\C\H\F 0.00;\C\H\F 0.00"
1580            Else
1590              strTmp01 = .currsym_symbol & "#,##0.00;(" & .currsym_symbol & "#,##0.00);" & _
                    .currsym_symbol & "0.00;" & .currsym_symbol & "0.00"
1600            End If
1610          End If
1620          .Parent.icash_subtot.Format = strTmp01
1630          .Parent.pcash_subtot.Format = strTmp01
1640          .Parent.cost_subtot.Format = strTmp01
1650        End If
1660      Else
1670  On Error GoTo ERRH
1680      End If
1690    End With

EXITP:
1700    Exit Sub

ERRH:
1710    Select Case ERR.Number
        Case Else
1720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1730    End Select
1740    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1810    intShiftDown = (Shift And acShiftMask) > 0
1820    intAltDown = (Shift And acAltMask) > 0
1830    intCtrlDown = (Shift And acCtrlMask) > 0

1840    intRetVal = KeyCode

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Focus Holder:     D {FocusHolder}

        ' ** Plain keys.
1850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1860      Select Case intRetVal
          Case vbKeyUp
1870        intRetVal = 0
1880        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1890      Case vbKeyDown
1900        intRetVal = 0
1910        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1920      End Select
1930    End If

        ' ** Alt keys.
1940    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1950      Select Case intRetVal
          Case vbKeyX
1960        With Me
1970          intRetVal = 0
1980          DoCmd.SelectObject acForm, .Parent.Name, False
1990          .Parent.cmdClose.SetFocus
2000          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountHideTrans2
2010        End With
2020      End Select
2030    End If

        ' ** Ctrl keys.
2040    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2050      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2060        With Me
2070          intRetVal = 0
2080          DoCmd.SelectObject acForm, .Parent.Name, False
2090          .Parent.chkShowJournalNo.SetFocus
2100        End With
2110      Case vbKeyUp, vbKeyPageUp, vbKeyHome
2120        intRetVal = 0
2130        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2140      Case vbKeyDown, vbKeyPageDown, vbKeyEnd
2150        intRetVal = 0
2160        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2170      End Select
2180    End If

        ' ** Ctrl-Shift keys.
2190    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2200      Select Case intRetVal
          Case vbKeyF
2210        With Me
2220          intRetVal = 0
2230          DoCmd.SelectObject acForm, .Parent.Name, False
2240          .Parent.FocusHolder.SetFocus
2250        End With
2260      Case vbKeyTab, vbKeyReturn
2270        With Me
2280          intRetVal = 0
2290          DoCmd.SelectObject acForm, .Parent.Name
2300          .Parent.cmdClose.SetFocus
2310        End With
2320      End Select
2330    End If

EXITP:
2340    KeyCode = intRetVal
2350    Exit Sub

ERRH:
2360    intRetVal = 0
2370    Select Case ERR.Number
        Case Else
2380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2390    End Select
2400    Resume EXITP

End Sub

Private Sub Form_Timer()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strTmp01 As String

2510    With Me
2520      .TimerInterval = 0&
2530      If blnRequery = True Then
2540        blnRequery = False
2550  On Error Resume Next
2560        strTmp01 = .RecordSource
2570        .RecordSource = vbNullString
2580        .RecordSource = strTmp01
2590        DoEvents
2600        lngRecsCur = RecCnt  ' ** Function: Below.
2610  On Error GoTo ERRH
2620        If lngRecsCur = 0& Then
2630          blnRequery = True
2640          .TimerInterval = 1000&
2650        Else
2660          SortNow "Form_Load"  ' ** Procedure: Below
2670        End If
2680      End If
2690    End With

EXITP:
2700    Exit Sub

ERRH:
2710    Select Case ERR.Number
        Case Else
2720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2730    End Select
2740    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
2810    intShiftDown = (Shift And acShiftMask) > 0
2820    intAltDown = (Shift And acAltMask) > 0
2830    intCtrlDown = (Shift And acCtrlMask) > 0

2840    intRetVal = KeyCode

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2870        With Me
2880          intRetVal = 0
2890          lngRecsCur = RecCnt  ' ** Function: Below.
2900          If .CurrentRecord < lngRecsCur Then
2910            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2920            .FocusHolder.SetFocus
2930          Else
2940            DoCmd.SelectObject acForm, .Parent.Name, False
2950            .Parent.chkShowJournalNo.SetFocus
2960          End If
2970        End With
2980      End Select
2990    End If

        ' ** Shift keys.
3000    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3010      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3020        With Me
3030          intRetVal = 0
3040          If .CurrentRecord > 1 Then
3050            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3060            .FocusHolder.SetFocus
3070          Else
3080            DoCmd.SelectObject acForm, .Parent.Name, False
3090            .Parent.cmdClose.SetFocus
3100          End If
3110        End With
3120      End Select
3130    End If

EXITP:
3140    KeyCode = intRetVal
3150    Exit Sub

ERRH:
3160    intRetVal = 0
3170    Select Case ERR.Number
        Case Else
3180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3190    End Select
3200    Resume EXITP

End Sub

Private Sub ledghid_grpnum_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "ledghid_grpnum_lbl_DblClick"

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

Private Sub PriorPeriod_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "PriorPeriod_lbl_DblClick"

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

Private Sub IsArch_lbl_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "IsArch_lbl_DblClick"

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

Private Sub journalno_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_lbl_DblClick"

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

Private Sub transdate_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

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

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

3810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

3910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

4010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

4110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

4210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

4310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

4410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_lbl_DblClick"

4510    If blnEmptyHidden = False Then
4520      SortNow THIS_PROC  ' ** Procedure: Below.
4530    End If

EXITP:
4540    Exit Sub

ERRH:
4550    Select Case ERR.Number
        Case Else
4560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4570    End Select
4580    Resume EXITP

End Sub

Private Sub ledger_HIDDEN_lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ** Easter Egg (Secret Key) to empty the LedgerHidden and tblLedgerHidden tables, so that they can be regenerated.

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "ledger_HIDDEN_lbl_MouseUp"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

4610    If Button = acRightButton And Shift = acCtrlMask + acAltMask Then
4620      blnEmptyHidden = True
4630      Beep
4640      If MsgBox("Rebuild the list of of hidden ledger entries?" & vbCrLf & vbCrLf & _
              "If you've received a 'Group Size Not Accommodated' message," & vbCrLf & _
              "or other error, when entering the hidden transaction window," & vbCrLf & _
              "rebuilding this list may solve the problem." & vbCrLf & vbCrLf & _
              "This will not change the status of any entries that are currently" & vbCrLf & _
              "hidden, only rebuild a secondary list of specific matches.", _
              vbQuestion + vbOKCancel, ("Rebuild Hidden Ledger List" & Space(50))) = vbOK Then
4650        DoCmd.Hourglass True
4660        DoEvents
4670        Set dbs = CurrentDb
4680        With dbs
              ' ** Empty LedgerHidden (in prep for regen).
4690          Set qdf = .QueryDefs("qryAccountHide_11j")
4700          qdf.Execute
4710          Set qdf = Nothing
              ' ** Empty tblLedgerHidden (in prep for regen).
4720          Set qdf = .QueryDefs("qryAccountHide_11k")
4730          qdf.Execute
4740          Set qdf = Nothing
4750          .Close
4760        End With
4770        Set dbs = Nothing
4780        DoCmd.Hourglass False
4790        MsgBox "Finished." & vbCrLf & "This secondary list will now regenerate the next" & vbCrLf & _
              "time you click the 'Hide Transactions' button.", vbInformation + vbOKOnly, ("Finished" & Space(40))
4800        blnEmptyHidden = False
4810      Else
4820        blnEmptyHidden = False
4830      End If
4840    End If

EXITP:
4850    Set qdf = Nothing
4860    Set dbs = Nothing
4870    Exit Sub

ERRH:
4880    DoCmd.Hourglass False
4890    Select Case ERR.Number
        Case Else
4900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4910    End Select
4920    Resume EXITP

End Sub

Public Sub ShowJournalNo_Sub(blnShow As Boolean)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim blnCurrID As Boolean

5010    With Me
5020      DoCmd.Hourglass True
5030      DoEvents
5040      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
5050        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5060      End If
5070      gblnHasForExThis = HasForEx_Acct(gstrAccountNo, "L")  ' ** Module Function: modCurrencyFuncs.
5080      Select Case gblnHasForEx
          Case True
5090        Select Case gblnHasForExThis
            Case True
5100          blnCurrID = True
5110        Case False
5120          blnCurrID = False
5130        End Select
5140      Case False
5150        blnCurrID = False
5160      End Select
5170      blnSortHere = False: blnResort = False
5180      Select Case blnShow
          Case True
5190        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
5200        .totdesc.Left = lngTotDesc_Left
5210        .totdesc_lbl.Left = .totdesc.Left
5220        .totdesc_lbl_line.Left = .totdesc_lbl.Left
5230        Select Case blnCurrID
            Case True
5240          .totdesc.Width = lngTotDesc_Width  ' ** Design View width
5250        Case False
5260          .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)  ' ** Without curr_id, but on open,
5270        End Select                                               ' ** other fields haven't been moved yet
5280        .totdesc_lbl.Width = .totdesc.Width
5290        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
5300        If blnSortHere = True Then
5310          .Sort_line.Left = .totdesc_lbl.Left
5320          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
5330        End If
5340        blnSortHere = False
5350        If .Sort_line.Left = .journaltype_lbl.Left Then blnSortHere = True
5360        .journaltype.Left = lngJournalType_Left
5370        .journaltype_lbl.Left = .journaltype.Left
5380        .journaltype_lbl_line.Left = .journaltype_lbl.Left
5390        If blnSortHere = True Then
5400          .Sort_line.Left = .journaltype_lbl.Left
5410          .Sort_lbl.Left = ((.journaltype_lbl.Left + .journaltype_lbl.Width) - lngSortLbl_Width)
5420        End If
5430        blnSortHere = False
5440        If .Sort_line.Left = .transdate_lbl.Left Then blnSortHere = True
5450        .transdate.Left = lngTransdate_Left
5460        .transdate_lbl.Left = .transdate.Left
5470        .transdate_lbl_line.Left = .transdate_lbl.Left
5480        If blnSortHere = True Then
5490          .Sort_line.Left = .transdate_lbl.Left
5500          .Sort_lbl.Left = ((.transdate_lbl.Left + .transdate_lbl.Width) - lngSortLbl_Width)
5510        End If
5520        blnSortHere = False
5530        .journalno.Visible = True
5540        .journalno_lbl.Visible = True
5550        .journalno_lbl2.Visible = True
5560        .journalno_lbl_line.Visible = True
5570      Case False
5580        If .journalno_Sort_line1.Visible = True Then blnSortHere = True
5590        .journalno.Visible = False
5600        .journalno_lbl.Visible = False
5610        .journalno_lbl2.Visible = False
5620        .journalno_lbl_line.Visible = False
5630        If blnSortHere = True Then
5640          .journalno_Sort_line1.Visible = False
5650          .journalno_Sort_line2.Visible = False
5660          .Sort_lbl.Visible = False
5670          blnResort = True
5680        End If
5690        blnSortHere = False
5700        If .Sort_line.Left = .transdate_lbl.Left Then blnSortHere = True
5710        .transdate.Left = (lngTransdate_Left - lngJournalNo_Width)
5720        .transdate_lbl.Left = .transdate.Left
5730        .transdate_lbl_line.Left = .transdate_lbl.Left
5740        If blnSortHere = True Then
5750          .Sort_line.Left = .transdate_lbl.Left
5760          .Sort_lbl.Left = ((.transdate_lbl.Left + .transdate_lbl.Width) - lngSortLbl_Width)
5770        End If
5780        blnSortHere = False
5790        If .Sort_line.Left = .journaltype_lbl.Left Then blnSortHere = True
5800        .journaltype.Left = (lngJournalType_Left - lngJournalNo_Width)
5810        .journaltype_lbl.Left = .journaltype.Left
5820        .journaltype_lbl_line.Left = .journaltype_lbl.Left
5830        If blnSortHere = True Then
5840          .Sort_line.Left = .journaltype_lbl.Left
5850          .Sort_lbl.Left = ((.journaltype_lbl.Left + .journaltype_lbl.Width) - lngSortLbl_Width)
5860        End If
5870        blnSortHere = False
5880        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
5890        .totdesc.Left = (lngTotDesc_Left - lngJournalNo_Width)
5900        .totdesc_lbl.Left = .totdesc.Left
5910        .totdesc_lbl_line.Left = .totdesc_lbl.Left
5920        Select Case blnCurrID
            Case True
5930          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)  ' ** Design View minus journalno.
5940        Case False
5950          .totdesc.Width = ((lngTotDesc_Width + lngJournalNo_Width) + lngCurrID_Width)  ' ** Minus both journalno and curr_id, but on open,
5960        End Select                                                                      ' ** other fields haven't been moved yet.
5970        .totdesc_lbl.Width = .totdesc.Width
5980        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
5990        If blnSortHere = True Then
6000          .Sort_line.Left = .totdesc_lbl.Left
6010          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
6020        End If
6030        blnSortHere = False
6040      End Select
6050      If blnResort = True Then
6060        SortNow "Form_Load"  ' ** Procedure: Below.
6070      End If
6080    End With

EXITP:
6090    Exit Sub

ERRH:
6100    DoCmd.Hourglass False
6110    Select Case ERR.Number
        Case Else
6120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6130    End Select
6140    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnShow As Boolean)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim blnJournalNo As Boolean

6210    With Me
6220      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
6230        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6240      End If
6250      blnSortHere = False: blnResort = False
6260      blnJournalNo = .Parent.chkShowJournalNo
6270      Select Case blnShow
          Case True
6280        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
6290        Select Case blnJournalNo
            Case True
6300          .totdesc.Width = lngTotDesc_Width
6310        Case False
6320          .totdesc.Width = (lngTotDesc_Width + lngJournalNo_Width)
6330        End Select
6340        .totdesc_lbl.Width = .totdesc.Width
6350        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
6360        If blnSortHere = True Then
6370          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
6380          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
6390        End If
6400        blnSortHere = False
6410        If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
6420        .shareface_0.Left = lngShareface_Left
6430        .shareface_1.Left = .shareface_0.Left
6440        .shareface_2.Left = .shareface_0.Left
6450        .shareface_3.Left = .shareface_0.Left
6460        .shareface_4.Left = .shareface_0.Left
6470        .shareface_box.Left = .shareface_0.Left
6480        .IsInvalid_box1.Left = (.shareface_0.Left - (3& * lngTpp))
6490        .IsInvalid_box2.Left = .IsInvalid_box1.Left
6500        .shareface_lbl.Left = .shareface_0.Left
6510        .shareface_lbl_line.Left = .shareface_0.Left
6520        If blnSortHere = True Then
6530          .Sort_line.Left = .shareface_lbl.Left
6540          .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
6550        End If
6560        blnSortHere = False
6570        If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
6580        .ICash.Left = lngICash_Left
6590        .icash_str.Left = .ICash.Left
6600        .icash_str.Visible = True
6610        .ICash.Visible = False
6620        .icash_lbl.Left = .ICash.Left
6630        .icash_lbl_line.Left = .icash_lbl.Left
6640        If blnSortHere = True Then
6650          .Sort_line.Left = .icash_lbl.Left
6660          .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
6670        End If
6680        blnSortHere = False
6690        If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
6700        .PCash.Left = lngPCash_Left
6710        .pcash_str.Left = .PCash.Left
6720        .pcash_str.Visible = True
6730        .PCash.Visible = False
6740        .pcash_lbl.Left = .PCash.Left
6750        .pcash_lbl_line.Left = .pcash_lbl.Left
6760        If blnSortHere = True Then
6770          .Sort_line.Left = .pcash_lbl.Left
6780          .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
6790        End If
6800        blnSortHere = False
6810        If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
6820        .Cost.Left = lngCost_Left
6830        .cost_str.Left = .Cost.Left
6840        .cost_str.Visible = True
6850        .Cost.Visible = False
6860        .cost_lbl.Left = .Cost.Left
6870        .cost_lbl_line.Left = .cost_lbl.Left
6880        If blnSortHere = True Then
6890          .Sort_line.Left = .cost_lbl.Left
6900          .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
6910        End If
6920        blnSortHere = False
6930        .curr_id.Left = lngCurrID_Left
6940        .curr_id_lbl.Left = .curr_id.Left
6950        .curr_id_lbl_line.Left = .curr_id_lbl.Left
6960        .curr_id.Visible = True
6970        .curr_id_lbl.Visible = True
6980        .curr_id_lbl_line.Visible = True
6990      Case False
7000        If .Sort_line.Left = .curr_id_lbl.Left Then blnSortHere = True
7010        .curr_id.Visible = False
7020        .curr_id_lbl.Visible = False
7030        .curr_id_lbl_line.Visible = False
7040        .curr_id.Left = 0&
7050        .curr_id_lbl.Left = 0&
7060        .curr_id_lbl_line.Left = 0&
7070        If blnSortHere = True Then
7080          .Sort_line.Visible = False
7090          .Sort_lbl.Visible = False
7100          blnResort = True
7110        End If
7120        blnSortHere = False
7130        If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
7140        .Cost.Left = (lngCost_Left + lngCurrID_Width)
7150        .cost_str.Left = .Cost.Left
7160        .Cost.Visible = True
7170        .cost_str.Visible = False
7180        .cost_lbl.Left = .Cost.Left
7190        .cost_lbl_line.Left = .cost_lbl.Left
7200        If blnSortHere = True Then
7210          .Sort_line.Left = .cost_lbl.Left
7220          .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
7230        End If
7240        blnSortHere = False
7250        If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
7260        .PCash.Left = (lngPCash_Left + lngCurrID_Width)
7270        .pcash_str.Left = .PCash.Left
7280        .PCash.Visible = True
7290        .pcash_str.Visible = False
7300        .pcash_lbl.Left = .PCash.Left
7310        .pcash_lbl_line.Left = .pcash_lbl.Left
7320        If blnSortHere = True Then
7330          .Sort_line.Left = .pcash_lbl.Left
7340          .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
7350        End If
7360        blnSortHere = False
7370        If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
7380        .ICash.Left = (lngICash_Left + lngCurrID_Width)
7390        .icash_str.Left = .ICash.Left
7400        .ICash.Visible = True
7410        .icash_str.Visible = False
7420        .icash_lbl.Left = .ICash.Left
7430        .icash_lbl_line.Left = .icash_lbl.Left
7440        If blnSortHere = True Then
7450          .Sort_line.Left = .icash_lbl.Left
7460          .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
7470        End If
7480        blnSortHere = False
7490        If .Sort_line.Left = .shareface_lbl.Left Then blnSortHere = True
7500        .shareface_0.Left = (lngShareface_Left + lngCurrID_Width)
7510        .shareface_1.Left = .shareface_0.Left
7520        .shareface_2.Left = .shareface_0.Left
7530        .shareface_3.Left = .shareface_0.Left
7540        .shareface_4.Left = .shareface_0.Left
7550        .shareface_box.Left = .shareface_0.Left
7560        .IsInvalid_box1.Left = (.shareface_0.Left - (3& * lngTpp))
7570        .IsInvalid_box2.Left = .IsInvalid_box1.Left
7580        .shareface_lbl.Left = .shareface_0.Left
7590        .shareface_lbl_line.Left = .shareface_0.Left
7600        If blnSortHere = True Then
7610          .Sort_line.Left = .shareface_lbl.Left
7620          .Sort_lbl.Left = ((.shareface_lbl.Left + .shareface_lbl.Width) - lngSortLbl_Width)
7630        End If
7640        blnSortHere = False
7650        If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
7660        .totdesc.Width = (lngTotDesc_Width + lngCurrID_Width)
7670        .totdesc_lbl.Width = .totdesc.Width
7680        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
7690        If blnSortHere = True Then
7700          .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
7710          .Sort_lbl.Left = ((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width)
7720        End If
7730        blnSortHere = False
7740      End Select
7750    End With

EXITP:
7760    Exit Sub

ERRH:
7770    Select Case ERR.Number
        Case Else
7780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7790    End Select
7800    Resume EXITP

End Sub

Public Function RecCnt() As Long

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

7910    lngRetVal = 0&

7920    With Me
7930  On Error Resume Next
7940      Set rst = .RecordsetClone
7950      If ERR.Number = 0 Then
7960  On Error GoTo ERRH
7970        With rst
7980          If .BOF = True And .EOF = True Then
                ' ** Shouldn't really be here if that's the case.
7990          Else
8000            .MoveLast
8010            lngRetVal = .RecordCount
8020          End If
8030          .Close
8040        End With
8050      Else
8060  On Error GoTo ERRH
8070      End If
8080    End With

EXITP:
8090    Set rst = Nothing
8100    RecCnt = lngRetVal
8110    Exit Function

ERRH:
8120    DoCmd.Hourglass False
8130    lngRetVal = 0&
8140    Select Case ERR.Number
        Case Else
8150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8160    End Select
8170    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

8210    Select Case intWhere
        Case acCmdRecordsGoToFirst
8220      DoCmd.RunCommand acCmdRecordsGoToFirst
8230    Case acCmdRecordsGoToLast
8240      DoCmd.RunCommand acCmdRecordsGoToLast
8250    Case acCmdRecordsGoToPrevious
8260      DoCmd.RunCommand acCmdRecordsGoToPrevious
8270    Case acCmdRecordsGoToNext
8280      DoCmd.RunCommand acCmdRecordsGoToNext
8290    Case Else
8300      If IsMissing(varID) = False Then
8310        With Me
8320          Set rst = .RecordsetClone
8330          With rst
8340            .FindFirst "[ledghid_id] = " & CStr(varID)
8350            If .NoMatch = False Then
8360              Me.Bookmark = .Bookmark
8370            End If
8380            .Close
8390          End With
8400        End With
8410      End If
8420    End Select

EXITP:
8430    Set rst = Nothing
8440    Exit Sub

ERRH:
8450    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8460    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8470    Case Else
8480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8490    End Select
8500    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [ledghid_grpnum], [ledghid_ord]"

8610    With Me
8620      .Controls(strSortLbl).Visible = False
8630      .Controls(strSortLine).Visible = False
8640      .journalno_Sort_line1.Visible = False
8650      .journalno_Sort_line2.Visible = False
8660      .PriorPeriod_Sort_line1.Visible = False
8670      .PriorPeriod_Sort_line2.Visible = False
8680      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
8690      If .ledger_HIDDEN_lbl.Left <> lngHiddenLbl_Left Then
8700        .ledger_HIDDEN_lbl.Left = lngHiddenLbl_Left
8710      End If
8720      DoEvents
8730      If strProc = "Form_Load" Then
8740        strCalled = "ledghid_grpnum"
8750        strSortNow = strSortOrig
8760        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - (3& * lngTpp))
8770        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
8780        lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
8790        lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
8800        lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
8810        .Controls(strSortLbl).Top = lngSortLbl_Top
8820        .Controls(strSortLbl).Left = lngSortLbl_Left
8830        .Controls(strSortLine).Left = lngSortLine_Left
8840        .Controls(strSortLine).Width = lngSortLine_Width
8850        .Controls(strSortLbl).Caption = strArwUp
8860        .Controls(strSortLbl).ForeColor = CLR_DKBLU
8870        DoEvents
8880      Else
8890        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: revcode_ID_lbl_DblClick
8900        lngSortLbl_Top = .Controls(strCalled & "_lbl").Top
8910        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
8920        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
8930        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
8940        .Controls(strSortLbl).Top = lngSortLbl_Top
8950        .Controls(strSortLbl).Left = lngSortLbl_Left
8960        .Controls(strSortLine).Left = lngSortLine_Left
8970        .Controls(strSortLine).Width = lngSortLine_Width
8980        If strCalled = "ledghid_grpnum" Then
8990          strSortAsc = strSortOrig
9000          lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - (3& * lngTpp))
9010          lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
9020          lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
9030          lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
9040          .Controls(strSortLbl).Top = lngSortLbl_Top
9050          .Controls(strSortLbl).Left = lngSortLbl_Left
9060          .Controls(strSortLine).Left = lngSortLine_Left
9070          .Controls(strSortLine).Width = lngSortLine_Width
9080          If strSortNow = strSortAsc Then
9090            strSortNow = "[ledghid_grpnum] DESC, [ledghid_ord] DESC"
9100            .Controls(strSortLbl).Caption = strArwDn
9110            .Controls(strSortLbl).ForeColor = CLR_DKRED
9120          Else
9130            strSortNow = strSortAsc
9140            .Controls(strSortLbl).Caption = strArwUp
9150            .Controls(strSortLbl).ForeColor = CLR_DKBLU
9160          End If
9170          DoEvents
9180        Else
9190          .Controls(strSortLbl).Caption = strArwUp
9200          .Controls(strSortLbl).ForeColor = CLR_DKBLU
9210          Select Case strCalled
              Case "PriorPeriod"
9220            strSortAsc = "[PriorPeriod]" & strStdAsc
9230            lngSortLbl_Top = (lngSortLbl_Top + (1& * lngTpp))
9240            lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
9250            .Controls(strSortLbl).Top = lngSortLbl_Top
9260            .Controls(strSortLbl).Left = lngSortLbl_Left
9270          Case "IsArch"
9280            strSortAsc = "[IsArch]" & strStdAsc
9290            lngSortLbl_Top = (lngSortLbl_Top - (3& * lngTpp))
9300            lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
9310            .Controls(strSortLbl).Top = lngSortLbl_Top
9320            .Controls(strSortLbl).Left = lngSortLbl_Left
9330          Case "journalno"
9340            strSortAsc = "[journalno]"
9350            lngSortLbl_Top = (lngSortLbl_Top + (1& * lngTpp))
9360            lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
9370            .Controls(strSortLbl).Top = lngSortLbl_Top
9380            .Controls(strSortLbl).Left = lngSortLbl_Left
9390          Case "transdate"
9400            strSortAsc = "[transdate]" & strStdAsc
9410            lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
9420            .Controls(strSortLbl).Top = lngSortLbl_Top
9430          Case "journaltype"
9440            strSortAsc = "[journaltype]" & strStdAsc
9450            lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
9460            .Controls(strSortLbl).Top = lngSortLbl_Top
9470          Case "shareface"
9480            strSortAsc = "[shareface]" & strStdAsc
9490            lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
9500            .Controls(strSortLbl).Top = lngSortLbl_Top
9510          Case "totdesc"
9520            strSortAsc = "[totdesc]" & strStdAsc
9530            lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
9540            .Controls(strSortLbl).Top = lngSortLbl_Top
9550          Case "icash"
9560            strSortAsc = "[icash]" & strStdAsc
9570            lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
9580            .Controls(strSortLbl).Top = lngSortLbl_Top
9590          Case "pcash"
9600            strSortAsc = "[pcash]" & strStdAsc
9610            lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
9620            .Controls(strSortLbl).Top = lngSortLbl_Top
9630          Case "cost"
9640            strSortAsc = "[cost]" & strStdAsc
9650            lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
9660            .Controls(strSortLbl).Top = lngSortLbl_Top
9670          Case "curr_id"
9680            strSortAsc = "[curr_code]"
9690            lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
9700            .Controls(strSortLbl).Left = lngSortLbl_Left
9710          Case "ledger_HIDDEN"
9720            strSortAsc = "[ledger_HIDDEN]" & strStdAsc
9730            lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
9740            lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
9750            .Controls(strSortLbl).Top = lngSortLbl_Top
9760            .Controls(strSortLbl).Left = lngSortLbl_Left
9770          End Select
9780          If strSortNow = strSortAsc Then
9790            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
9800            Select Case intCnt
                Case 1
9810              strTmp01 = strSortAsc & " DESC"
9820            Case 2
9830              intPos01 = InStr(strSortAsc, ",")
9840              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9850              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
9860            Case 3
9870              intPos01 = InStr(strSortAsc, ",")
9880              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9890              strTmp02 = Mid(strSortAsc, intPos01)
9900              intPos01 = InStr(2, strTmp02, ",")
9910              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9920              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
9930            Case 4
9940              intPos01 = InStr(strSortAsc, ",")
9950              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9960              strTmp02 = Mid(strSortAsc, intPos01)
9970              intPos01 = InStr(2, strTmp02, ",")
9980              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9990              strTmp02 = Mid(strTmp02, intPos01)
10000             intPos01 = InStr(2, strTmp02, ",")
10010             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10020             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10030           End Select
10040           strSortNow = strTmp01
10050           .Controls(strSortLbl).Caption = strArwDn
10060           .Controls(strSortLbl).ForeColor = CLR_DKRED
10070         Else
10080           strSortNow = strSortAsc
10090         End If
10100         DoEvents
10110       End If
10120     End If
10130     Select Case strCalled
          Case "PriorPeriod"
10140       .PriorPeriod_Sort_line1.Visible = True
10150       .PriorPeriod_Sort_line2.Visible = True
10160     Case "journalno"
10170       .journalno_Sort_line1.Visible = True
10180       .journalno_Sort_line2.Visible = True
10190     Case Else
10200       .Controls(strSortLine).Visible = True
10210     End Select
10220     .Controls(strSortLbl).Visible = True
10230     DoEvents
10240     .OrderBy = strSortNow
10250     .OrderByOn = True
10260   End With

        ' ** Potential Sorts:
        ' **   [ledghid_grpnum], [ledghid_ord]
        ' **   [PriorPeriod], [ledghid_grpnum], [ledghid_ord]
        ' **   [IsArch], [ledghid_grpnum], [ledghid_ord]
        ' **   [journalno]
        ' **   [transdate], [ledghid_grpnum], [ledghid_ord]
        ' **   [journaltype], [ledghid_grpnum], [ledghid_ord]
        ' **   [shareface], [ledghid_grpnum], [ledghid_ord]
        ' **   [totdesc], [ledghid_grpnum], [ledghid_ord]
        ' **   [icash], [ledghid_grpnum], [ledghid_ord]
        ' **   [pcash], [ledghid_grpnum], [ledghid_ord]
        ' **   [cost], [ledghid_grpnum], [ledghid_ord]
        ' **   [ledger_HIDDEN], [ledghid_grpnum], [ledghid_ord]

EXITP:
10270   Exit Sub

ERRH:
10280   Select Case ERR.Number
        Case Else
10290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10300   End Select
10310   Resume EXITP

End Sub

Public Function SortNow_Get() As String

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

10410   strRetVal = strSortNow

EXITP:
10420   SortNow_Get = strRetVal
10430   Exit Function

ERRH:
10440   Select Case ERR.Number
        Case Else
10450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10460   End Select
10470   Resume EXITP

End Function

Public Function SortOrig_Get() As String

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "SortOrig_Get"

        Dim strRetVal As String

10510   strRetVal = strSortOrig

EXITP:
10520   SortOrig_Get = strRetVal
10530   Exit Function

ERRH:
10540   Select Case ERR.Number
        Case Else
10550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10560   End Select
10570   Resume EXITP

End Function
