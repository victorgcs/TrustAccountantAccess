VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountTransactions_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountTransactions_Sub"

'VGC 04/15/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmAccountTransactions}
' **   Include Archive:  H {chkIncludeArchive on frmAccountTransactions}
' **   Include Currency: U {chkIncludeCurrency on frmAccountTransactions}
' **   Exit:             X {cmdClose on frmAccountTransactions}

Private Const strSortOrig As String = "[transdate] DESC, [JournalType_Order], [journalno]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngForm_Width As Long, lngCurrID_Left As Long, lngCurrIDLbl_Left As Long, lngTotWidth_Diff As Long
Private lngICashUSD_Left As Long, lngPCashUSD_Left As Long, lngCostUSD_Left As Long, lngVLine_Offset As Long

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
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
150       If strCallingForm = "frmAccountTransactions" Then

160         lngSortLbl_Width = .Sort_lbl.Width

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180         .form_width_line.Left = 0&  ' ** Just get it out of the way.

190         lngForm_Width = .Width
200         lngCurrID_Left = .curr_id.Left
210         lngCurrIDLbl_Left = .curr_id_lbl.Left
220         lngICashUSD_Left = .icash_usd.Left
230         lngPCashUSD_Left = .pcash_usd.Left
240         lngCostUSD_Left = .cost_usd.Left
250         lngTotWidth_Diff = ((.cost_usd.Left + .cost_usd.Width) - (.Cost.Left + .Cost.Width))
260         lngVLine_Offset = (.Cost.Left - .Detail_vline05.Left)

270       Else
280         Cancel = -1
290       End If
300     End With

310     If Cancel = -1 Then
320       Beep
330       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
340     End If

EXITP:
350     Exit Sub

ERRH:
360     Select Case ERR.Number
        Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
380     End Select
390     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
420     Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyUp
570         intRetVal = 0
580         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
590       Case vbKeyDown
600         intRetVal = 0
610         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
620       Case vbKeyEscape
630         With Me
640           intRetVal = 0
650           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountTransactions.
660         End With
670       End Select
680     End If

        ' ** Alt keys.
690     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
700       Select Case intRetVal
          Case vbKeyX
710         With Me
720           intRetVal = 0
730           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountTransactions.
740         End With
750       End Select
760     End If

        ' ** Ctrl keys.
770     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
780       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
790         With Me
800           intRetVal = 0
810           DoCmd.SelectObject acForm, .Parent.Name, False
820           If .Parent.chkIncludeCurrency.Enabled = True Then
830             .Parent.chkIncludeCurrency.SetFocus
840           ElseIf .Parent.chkIncludeArchive.Enabled = True Then
850             .Parent.chkIncludeArchive.SetFocus
860           Else
870             .Parent.cmdClose.SetFocus
880           End If
890         End With
900       Case vbKeyPageUp, vbKeyHome
910         intRetVal = 0
920         MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
930       Case vbKeyPageDown, vbKeyEnd
940         intRetVal = 0
950         MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
960       End Select
970     End If

        ' ** Ctrl-Shift keys.
980     If intCtrlDown And (Not intAltDown) And intShiftDown Then
990       Select Case intRetVal
          Case vbKeyF
1000        With Me
1010          intRetVal = 0
1020          DoCmd.SelectObject acForm, .Parent.Name, False
1030          .Parent.FocusHolder.SetFocus
1040        End With
1050      Case vbKeyTab, vbKeyReturn
1060        With Me
1070          intRetVal = 0
1080          DoCmd.SelectObject acForm, .Parent.Name, False
1090          .Parent.cmdClose.SetFocus
1100        End With
1110      End Select
1120    End If

EXITP:
1130    KeyCode = intRetVal
1140    Exit Sub

ERRH:
1150    intRetVal = 0
1160    Select Case ERR.Number
        Case Else
1170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1180    End Select
1190    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1270        With Me
1280          intRetVal = 0
1290          .transdate.SetFocus
1300        End With
1310      End Select
1320    End If

        ' ** Shift keys.
1330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1350        With Me
1360          intRetVal = 0
1370          If .CurrentRecord > 1 Then
1380            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1390            If .cost_usd.Visible = True And .cost_usd.Enabled = True Then
1400              .cost_usd.SetFocus
1410            Else
1420              .Cost.SetFocus
1430            End If
1440          Else
1450            DoCmd.SelectObject acForm, .Parent.Name, False
1460            .Parent.cmdClose.SetFocus
1470          End If
1480        End With
1490      End Select
1500    End If

EXITP:
1510    KeyCode = intRetVal
1520    Exit Sub

ERRH:
1530    intRetVal = 0
1540    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
1550    Case Else
1560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1570    End Select
1580    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

1610    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
1620    Exit Sub

ERRH:
1630    Select Case ERR.Number
        Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub transdate_KeyDown(KeyCode As Integer, Shift As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1720    intShiftDown = (Shift And acShiftMask) > 0
1730    intAltDown = (Shift And acAltMask) > 0
1740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1770        With Me
1780          intRetVal = 0
1790          .journaltype.SetFocus
1800        End With
1810      End Select
1820    End If

        ' ** Shift keys.
1830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1850        With Me
1860          intRetVal = 0
1870          If .CurrentRecord > 1 Then
1880            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1890            If .cost_usd.Visible = True And .cost_usd.Enabled = True Then
1900              .cost_usd.SetFocus
1910            Else
1920              .Cost.SetFocus
1930            End If
1940          Else
1950            DoCmd.SelectObject acForm, .Parent.Name, False
1960            .Parent.cmdClose.SetFocus
1970          End If
1980        End With
1990      End Select
2000    End If

EXITP:
2010    KeyCode = intRetVal
2020    Exit Sub

ERRH:
2030    intRetVal = 0
2040    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2050    Case Else
2060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2070    End Select
2080    Resume EXITP

End Sub

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

2110    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

2210    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

2310    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub icash_str_Enter()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_str_Enter"

2410    With Me
2420      If .curr_id = 150 Then
2430        .icash_str.BackColor = CLR_WHT
2440      Else
2450        .icash_str.BackColor = CLR_LTGRN
2460      End If
2470    End With

EXITP:
2480    Exit Sub

ERRH:
2490    Select Case ERR.Number
        Case Else
2500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2510    End Select
2520    Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

2610    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub pcash_str_Enter()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_str_Enter"

2710    With Me
2720      If .curr_id = 150 Then
2730        .pcash_str.BackColor = CLR_WHT
2740      Else
2750        .pcash_str.BackColor = CLR_LTGRN
2760      End If
2770    End With

EXITP:
2780    Exit Sub

ERRH:
2790    Select Case ERR.Number
        Case Else
2800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2810    End Select
2820    Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

2910    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2920    Exit Sub

ERRH:
2930    Select Case ERR.Number
        Case Else
2940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950    End Select
2960    Resume EXITP

End Sub

Private Sub cost_str_Enter()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_str_Enter"

3010    With Me
3020      If .curr_id = 150 Then
3030        .cost_str.BackColor = CLR_WHT
3040      Else
3050        .cost_str.BackColor = CLR_LTGRN
3060      End If
3070    End With

EXITP:
3080    Exit Sub

ERRH:
3090    Select Case ERR.Number
        Case Else
3100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3110    End Select
3120    Resume EXITP

End Sub

Private Sub cost_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3220    intShiftDown = (Shift And acShiftMask) > 0
3230    intAltDown = (Shift And acAltMask) > 0
3240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3270        With Me
3280          intRetVal = 0
3290          If .curr_id.Visible = True And .curr_id.Enabled = True Then
3300            .curr_id.SetFocus
3310          Else
3320            lngRecsCur = RecCnt  ' ** Function: Below.
3330            If .CurrentRecord < lngRecsCur Then
3340              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3350              .transdate.SetFocus
3360            Else
3370              DoCmd.SelectObject acForm, .Parent.Name, False
3380              If .Parent.chkIncludeArchive.Enabled = True Then
3390                .Parent.chkIncludeArchive.SetFocus
3400              ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
3410                .Parent.chkIncludeCurrency.SetFocus
3420              Else
3430                .Parent.cmdClose.SetFocus
3440              End If
3450            End If
3460          End If
3470        End With
3480      End Select
3490    End If

        ' ** Shift keys.
3500    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3510      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3520        With Me
3530          intRetVal = 0
3540          .PCash.SetFocus
3550        End With
3560      End Select
3570    End If

EXITP:
3580    KeyCode = intRetVal
3590    Exit Sub

ERRH:
3600    intRetVal = 0
3610    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
3620    Case Else
3630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3640    End Select
3650    Resume EXITP

End Sub

Private Sub cost_str_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_str_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3720    intShiftDown = (Shift And acShiftMask) > 0
3730    intAltDown = (Shift And acAltMask) > 0
3740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3770        With Me
3780          intRetVal = 0
3790          If .curr_id.Visible = True And .curr_id.Enabled = True Then
3800            .curr_id.SetFocus
3810          Else
3820            lngRecsCur = RecCnt  ' ** Function: Below.
3830            If .CurrentRecord < lngRecsCur Then
3840              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3850              .transdate.SetFocus
3860            Else
3870              DoCmd.SelectObject acForm, .Parent.Name, False
3880              If .Parent.chkIncludeArchive.Enabled = True Then
3890                .Parent.chkIncludeArchive.SetFocus
3900              ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
3910                .Parent.chkIncludeCurrency.SetFocus
3920              Else
3930                .Parent.cmdClose.SetFocus
3940              End If
3950            End If
3960          End If
3970        End With
3980      End Select
3990    End If

        ' ** Shift keys.
4000    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4010      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4020        With Me
4030          intRetVal = 0
4040          .pcash_str.SetFocus
4050        End With
4060      End Select
4070    End If

EXITP:
4080    KeyCode = intRetVal
4090    Exit Sub

ERRH:
4100    intRetVal = 0
4110    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
4120    Case Else
4130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4140    End Select
4150    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

4210    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub curr_id_Enter()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_Enter"

4310    With Me
4320      If .curr_id = 150& Then
4330        .curr_id.BackColor = CLR_WHT
4340      Else
4350        .curr_id.BackColor = CLR_LTGRN
4360      End If
4370    End With

EXITP:
4380    Exit Sub

ERRH:
4390    Select Case ERR.Number
        Case Else
4400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4410    End Select
4420    Resume EXITP

End Sub

Private Sub icash_usd_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_usd_lbl_DblClick"

4510    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub pcash_usd_lbl_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_usd_lbl_DblClick"

4610    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub cost_usd_lbl_DblClick(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_usd_lbl_DblClick"

4710    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub cost_usd_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_usd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4820    intShiftDown = (Shift And acShiftMask) > 0
4830    intAltDown = (Shift And acAltMask) > 0
4840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4870        With Me
4880          intRetVal = 0
4890          lngRecsCur = RecCnt  ' ** Function: Below.
4900          If .CurrentRecord < lngRecsCur Then
4910            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4920            .transdate.SetFocus
4930          Else
4940            DoCmd.SelectObject acForm, .Parent.Name, False
4950            If .Parent.chkIncludeArchive.Enabled = True Then
4960              .Parent.chkIncludeArchive.SetFocus
4970            ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
4980              .Parent.chkIncludeCurrency.SetFocus
4990            Else
5000              .Parent.cmdClose.SetFocus
5010            End If
5020          End If
5030        End With
5040      End Select
5050    End If

        ' ** Shift keys.
5060    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5070      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5080        With Me
5090          intRetVal = 0
5100          .pcash_usd.SetFocus
5110        End With
5120      End Select
5130    End If

EXITP:
5140    KeyCode = intRetVal
5150    Exit Sub

ERRH:
5160    intRetVal = 0
5170    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
5180    Case Else
5190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5200    End Select
5210    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnShow As Boolean)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean

5310    With Me
5320      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5330        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5340      End If
5350      blnSortHere = False: blnResort = False
5360      Select Case blnShow
          Case True
5370        .FocusHolder.SetFocus
5380        DoEvents
5390        .Width = lngForm_Width
5400        .cost_usd.Left = lngCostUSD_Left
5410        .cost_usd_lbl.Left = .cost_usd.Left
5420        .cost_usd_lbl_line.Left = .cost_usd_lbl.Left
5430        .Detail_vline09.Left = (.cost_usd.Left - lngVLine_Offset)
5440        .cost_usd.Visible = True
5450        .cost_usd_lbl.Visible = True
5460        .cost_usd_lbl_line.Visible = True
5470        .Detail_vline09.Visible = True
5480        .pcash_usd.Left = lngPCashUSD_Left
5490        .pcash_usd_lbl.Left = .pcash_usd.Left
5500        .pcash_usd_lbl_line.Left = .pcash_usd_lbl.Left
5510        .Detail_vline08.Left = (.pcash_usd.Left - lngVLine_Offset)
5520        .pcash_usd.Visible = True
5530        .pcash_usd_lbl.Visible = True
5540        .pcash_usd_lbl_line.Visible = True
5550        .Detail_vline08.Visible = True
5560        .icash_usd.Left = lngICashUSD_Left
5570        .icash_usd_lbl.Left = .icash_usd.Left
5580        .icash_usd_lbl_line.Left = .icash_usd_lbl.Left
5590        .Detail_vline07.Left = (.icash_usd.Left - lngVLine_Offset)
5600        .icash_usd.Visible = True
5610        .icash_usd_lbl.Visible = True
5620        .icash_usd_lbl_line.Visible = True
5630        .Detail_vline07.Visible = True
5640        .curr_id.Left = lngCurrID_Left
5650        .curr_id_lbl.Left = lngCurrIDLbl_Left
5660        .curr_id_lbl_line.Left = .curr_id.Left
5670        .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
5680        .Detail_vline06.Left = (.curr_id.Left - lngVLine_Offset)
5690        .curr_id.Visible = True
5700        .curr_id_lbl.Visible = True
5710        .curr_id_lbl_line.Visible = True
5720        .curr_id_forex_bg.Visible = True
5730        .Detail_vline06.Visible = True
5740        .IsArch_bg1.Width = .Width
5750        .IsArch_bg2.Width = .Width
5760      Case False
5770        .FocusHolder.SetFocus
5780        DoEvents
5790        If .Sort_line.Left = .cost_usd_lbl.Left Then blnSortHere = True
5800        .cost_usd.Visible = False
5810        .cost_usd_lbl.Visible = False
5820        .cost_usd_lbl_line.Visible = False
5830        .Detail_vline09.Visible = False
5840        .cost_usd.Left = 0&
5850        .cost_usd_lbl.Left = 0&
5860        .cost_usd_lbl_line.Left = 0&
5870        .Detail_vline09.Left = 0&
5880        If blnSortHere = True Then
5890          .Sort_line.Visible = False
5900          .Sort_lbl.Visible = False
5910          blnResort = True
5920        End If
5930        blnSortHere = False
5940        If .Sort_line.Left = .pcash_usd_lbl.Left Then blnSortHere = True
5950        .pcash_usd.Visible = False
5960        .pcash_usd_lbl.Visible = False
5970        .pcash_usd_lbl_line.Visible = False
5980        .Detail_vline08.Visible = False
5990        .pcash_usd.Left = 0&
6000        .pcash_usd_lbl.Left = 0&
6010        .pcash_usd_lbl_line.Left = 0&
6020        .Detail_vline08.Left = 0&
6030        If blnSortHere = True Then
6040          .Sort_line.Visible = False
6050          .Sort_lbl.Visible = False
6060          blnResort = True
6070        End If
6080        blnSortHere = False
6090        If .Sort_line.Left = .icash_usd_lbl.Left Then blnSortHere = True
6100        .icash_usd.Visible = False
6110        .icash_usd_lbl.Visible = False
6120        .icash_usd_lbl_line.Visible = False
6130        .Detail_vline07.Visible = False
6140        .icash_usd.Left = 0&
6150        .icash_usd_lbl.Left = 0&
6160        .icash_usd_lbl_line.Left = 0&
6170        .Detail_vline07.Left = 0&
6180        If blnSortHere = True Then
6190          .Sort_line.Visible = False
6200          .Sort_lbl.Visible = False
6210          blnResort = True
6220        End If
6230        blnSortHere = False
6240        If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
6250        .curr_id.Visible = False
6260        .curr_id_lbl.Visible = False
6270        .curr_id_lbl_line.Visible = False
6280        .curr_id_forex_bg.Visible = False
6290        .Detail_vline06.Visible = False
6300        .curr_id.Left = 0&
6310        .curr_id_lbl.Left = 0&
6320        .curr_id_lbl_line.Left = 0&
6330        .curr_id_forex_bg.Left = 0&
6340        .Detail_vline06.Left = 0&
6350        If blnSortHere = True Then
6360          .Sort_line.Visible = False
6370          .Sort_lbl.Visible = False
6380          blnResort = True
6390        End If
6400        blnSortHere = False
6410        .IsArch_bg1.Width = (lngForm_Width - lngTotWidth_Diff)
6420        .IsArch_bg2.Width = (lngForm_Width - lngTotWidth_Diff)
6430        .Width = (lngForm_Width - lngTotWidth_Diff)
6440      End Select
6450      DoEvents
6460      If blnResort = True Then
6470        SortNow "Form_Load"  ' ** Procedure: Below.
6480      End If
6490    End With

EXITP:
6500    Exit Sub

ERRH:
6510    Select Case ERR.Number
        Case Else
6520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6530    End Select
6540    Resume EXITP

End Sub

Public Function RecCnt() As Long

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

6610    lngRetVal = 0&

6620    With Me
6630      Set rst = .RecordsetClone
6640      With rst
6650        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
6660        Else
6670          .MoveLast
6680          lngRetVal = .RecordCount
6690        End If
6700        .Close
6710      End With
6720    End With

EXITP:
6730    Set rst = Nothing
6740    RecCnt = lngRetVal
6750    Exit Function

ERRH:
6760    lngRetVal = 0&
6770    Select Case ERR.Number
        Case Else
6780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6790    End Select
6800    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

6910    Select Case intWhere
        Case acCmdRecordsGoToFirst
6920      DoCmd.RunCommand acCmdRecordsGoToFirst
6930    Case acCmdRecordsGoToLast
6940      DoCmd.RunCommand acCmdRecordsGoToLast
6950    Case acCmdRecordsGoToPrevious
6960      DoCmd.RunCommand acCmdRecordsGoToPrevious
6970    Case acCmdRecordsGoToNext
6980      DoCmd.RunCommand acCmdRecordsGoToNext
6990    Case Else
7000      If IsMissing(varID) = False Then
7010        With Me
7020          Set rst = .RecordsetClone
7030          With rst
7040            .FindFirst "[journalno] = " & CStr(varID)
7050            If .NoMatch = False Then
7060              Me.Bookmark = .Bookmark
7070            End If
7080            .Close
7090          End With
7100        End With
7110      End If
7120    End Select

EXITP:
7130    Set rst = Nothing
7140    Exit Sub

ERRH:
7150    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7160    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7170    Case Else
7180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7190    End Select
7200    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [JournalType_Order], [journalno]"

7310    With Me
7320      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7330        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7340      End If
7350      .Controls(strSortLbl).Visible = False
7360      .Controls(strSortLine).Visible = False
7370      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
7380      If strProc = "Form_Load" Then
7390        strCalled = "transdate"
7400        strSortNow = strSortOrig
7410        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7420        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7430        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7440        .Controls(strSortLbl).Left = lngSortLbl_Left
7450        .Controls(strSortLine).Left = lngSortLine_Left
7460        .Controls(strSortLine).Width = lngSortLine_Width
7470        .Controls(strSortLbl).Caption = strArwDn
7480        .Controls(strSortLbl).ForeColor = CLR_DKRED
7490      Else
7500        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
7510        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7520        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7530        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7540        .Controls(strSortLbl).Left = lngSortLbl_Left
7550        .Controls(strSortLine).Left = lngSortLine_Left
7560        .Controls(strSortLine).Width = lngSortLine_Width
7570        If strCalled = "transdate" Then
7580          strSortAsc = strSortOrig
7590          If strSortNow = strSortAsc Then
7600            strSortNow = "[transdate], [JournalType_Order], [journalno]"
7610            .Controls(strSortLbl).Caption = strArwUp
7620            .Controls(strSortLbl).ForeColor = CLR_DKBLU
7630          Else
7640            strSortNow = strSortAsc
7650            .Controls(strSortLbl).Caption = strArwDn
7660            .Controls(strSortLbl).ForeColor = CLR_DKRED
7670          End If
7680        Else
7690          .Controls(strSortLbl).Caption = strArwUp
7700          .Controls(strSortLbl).ForeColor = CLR_DKBLU
7710          Select Case strCalled
              Case "journaltype"
7720            strSortAsc = "[JournalType_Order], [transdate], [journalno]"
7730          Case "totdesc"
7740            strSortAsc = "[totdesc]" & strStdAsc
7750          Case "icash"
7760            strSortAsc = "[icash]" & strStdAsc
7770          Case "pcash"
7780            strSortAsc = "[pcash]" & strStdAsc
7790          Case "cost"
7800            strSortAsc = "[cost]" & strStdAsc
7810          Case "curr_id"
7820            strSortAsc = "[curr_code]" & strStdAsc
7830            lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
7840            lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
7850            lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
7860            .Controls(strSortLbl).Left = lngSortLbl_Left
7870            .Controls(strSortLine).Left = lngSortLine_Left
7880            .Controls(strSortLine).Width = lngSortLine_Width
7890          Case "icash_usd"
7900            strSortAsc = "[icash_usd]" & strStdAsc
7910            lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
7920            .Controls(strSortLbl).Left = lngSortLbl_Left
7930          Case "pcash_usd"
7940            strSortAsc = "[pcash_usd]" & strStdAsc
7950            lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
7960            .Controls(strSortLbl).Left = lngSortLbl_Left
7970          Case "cost_usd"
7980            strSortAsc = "[cost_usd]" & strStdAsc
7990          End Select
8000          If strSortNow = strSortAsc Then
8010            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
8020            Select Case intCnt
                Case 1
8030              strTmp01 = strSortAsc & " DESC"
8040            Case 2
8050              intPos01 = InStr(strSortAsc, ",")
8060              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8070              strTmp01 = Mid(strSortAsc, intPos01) & " DESC"
8080            Case 3
8090              intPos01 = InStr(strSortAsc, ",")
8100              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8110              strTmp02 = Mid(strSortAsc, intPos01)
8120              intPos01 = InStr(2, strTmp02, ",")
8130              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8140              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
8150            Case 4
8160              intPos01 = InStr(strSortAsc, ",")
8170              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8180              strTmp02 = Mid(strSortAsc, intPos01)
8190              intPos01 = InStr(2, strTmp02, ",")
8200              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8210              strTmp02 = Mid(strTmp02, intPos01)
8220              intPos01 = InStr(2, strTmp02, ",")
8230              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8240              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
8250            End Select
8260            strSortNow = strTmp01
8270            .Controls(strSortLbl).Caption = strArwDn
8280            .Controls(strSortLbl).ForeColor = CLR_DKRED
8290          Else
8300            strSortNow = strSortAsc
8310          End If
8320        End If
8330      End If
8340      .Controls(strSortLbl).Visible = True
8350      .Controls(strSortLine).Visible = True
8360      .OrderBy = strSortNow
8370      .OrderByOn = True
8380    End With

EXITP:
8390    Exit Sub

ERRH:
8400    Select Case ERR.Number
        Case Else
8410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8420    End Select
8430    Resume EXITP

End Sub
