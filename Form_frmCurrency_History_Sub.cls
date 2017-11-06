VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_History_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_History_Sub"

'VGC 12/30/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Active Only Date: A {chkActiveOnly_Dates on frmCurrency_History}
' **   Close:            C {cmdClose on frmCurrency_History}
' **   Prior Dates:      D {lbxHistoryDate on frmCurrency_History}
' **   Currency Code:    O {curr_code}
' **   Print Report:     P {cmdPrintReport on frmCurrency_History}
' **   Prior Rates:      R {lbxHistoryRate on frmCurrency_History}
' **   Active Only Rate: V {chkActiveOnly_Rates on frmCurrency_History}
' **   Exit:             X {cmdClose on frmCurrency_History}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[curr_date] DESC, [curr_code]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmCurrency_History" Then

160         blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180         lngSortLbl_Width = .Sort_lbl.Width

190       Else
200         Cancel = -1
210       End If
220     End With

230     If Cancel = -1 Then
240       Beep
250       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
260     End If

EXITP:
270     Exit Sub

ERRH:
280     Select Case ERR.Number
        Case Else
290       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
300     End Select
310     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me

420       CLR_DISABLED_FG = CLR_DKGRY
430       CLR_DISABLED_BG = CLR_LTTEAL

440       .curr_date.ForeColor = CLR_DISABLED_FG
450       .curr_date.BackColor = CLR_DISABLED_BG
460       .curr_code.ForeColor = CLR_DISABLED_FG
470       .curr_code.BackColor = CLR_DISABLED_BG
480       .curr_name.ForeColor = CLR_DISABLED_FG
490       .curr_name.BackColor = CLR_DISABLED_BG
500       .curr_rate1.ForeColor = CLR_DISABLED_FG
510       .curr_rate1.BackColor = CLR_DISABLED_BG
520       .curr_rate2.ForeColor = CLR_DISABLED_FG
530       .curr_rate2.BackColor = CLR_DISABLED_BG
540       .currsym_symbol_display.ForeColor = CLR_DISABLED_FG
550       .currsym_symbol_display.BackColor = CLR_DISABLED_BG

560       SortNow THIS_PROC  ' ** Procedure: Below.

570     End With

EXITP:
580     Exit Sub

ERRH:
590     Select Case ERR.Number
        Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Prior Dates:      D {lbxHistoryDate on frmCurrency_History}
        ' **   Currency Code:    O {curr_code)
        ' **   Prior Rates:      R {lbxHistoryRate on frmCurrency_History}
        ' **   Exit:             X {cmdClose on frmCurrency_History}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         With Me
780           intRetVal = 0
790           .Parent.cmdClose_Click  ' ** Form Procedure: frmCurrency_History.
800         End With
810       Case vbKeyUp
820         With Me
830           intRetVal = 0
840           MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
850         End With
860       Case vbKeyDown
870         With Me
880           intRetVal = 0
890           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
900         End With
910       End Select
920     End If

        ' ** Alt keys.
930     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
940       Select Case intRetVal
          Case vbKeyD
950         With Me
960           intRetVal = 0
970           DoCmd.SelectObject acForm, .Parent.Name, False
980           .Parent.lbxHistoryDate.SetFocus
990         End With
1000      Case vbKeyO
1010        With Me
1020          intRetVal = 0
1030          .curr_code.SetFocus
1040        End With
1050      Case vbKeyR
1060        With Me
1070          intRetVal = 0
1080          DoCmd.SelectObject acForm, .Parent.Name, False
1090          .Parent.lbxHistoryRate.SetFocus
1100        End With
1110      Case vbKeyX
1120        With Me
1130          intRetVal = 0
1140          .Parent.cmdClose_Click  ' ** Form Procedure: frmCurrency_History.
1150        End With
1160      End Select
1170    End If

        ' ** Ctrl keys.
1180    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1190      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1200        With Me
1210          intRetVal = 0
1220          DoCmd.SelectObject acForm, .Parent.Name, False
1230          If .Parent.cmdPrintReport.Enabled = True Then
1240            .Parent.cmdPrintReport.SetFocus
1250          ElseIf .Parent.cmdRestore.Enabled = True Then
1260            .Parent.cmdRestore.SetFocus
1270          ElseIf .Parent.cmdEdit.Enabled = True Then
1280            .Parent.cmdEdit.SetFocus
1290          Else
1300            .Parent.cmdClose.SetFocus
1310          End If
1320        End With
1330      End Select
1340    End If

        ' ** Ctrl-Shift keys.
1350    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1370        With Me
1380          intRetVal = 0
1390          DoCmd.SelectObject acForm, .Parent.Name, False
1400          .Parent.chkActiveOnly_Rates.SetFocus
1410        End With
1420      End Select
1430    End If

EXITP:
1440    KeyCode = intRetVal
1450    Exit Sub

ERRH:
1460    intRetVal = 0
1470    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1480    Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_MouseMove"

        Dim ctl As Access.Control
        Dim strControl As String

1610    With Me
1620  On Error Resume Next
1630      strControl = Screen.ActiveControl.Name
1640  On Error GoTo ERRH
1650      If strControl <> vbNullString Then
1660        DoCmd.SelectObject acForm, .Parent.Name, False
1670        .Parent.frmCurrency_History_Sub.SetFocus
1680      End If
1690    End With

EXITP:
1700    Set ctl = Nothing
1710    Exit Sub

ERRH:
1720    Select Case ERR.Number
        Case Else
1730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1740    End Select
1750    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

1810    Form_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1910    Form_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
1920    Exit Sub

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub curr_name_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_MouseMove"

2010    Form_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2170        With Me
2180          intRetVal = 0
2190          .curr_date.SetFocus
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2250        With Me
2260          intRetVal = 0
2270          If .CurrentRecord > 1 Then
2280            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2290            .curr_active.SetFocus
2300          Else
2310            DoCmd.SelectObject acForm, .Parent.Name, False
2320            .Parent.chkActiveOnly_Rates.SetFocus
2330          End If
2340        End With
2350      End Select
2360    End If

EXITP:
2370    KeyCode = intRetVal
2380    Exit Sub

ERRH:
2390    intRetVal = 0
2400    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2410    Case Else
2420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2430    End Select
2440    Resume EXITP

End Sub

Private Sub curr_date_lbl_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_lbl_DblClick"

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

Private Sub curr_date_GotFocus()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_GotFocus"

2610    With Me
2620      .curr_date.SelLength = 0
2630      .curr_date.SelStart = 19
2640    End With

EXITP:
2650    Exit Sub

ERRH:
2660    Select Case ERR.Number
        Case Else
2670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2680    End Select
2690    Resume EXITP

End Sub

Private Sub curr_date_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2770        With Me
2780          intRetVal = 0
2790          .curr_code.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2850        With Me
2860          intRetVal = 0
2870          If .CurrentRecord > 1 Then
2880            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2890            .curr_active.SetFocus
2900          Else
2910            DoCmd.SelectObject acForm, .Parent.Name, False
2920            .Parent.chkActiveOnly_Rates.SetFocus
2930          End If
2940        End With
2950      End Select
2960    End If

EXITP:
2970    KeyCode = intRetVal
2980    Exit Sub

ERRH:
2990    intRetVal = 0
3000    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3010    Case Else
3020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3030    End Select
3040    Resume EXITP

End Sub

Private Sub curr_code_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_lbl_DblClick"

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

Private Sub curr_code_GotFocus()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_GotFocus"

3210    With Me
3220      .curr_code.SelLength = 0
3230      .curr_code.SelStart = 9
3240    End With

EXITP:
3250    Exit Sub

ERRH:
3260    Select Case ERR.Number
        Case Else
3270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3280    End Select
3290    Resume EXITP

End Sub

Private Sub curr_name_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_lbl_DblClick"

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

Private Sub curr_name_GotFocus()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_GotFocus"

3410    With Me
3420      .curr_name.SelLength = 0
3430      .curr_name.SelStart = 99
3440    End With

EXITP:
3450    Exit Sub

ERRH:
3460    Select Case ERR.Number
        Case Else
3470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3480    End Select
3490    Resume EXITP

End Sub

Private Sub curr_rate1_lbl_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_lbl_DblClick"

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

Private Sub curr_rate1_GotFocus()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_GotFocus"

3610    With Me
3620      .curr_rate1.SelLength = 0
3630      .curr_rate1.SelStart = 19
3640    End With

EXITP:
3650    Exit Sub

ERRH:
3660    Select Case ERR.Number
        Case Else
3670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3680    End Select
3690    Resume EXITP

End Sub

Private Sub curr_rate2_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_lbl_DblClick"

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

Private Sub curr_rate2_GotFocus()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_GotFocus"

3810    With Me
3820      .curr_rate2.SelLength = 0
3830      .curr_rate2.SelStart = 19
3840    End With

EXITP:
3850    Exit Sub

ERRH:
3860    Select Case ERR.Number
        Case Else
3870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3880    End Select
3890    Resume EXITP

End Sub

Private Sub currsym_symbol_display_lbl_DblClick(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "currsym_symbol_display_lbl_DblClick"

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

Private Sub currsym_symbol_display_GotFocus()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "currsym_symbol_display_GotFocus"

4010    With Me
4020      If IsNull(.currsym_symbol_display) = False Then
4030        .currsym_symbol_display.SelLength = 0
4040        .currsym_symbol_display.SelStart = 9
4050      Else
4060        .curr_active.SetFocus
4070      End If
4080    End With

EXITP:
4090    Exit Sub

ERRH:
4100    Select Case ERR.Number
        Case Else
4110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4120    End Select
4130    Resume EXITP

End Sub

Private Sub curr_active_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_active_lbl_DblClick"

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

Private Sub curr_active_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_active_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4320    intShiftDown = (Shift And acShiftMask) > 0
4330    intAltDown = (Shift And acAltMask) > 0
4340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4370        With Me
4380          intRetVal = 0
4390          lngRecsCur = RecCnt  ' ** Function: Below.
4400          If .CurrentRecord < lngRecsCur Then
4410            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4420            .curr_date.SetFocus
4430          Else
4440            DoCmd.SelectObject acForm, .Parent.Name, False
4450            If .Parent.cmdPrintReport.Enabled = True Then
4460              .Parent.cmdPrintReport.SetFocus
4470            ElseIf .Parent.cmdRestore.Enabled = True Then
4480              .Parent.cmdRestore.SetFocus
4490            ElseIf .Parent.cmdEdit.Enabled = True Then
4500              .Parent.cmdEdit.SetFocus
4510            Else
4520              .Parent.cmdClose.SetFocus
4530            End If
4540          End If
4550        End With
4560      End Select
4570    End If

        ' ** Shift keys.
4580    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4590      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4600        With Me
4610          intRetVal = 0
4620          If IsNull(.currsym_symbol_display) = False Then
4630            .currsym_symbol_display.SetFocus
4640          Else
4650            .curr_rate2.SetFocus
4660          End If
4670        End With
4680      End Select
4690    End If

EXITP:
4700    KeyCode = intRetVal
4710    Exit Sub

ERRH:
4720    intRetVal = 0
4730    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4740    Case Else
4750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4760    End Select
4770    Resume EXITP

End Sub

Public Function RecCnt() As Long

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

4810    lngRetVal = 0&

4820    With Me
4830      Set rst = .RecordsetClone
4840      With rst
4850        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
4860        Else
4870          .MoveLast
4880          lngRetVal = .RecordCount
4890        End If
4900        .Close
4910      End With
4920    End With

EXITP:
4930    Set rst = Nothing
4940    RecCnt = lngRetVal
4950    Exit Function

ERRH:
4960    lngRetVal = 0&
4970    Select Case ERR.Number
        Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

5110    Select Case intWhere
        Case acCmdRecordsGoToFirst
5120      DoCmd.RunCommand acCmdRecordsGoToFirst
5130    Case acCmdRecordsGoToLast
5140      DoCmd.RunCommand acCmdRecordsGoToLast
5150    Case acCmdRecordsGoToPrevious
5160      DoCmd.RunCommand acCmdRecordsGoToPrevious
5170    Case acCmdRecordsGoToNext
5180      DoCmd.RunCommand acCmdRecordsGoToNext
5190    Case Else
5200      If IsMissing(varID) = False Then
5210        With Me
5220          Set rst = .RecordsetClone
5230          With rst
5240            .FindFirst "[currhist_id] = '" & varID & "'"
5250            If .NoMatch = False Then
5260              Me.Bookmark = .Bookmark
5270            End If
5280            .Close
5290          End With
5300        End With
5310      End If
5320    End Select

EXITP:
5330    Set rst = Nothing
5340    Exit Sub

ERRH:
5350    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5360    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5370    Case Else
5380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5390    End Select
5400    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [curr_date] DESC, [curr_code]"

5510    With Me
5520      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5530        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5540      End If
5550      .Controls(strSortLbl).Visible = False
5560      .Controls(strSortLine).Visible = False
5570      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
5580      If strProc = "Form_Load" Then
5590        strCalled = "curr_date"
5600        strSortNow = strSortOrig
5610        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5620        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5630        lngSortLbl_Left = (lngSortLbl_Left + lngTpp)
5640        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5650        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5660        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5670        .Controls(strSortLbl).Top = lngSortLbl_Top
5680        .Controls(strSortLbl).Left = lngSortLbl_Left
5690        .Controls(strSortLine).Top = lngSortLine_Top
5700        .Controls(strSortLine).Left = lngSortLine_Left
5710        .Controls(strSortLine).Width = lngSortLine_Width
5720        .Controls(strSortLbl).Caption = strArwDn
5730        .Controls(strSortLbl).ForeColor = CLR_DKRED
5740      Else
5750        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))
5760        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5770        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5780        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5790        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5800        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5810        .Controls(strSortLbl).Top = lngSortLbl_Top
5820        .Controls(strSortLbl).Left = lngSortLbl_Left
5830        .Controls(strSortLine).Top = lngSortLine_Top
5840        .Controls(strSortLine).Left = lngSortLine_Left
5850        .Controls(strSortLine).Width = lngSortLine_Width
5860        If strCalled = "curr_date" Then
5870          strSortAsc = strSortOrig
5880          If strSortNow = strSortAsc Then
5890            strSortNow = "[curr_date], [curr_code]"
5900            lngSortLbl_Left = (lngSortLbl_Left + lngTpp)
5910            .Controls(strSortLbl).Left = lngSortLbl_Left
5920            .Controls(strSortLbl).Caption = strArwUp
5930            .Controls(strSortLbl).ForeColor = CLR_DKBLU
5940          Else
5950            strSortNow = strSortOrig
5960            .Controls(strSortLbl).Caption = strArwDn
5970            .Controls(strSortLbl).ForeColor = CLR_DKRED
5980          End If
5990        Else
6000          .Controls(strSortLbl).Caption = strArwUp
6010          .Controls(strSortLbl).ForeColor = CLR_DKBLU
6020          Select Case strCalled
              Case "curr_code"
6030            strSortAsc = "[curr_code]" & strStdAsc
6040            lngSortLbl_Left = (lngSortLbl_Left + (9& * lngTpp))
6050            .Controls(strSortLbl).Left = lngSortLbl_Left
6060          Case "curr_name"
6070            strSortAsc = "[curr_name]" & strStdAsc
6080          Case "curr_rate1"
6090            strSortAsc = "[curr_rate1]" & strStdAsc
6100          Case "curr_rate2"
6110            strSortAsc = "[curr_rate2]" & strStdAsc
6120          Case "currsym_symbol_display"
6130            strSortAsc = "[currsym_symbol]" & strStdAsc
6140            lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
6150            .Controls(strSortLbl).Left = lngSortLbl_Left
6160          Case "curr_active"
6170            strSortAsc = "[curr_active]" & strStdAsc
6180            lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
6190            .Controls(strSortLbl).Left = lngSortLbl_Left
6200          End Select
6210          If strSortNow = strSortAsc Then
6220            intCnt = (CharCnt(strSortAsc, ",") + 1)  ' ** Module Function: modStringFuncs.
6230            Select Case intCnt
                Case 1
6240              strTmp01 = strSortAsc & " DESC"
6250            Case 2
6260              intPos01 = InStr(strSortAsc, ",")
6270              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
6280            Case 3
6290              intPos01 = InStr(strSortAsc, ",")
6300              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6310              strTmp02 = Mid(strSortAsc, intPos01)
6320              intPos01 = InStr(2, strTmp02, ",")
6330              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
6340              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
6350            End Select
6360            strSortNow = strTmp01
6370            .Controls(strSortLbl).Caption = strArwDn
6380            .Controls(strSortLbl).ForeColor = CLR_DKRED
6390          Else
6400            strSortNow = strSortAsc
6410          End If
6420        End If
6430      End If
6440      .Controls(strSortLbl).Visible = True
6450      .Controls(strSortLine).Visible = True
6460      .OrderBy = strSortNow
6470      .OrderByOn = True
6480    End With

EXITP:
6490    Exit Sub

ERRH:
6500    Select Case ERR.Number
        Case Else
6510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6520    End Select
6530    Resume EXITP

End Sub

Public Function SortNow_Get() As String

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

6610    strRetVal = strSortNow

EXITP:
6620    SortNow_Get = strRetVal
6630    Exit Function

ERRH:
6640    strRetVal = vbNullString
6650    Select Case ERR.Number
        Case Else
6660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6670    End Select
6680    Resume EXITP

End Function
