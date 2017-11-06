VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_Symbol_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_Symbol_Sub"

'VGC 04/17/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmCurrency_Symbol}
' **   Active Exclude:   E {opgActive_optExclude on frmCurrency_Symbol}
' **   Active Include:   I {opgActive_optInclude on frmCurrency_Symbol}
' **   Country:          N {country_name1}
' **   Active Only:      O {opgActive_optOnly on frmCurrency_Symbol}
' **   Print Report:     P {cmdPrintReport on frmCurrency_Symbol}
' **   Currency:         R {curr_word1}
' **   Exit:             X {cmdClose on frmCurrency_Symbol}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[country_name_sort]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngTpp As Long
Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmCurrency_Symbol" Then

160         lngSortLbl_Width = .Sort_lbl.Width

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200       Else
210         Cancel = -1
220       End If
230     End With

240     If Cancel = -1 Then
250       Beep
260       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
270     End If

EXITP:
280     Exit Sub

ERRH:
290     Select Case ERR.Number
        Case Else
300       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
310     End Select
320     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .curr_code.ForeColor = CLR_DISABLED_FG
430       .curr_code.BackColor = CLR_DISABLED_BG
440       .curr_word1.ForeColor = CLR_DISABLED_FG
450       .curr_word1.BackColor = CLR_DISABLED_BG
460       .curr_word2.ForeColor = CLR_DISABLED_FG
470       .curr_word2_box.BackColor = CLR_DISABLED_BG
480       .country_code3.ForeColor = CLR_DISABLED_FG
490       .country_code3.BackColor = CLR_DISABLED_BG
500       .country_name1.ForeColor = CLR_DISABLED_FG
510       .country_name1.BackColor = CLR_DISABLED_BG
520       .country_name2.ForeColor = CLR_DISABLED_FG
530       .country_name2_box.BackColor = CLR_DISABLED_BG
540     End With

550     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
560     Exit Sub

ERRH:
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

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
        ' **   Name:             N {country_name1}
        ' **   Currency:         R {curr_word1}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyUp
770         intRetVal = 0
780         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
790       Case vbKeyDown
800         intRetVal = 0
810         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
820       Case vbKeyEscape
830         With Me
840           intRetVal = 0
850           .Parent.cmdClose_Click  ' ** Form Procedure: frmCountryCode.
860         End With
870       End Select
880     End If

        ' ** Alt keys.
890     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
900       Select Case intRetVal
          Case vbKeyN
910         With Me
920           intRetVal = 0
930           .country_name1.SetFocus
940         End With
950       Case vbKeyR
960         With Me
970           intRetVal = 0
980           .curr_word1.SetFocus
990         End With
1000      Case vbKeyX
1010        With Me
1020          intRetVal = 0
1030          .Parent.cmdClose_Click  ' ** Form Procedure: frmCountryCode.
1040        End With
1050      End Select
1060    End If

        ' ** Ctrl keys.
1070    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1080      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1090        With Me
1100          intRetVal = 0
1110          DoCmd.SelectObject acForm, .Parent.Name, False
1120          .Parent.opgActive.SetFocus
1130        End With
1140      Case vbKeyUp, vbKeyPageUp
1150        intRetVal = 0
1160        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1170      Case vbKeyDown, vbKeyPageDown
1180        intRetVal = 0
1190        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1200      End Select
1210    End If

        ' ** Ctrl-Shift keys.
1220    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1230      Select Case intRetVal
          Case vbKeyF
1240        With Me
1250          intRetVal = 0
1260          DoCmd.SelectObject acForm, .Parent.Name, False
1270          .Parent.FocusHolder.SetFocus
1280        End With
1290      Case vbKeyTab, vbKeyReturn
1300        With Me
1310          intRetVal = 0
1320          DoCmd.SelectObject acForm, .Parent.Name, False
1330          .Parent.cmdClose.SetFocus
1340        End With
1350      End Select
1360    End If

EXITP:
1370    KeyCode = intRetVal
1380    Exit Sub

ERRH:
1390    intRetVal = 0
1400    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1410    Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1520    intShiftDown = (Shift And acShiftMask) > 0
1530    intAltDown = (Shift And acAltMask) > 0
1540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1570        With Me
1580          intRetVal = 0
1590          lngRecsCur = RecCnt  ' ** Function: Below.
1600          If .CurrentRecord < lngRecsCur Then
1610            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1620            .FocusHolder.SetFocus
1630          Else
1640            DoCmd.SelectObject acForm, .Parent.Name, False
1650            .Parent.opgActive.SetFocus
1660          End If
1670        End With
1680      End Select
1690    End If

        ' ** Shift keys.
1700    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1710      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1720        With Me
1730          intRetVal = 0
1740          If .CurrentRecord > 1 Then
1750            MoveRec acCmdRecordsGoToPrevious
1760            .FocusHolder.SetFocus
1770          Else
1780            DoCmd.SelectObject acForm, .Parent.Name, False
1790            .Parent.cmdClose.SetFocus
1800          End If
1810        End With
1820      End Select
1830    End If

EXITP:
1840    KeyCode = intRetVal
1850    Exit Sub

ERRH:
1860    intRetVal = 0
1870    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1880    Case Else
1890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1900    End Select
1910    Resume EXITP

End Sub

Private Sub curr_code_lbl_DblClick(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_lbl_DblClick"

2010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub curr_code_GotFocus()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_GotFocus"

2110    With Me
2120      .curr_code.SelLength = 0
2130      .curr_code.SelStart = 9
2140    End With

EXITP:
2150    Exit Sub

ERRH:
2160    Select Case ERR.Number
        Case Else
2170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2180    End Select
2190    Resume EXITP

End Sub

Private Sub curr_code_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2270        With Me
2280          intRetVal = 0
2290          .currsym_symbol_display.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2350        With Me
2360          intRetVal = 0
2370          If .CurrentRecord > 1 Then
2380            MoveRec acCmdRecordsGoToPrevious
2390            .curr_code.SetFocus
2400          Else
2410            DoCmd.SelectObject acForm, .Parent.Name, False
2420            .Parent.cmdClose.SetFocus
2430          End If
2440        End With
2450      End Select
2460    End If

EXITP:
2470    KeyCode = intRetVal
2480    Exit Sub

ERRH:
2490    intRetVal = 0
2500    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2510    Case Else
2520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2530    End Select
2540    Resume EXITP

End Sub

Private Sub currsym_symbol_display_lbl_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "currsym_symbol_display_lbl_DblClick"

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

Private Sub currsym_symbol_display_GotFocus()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "currsym_symbol_display_GotFocus"

2710    With Me
2720      .currsym_symbol_display.SelLength = 0
2730      .currsym_symbol_display.SelStart = 9
2740    End With

EXITP:
2750    Exit Sub

ERRH:
2760    Select Case ERR.Number
        Case Else
2770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2780    End Select
2790    Resume EXITP

End Sub

Private Sub curr_word1_lbl_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_word1_lbl_DblClick"

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

Private Sub curr_word1_GotFocus()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_word1_GotFocus"

2910    With Me
2920      .curr_word1.SelLength = 0
2930      .curr_word1.SelStart = 99
2940    End With

EXITP:
2950    Exit Sub

ERRH:
2960    Select Case ERR.Number
        Case Else
2970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2980    End Select
2990    Resume EXITP

End Sub

Private Sub country_code3_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "country_code3_lbl_DblClick"

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

Private Sub country_code3_GotFocus()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "country_code3_GotFocus"

3110    With Me
3120      .country_code3.SelLength = 0
3130      .country_code3.SelStart = 9
3140    End With

EXITP:
3150    Exit Sub

ERRH:
3160    Select Case ERR.Number
        Case Else
3170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3180    End Select
3190    Resume EXITP

End Sub

Private Sub country_name1_lbl_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "country_name1_lbl_DblClick"

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

Private Sub country_name1_GotFocus()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "country_name1_GotFocus"

3310    With Me
3320      .country_name1.SelLength = 0
3330      .country_name1.SelStart = 99
3340    End With

EXITP:
3350    Exit Sub

ERRH:
3360    Select Case ERR.Number
        Case Else
3370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3380    End Select
3390    Resume EXITP

End Sub

Private Sub country_name1_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "country_name1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3420    intShiftDown = (Shift And acShiftMask) > 0
3430    intAltDown = (Shift And acAltMask) > 0
3440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3470        With Me
3480          intRetVal = 0
3490          lngRecsCur = RecCnt  ' ** Function: Below
3500          If .CurrentRecord < lngRecsCur Then
3510            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3520            .FocusHolder.SetFocus
3530          Else
3540            DoCmd.SelectObject acForm, .Parent.Name, False
3550            .Parent.opgActive.SetFocus
3560          End If
3570        End With
3580      End Select
3590    End If

        ' ** Shift keys.
3600    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3610      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3620        With Me
3630          intRetVal = 0
3640          .country_code3.SetFocus
3650        End With
3660      End Select
3670    End If

EXITP:
3680    KeyCode = intRetVal
3690    Exit Sub

ERRH:
3700    intRetVal = 0
3710    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3720    Case Else
3730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3740    End Select
3750    Resume EXITP

End Sub

Public Function RecCnt() As Long

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

3810    lngRetVal = 0&

3820    With Me
3830      Set rst = .RecordsetClone
3840      With rst
3850        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
3860        Else
3870          .MoveLast
3880          lngRetVal = .RecordCount
3890        End If
3900        .Close
3910      End With
3920    End With

EXITP:
3930    Set rst = Nothing
3940    RecCnt = lngRetVal
3950    Exit Function

ERRH:
3960    lngRetVal = 0&
3970    Select Case ERR.Number
        Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

4110    Select Case intWhere
        Case acCmdRecordsGoToFirst
4120      DoCmd.RunCommand acCmdRecordsGoToFirst
4130    Case acCmdRecordsGoToLast
4140      DoCmd.RunCommand acCmdRecordsGoToLast
4150    Case acCmdRecordsGoToPrevious
4160      DoCmd.RunCommand acCmdRecordsGoToPrevious
4170    Case acCmdRecordsGoToNext
4180      DoCmd.RunCommand acCmdRecordsGoToNext
4190    Case Else
4200      If IsMissing(varID) = False Then
4210        With Me
4220          Set rst = .RecordsetClone
4230          With rst
4240            .FindFirst "[currsym_id] = '" & varID & "'"
4250            If .NoMatch = False Then
4260              Me.Bookmark = .Bookmark
4270            End If
4280            .Close
4290          End With
4300        End With
4310      End If
4320    End Select

EXITP:
4330    Set rst = Nothing
4340    Exit Sub

ERRH:
4350    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4360    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4370    Case Else
4380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4390    End Select
4400    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim ctl As Access.Control
        Dim strCalled As String, strSortAsc As String

        Const strStdAsc As String = ", [country_name_sort]"

4510    With Me
4520      .Controls(strSortLbl).Visible = False
4530      .Controls(strSortLine).Visible = False
4540      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
4550      If strProc = "Form_Load" Then
4560        strCalled = "country_name1"
4570        strSortNow = strSortOrig
4580        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4590        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
4600        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4610        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4620        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4630        .Controls(strSortLbl).Top = lngSortLbl_Top
4640        .Controls(strSortLbl).Left = lngSortLbl_Left
4650        .Controls(strSortLbl).Caption = strArwUp
4660        .Controls(strSortLbl).ForeColor = CLR_DKBLU
4670        .Controls(strSortLine).Top = lngSortLine_Top
4680        .Controls(strSortLine).Left = lngSortLine_Left
4690        .Controls(strSortLine).Width = lngSortLine_Width
4700      Else
4710        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
4720        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4730        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
4740        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4750        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4760        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4770        .Controls(strSortLbl).Top = lngSortLbl_Top
4780        .Controls(strSortLbl).Left = lngSortLbl_Left
4790        .Controls(strSortLine).Top = lngSortLine_Top
4800        .Controls(strSortLine).Left = lngSortLine_Left
4810        .Controls(strSortLine).Width = lngSortLine_Width
4820        If strCalled = "country_name1" Then
4830          strSortAsc = strSortOrig
4840          If strSortNow = strSortAsc Then
4850            strSortNow = "[country_name] DESC"
4860            .Controls(strSortLbl).Caption = strArwDn
4870            .Controls(strSortLbl).ForeColor = CLR_DKRED
4880          Else
4890            strSortNow = strSortAsc
4900            .Controls(strSortLbl).Caption = strArwUp
4910            .Controls(strSortLbl).ForeColor = CLR_DKBLU
4920          End If
4930        Else
4940          .Controls(strSortLbl).Caption = strArwUp
4950          .Controls(strSortLbl).ForeColor = CLR_DKBLU
4960          Select Case strCalled
              Case "curr_code"
4970            strSortAsc = "[curr_code]"
4980            lngSortLbl_Left = (lngSortLbl_Left + (9& * lngTpp))
4990            .Controls(strSortLbl).Left = lngSortLbl_Left
5000          Case "currsym_symbol_display"
5010            strSortAsc = "[currsym_symbol]" & strStdAsc
5020            lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
5030            .Controls(strSortLbl).Left = lngSortLbl_Left
5040          Case "curr_word1"
5050            strSortAsc = "[curr_word1]" & strStdAsc
5060          Case "country_code3"
5070            strSortAsc = "[country_code3]"
5080            lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
5090            .Controls(strSortLbl).Left = lngSortLbl_Left
5100          Case "country_name1"
5110            strSortAsc = "[country_name]"
5120          End Select
5130          If strSortNow = strSortAsc Then
5140            If InStr(strSortAsc, ",") > 0 Then
5150              strSortNow = Left(strSortAsc, (InStr(strSortAsc, ",") - 1)) & " DESC" & Mid(strSortAsc, InStr(strSortAsc, ",")) & " DESC"
5160            Else
5170              strSortNow = strSortNow & " DESC"
5180            End If
5190            .Controls(strSortLbl).Caption = strArwDn
5200            .Controls(strSortLbl).ForeColor = CLR_DKRED
5210          Else
5220            strSortNow = strSortAsc
5230          End If
5240        End If
5250      End If
5260      .Controls(strSortLbl).Visible = True
5270      .Controls(strSortLine).Visible = True
5280      .OrderBy = strSortNow
5290      .OrderByOn = True
5300    End With

EXITP:
5310    Set ctl = Nothing
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Public Function SortNow_Get() As String

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

5410    strRetVal = strSortNow

EXITP:
5420    SortNow_Get = strRetVal
5430    Exit Function

ERRH:
5440    strRetVal = vbNullString
5450    Select Case ERR.Number
        Case Else
5460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5470    End Select
5480    Resume EXITP

End Function
