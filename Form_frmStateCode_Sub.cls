VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStateCode_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmStateCode_Sub"

'VGC 04/19/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmStateCode}
' **   Military:         M {chkIncludeMilitary on frmStateCode}
' **   Canada:           N {chkIncludeCanada on frmStateCode}
' **   Print Report:     P {cmdPrintReport on frmStateCode}
' **   Territories:      T {chkIncludeTerritories on frmStateCode}
' **   Exit:             X {cmdClose on frmStateCode}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[state_military] DESC, [state_canada] DESC, [state_territory] DESC, [state_multi], [state_name]"
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

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmStateCode" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         lngSortLbl_Width = .Sort_lbl.Width

180         CLR_DISABLED_FG = CLR_DKGRY
190         CLR_DISABLED_BG = CLR_LTTEAL

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
420       .state_code.ForeColor = CLR_DISABLED_FG
430       .state_code.BackColor = CLR_DISABLED_BG
440       .state_name.ForeColor = CLR_DISABLED_FG
450       .state_name.BackColor = CLR_DISABLED_BG
460     End With

470     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
480     Exit Sub

ERRH:
490     Select Case ERR.Number
        Case Else
500       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
510     End Select
520     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyUp
670         intRetVal = 0
680         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
690       Case vbKeyDown
700         intRetVal = 0
710         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
720       Case vbKeyEscape
730         With Me
740           intRetVal = 0
750           .Parent.cmdClose_Click  ' ** Form Procedure: frmStateCode.
760         End With
770       End Select
780     End If

        ' ** Alt keys.
790     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
800       Select Case intRetVal
          Case vbKeyX
810         With Me
820           intRetVal = 0
830           .Parent.cmdClose_Click  ' ** Form Procedure: frmStateCode.
840         End With
850       End Select
860     End If

        ' ** Ctrl keys.
870     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
880       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
890         With Me
900           intRetVal = 0
910           DoCmd.SelectObject acForm, .Parent.Name, False
920           If .Parent.cmdPrintReport.Enabled = True Then
930             .Parent.cmdPrintReport.SetFocus
940           Else
950             .Parent.chkIncludeTerritories.SetFocus
960           End If
970         End With
980       End Select
990     End If

        ' ** Ctrl-Shift keys.
1000    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1010      Select Case intRetVal
          Case vbKeyF
1020        With Me
1030          intRetVal = 0
1040          DoCmd.SelectObject acForm, .Parent.Name, False
1050          .Parent.FocusHolder.SetFocus
1060        End With
1070      Case vbKeyTab, vbKeyReturn
1080        With Me
1090          intRetVal = 0
1100          DoCmd.SelectObject acForm, .Parent.Name, False
1110          .Parent.cmdClose.SetFocus
1120        End With
1130      End Select
1140    End If

EXITP:
1150    KeyCode = intRetVal
1160    Exit Sub

ERRH:
1170    intRetVal = 0
1180    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1190    Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1320    intShiftDown = (Shift And acShiftMask) > 0
1330    intAltDown = (Shift And acAltMask) > 0
1340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1370        With Me
1380          intRetVal = 0
1390          lngRecsCur = RecCnt  ' ** Function: Below.
1400          If .CurrentRecord < lngRecsCur Then
1410            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1420            .FocusHolder.SetFocus
1430          Else
1440            DoCmd.SelectObject acForm, .Parent.Name, False
1450            If .Parent.cmdPrintReport.Enabled = True Then
1460              .Parent.cmdPrintReport.SetFocus
1470            Else
1480              .Parent.chkIncludeTerritories.SetFocus
1490            End If
1500          End If
1510        End With
1520      End Select
1530    End If

        ' ** Shift keys.
1540    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1550      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1560        With Me
1570          intRetVal = 0
1580          If .CurrentRecord > 1 Then
1590            MoveRec acCmdRecordsGoToPrevious
1600            .FocusHolder.SetFocus
1610          Else
1620            DoCmd.SelectObject acForm, .Parent.Name, False
1630            .Parent.cmdClose.SetFocus
1640          End If
1650        End With
1660      End Select
1670    End If

EXITP:
1680    KeyCode = intRetVal
1690    Exit Sub

ERRH:
1700    intRetVal = 0
1710    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1720    Case Else
1730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1740    End Select
1750    Resume EXITP

End Sub

Private Sub state_code_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "state_code_lbl_DblClick"

1810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub state_code_GotFocus()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "state_code_GotFocus"

1910    With Me
1920      .state_code.SelLength = 0
1930      .state_code.SelStart = 9
1940    End With

EXITP:
1950    Exit Sub

ERRH:
1960    Select Case ERR.Number
        Case Else
1970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1980    End Select
1990    Resume EXITP

End Sub

Private Sub state_code_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "state_code_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2020    intShiftDown = (Shift And acShiftMask) > 0
2030    intAltDown = (Shift And acAltMask) > 0
2040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2070        With Me
2080          intRetVal = 0
2090          .state_name.SetFocus
2100        End With
2110      End Select
2120    End If

        ' ** Shift keys.
2130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2150        With Me
2160          intRetVal = 0
2170          If .CurrentRecord > 1 Then
2180            MoveRec acCmdRecordsGoToPrevious
2190            .FocusHolder.SetFocus
2200          Else
2210            DoCmd.SelectObject acForm, .Parent.Name, False
2220            .Parent.cmdClose.SetFocus
2230          End If
2240        End With
2250      End Select
2260    End If

EXITP:
2270    KeyCode = intRetVal
2280    Exit Sub

ERRH:
2290    intRetVal = 0
2300    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2310    Case Else
2320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2330    End Select
2340    Resume EXITP

End Sub

Private Sub state_name_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "state_name_lbl_DblClick"

2410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub state_name_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "state_name_GotFocus"

2510    With Me
2520      .state_name.SelLength = 0
2530      .state_name.SelStart = 99
2540    End With

EXITP:
2550    Exit Sub

ERRH:
2560    Select Case ERR.Number
        Case Else
2570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2580    End Select
2590    Resume EXITP

End Sub

Private Sub state_name_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "state_name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2620    intShiftDown = (Shift And acShiftMask) > 0
2630    intAltDown = (Shift And acAltMask) > 0
2640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2670        With Me
2680          intRetVal = 0
2690          lngRecsCur = RecCnt  ' ** Function: Below
2700          If .CurrentRecord < lngRecsCur Then
2710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2720            .FocusHolder.SetFocus
2730          Else
2740            DoCmd.SelectObject acForm, .Parent.Name
2750            If .Parent.cmdPrintReport.Enabled = True Then
2760              .Parent.cmdPrintReport.SetFocus
2770            Else
2780              .Parent.chkIncludeTerritories.SetFocus
2790            End If
2800          End If
2810        End With
2820      End Select
2830    End If

        ' ** Shift keys.
2840    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2850      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2860        With Me
2870          intRetVal = 0
2880          .state_code.SetFocus
2890        End With
2900      End Select
2910    End If

EXITP:
2920    KeyCode = intRetVal
2930    Exit Sub

ERRH:
2940    intRetVal = 0
2950    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2960    Case Else
2970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2980    End Select
2990    Resume EXITP

End Sub

Private Sub state_territory_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "state_territory_lbl_DblClick"

3010    With Me
3020      If .state_territory.Locked = True Then
3030        SortNow THIS_PROC  ' ** Procedure: Below.
3040      End If
3050    End With

EXITP:
3060    Exit Sub

ERRH:
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub state_canada_lbl_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "state_canada_lbl_DblClick"

3210    With Me
3220      If .state_canada.Locked = True Then
3230        SortNow THIS_PROC  ' ** Procedure: Below.
3240      End If
3250    End With

EXITP:
3260    Exit Sub

ERRH:
3270    Select Case ERR.Number
        Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

End Sub

Private Sub state_military_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "state_military_lbl_DblClick"

3410    With Me
3420      If .state_military.Locked = True Then
3430        SortNow THIS_PROC  ' ** Procedure: Below.
3440      End If
3450    End With

EXITP:
3460    Exit Sub

ERRH:
3470    Select Case ERR.Number
        Case Else
3480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3490    End Select
3500    Resume EXITP

End Sub

Public Function FindCode(strTypeCode As String) As String
' ** Find first code matching specified type.

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "FindCode"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long
        Dim strRetVal As String

3610    strRetVal = vbNullString

3620    With Me
3630      If strTypeCode <> vbNullString Then
3640        Set rst = .RecordsetClone
3650        If rst.BOF = True And rst.EOF = True Then
              ' ** No recs?!
3660        Else
3670          With rst
3680            .MoveLast
3690            lngRecs = .RecordCount
3700            .MoveFirst
3710            For lngX = 1& To lngRecs
3720              Select Case strTypeCode
                  Case "US"
                    ' ** US states alone.
3730                If ![state_territory] = False And ![state_canada] = False And ![state_military] = False Then
3740                  strRetVal = ![state_code]
3750                  Exit For
3760                End If
3770              Case "TR"
                    ' ** US territories.
3780                If ![state_territory] = True Then
3790                  strRetVal = ![state_code]
3800                  Exit For
3810                End If
3820              Case "CN"
                    ' ** Canadian provinces.
3830                If ![state_canada] = True Then
3840                  strRetVal = ![state_code]
3850                  Exit For
3860                End If
3870              Case "ML"
                    ' ** US military codes.
3880                If ![state_military] = True Then
3890                  strRetVal = ![state_code]
3900                  Exit For
3910                End If
3920              End Select
3930              If lngX < lngRecs Then .MoveNext
3940            Next
3950          End With
3960        End If
3970        rst.Close
3980      End If
3990    End With

EXITP:
4000    Set rst = Nothing
4010    FindCode = strRetVal
4020    Exit Function

ERRH:
4030    strRetVal = vbNullString
4040    Select Case ERR.Number
        Case Else
4050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4060    End Select
4070    Resume EXITP

End Function

Public Function RecCnt() As Long

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

4110    lngRetVal = 0&

4120    With Me
4130      Set rst = .RecordsetClone
4140      With rst
4150        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
4160        Else
4170          .MoveLast
4180          lngRetVal = .RecordCount
4190        End If
4200        .Close
4210      End With
4220    End With

EXITP:
4230    Set rst = Nothing
4240    RecCnt = lngRetVal
4250    Exit Function

ERRH:
4260    lngRetVal = 0&
4270    Select Case ERR.Number
        Case Else
4280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4290    End Select
4300    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

4410    Select Case intWhere
        Case acCmdRecordsGoToFirst
4420      DoCmd.RunCommand acCmdRecordsGoToFirst
4430    Case acCmdRecordsGoToLast
4440      DoCmd.RunCommand acCmdRecordsGoToLast
4450    Case acCmdRecordsGoToPrevious
4460      DoCmd.RunCommand acCmdRecordsGoToPrevious
4470    Case acCmdRecordsGoToNext
4480      DoCmd.RunCommand acCmdRecordsGoToNext
4490    Case Else
4500      If IsMissing(varID) = False Then
4510        With Me
4520          Set rst = .RecordsetClone
4530          With rst
4540            .FindFirst "[state_code] = '" & varID & "'"
4550            If .NoMatch = False Then
4560              Me.Bookmark = .Bookmark
4570            End If
4580            .Close
4590          End With
4600        End With
4610      End If
4620    End Select

EXITP:
4630    Set rst = Nothing
4640    Exit Sub

ERRH:
4650    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4660    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4670    Case Else
4680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4690    End Select
4700    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim strTmp01 As String

        'Const strStdAsc As String = ", [state_military] DESC, [state_canada] DESC, [state_territory] DESC, [state_multi], [state_name]"

4810    With Me
4820      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
4830        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4840      End If
4850      .Controls(strSortLbl).Visible = False
4860      .Controls(strSortLine).Visible = False
4870      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the the right side of the form.
4880      .Sort_line2a.Visible = False
4890      .Sort_line2b.Visible = False
4900      If strProc = "Form_Load" Then
4910        strCalled = "state_name"
4920        strSortNow = strSortOrig
4930        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4940        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
4950        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4960        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4970        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4980        .Controls(strSortLbl).Top = lngSortLbl_Top
4990        .Controls(strSortLbl).Left = lngSortLbl_Left
5000        .Controls(strSortLine).Top = lngSortLine_Top
5010        .Controls(strSortLine).Left = lngSortLine_Left
5020        .Controls(strSortLine).Width = lngSortLine_Width
5030        .Controls(strSortLbl).Caption = strArwUp
5040        .Controls(strSortLbl).ForeColor = CLR_DKBLU
5050      Else
5060        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
5070        Select Case strCalled
            Case "state_territory", "state_canada", "state_military"
5080          lngSortLbl_Top = .Controls(strCalled & "_lbl").Top
5090          lngSortLine_Top = .Controls(strCalled & "_lbl").Top
5100        Case Else
5110          lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5120          lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5130        End Select
5140        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5150        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5160        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5170        .Controls(strSortLbl).Top = lngSortLbl_Top
5180        .Controls(strSortLbl).Left = lngSortLbl_Left
5190        Select Case strCalled
            Case "state_territory", "state_canada", "state_military"
5200          .Sort_line2a.Top = lngSortLine_Top
5210          .Sort_line2b.Top = (.Sort_line2a.Top + lngTpp)
5220          .Sort_line2a.Left = lngSortLine_Left
5230          .Sort_line2b.Left = .Sort_line2a.Left
5240          .Sort_line2a.Width = lngSortLine_Width
5250          .Sort_line2b.Width = .Sort_line2a.Width
5260        Case Else
5270          .Controls(strSortLine).Top = lngSortLine_Top
5280          .Controls(strSortLine).Left = lngSortLine_Left
5290          .Controls(strSortLine).Width = lngSortLine_Width
5300        End Select
5310        If strCalled = "state_name" Then
5320          strSortAsc = strSortOrig
5330          If strSortNow = strSortAsc Then
5340            strSortNow = "[state_military], [state_canada], [state_territory], [state_multi] DESC, [state_name] DESC"
5350            .Controls(strSortLbl).Caption = strArwDn
5360            .Controls(strSortLbl).ForeColor = CLR_DKRED
5370          Else
5380            strSortNow = strSortAsc
5390            .Controls(strSortLbl).Caption = strArwUp
5400            .Controls(strSortLbl).ForeColor = CLR_DKBLU
5410          End If
5420          lngSortLbl_Top = .Controls(strCalled & "_lbl").Top
5430          .Controls(strSortLbl).Top = lngSortLbl_Top
5440        Else
5450          .Controls(strSortLbl).Caption = strArwUp
5460          .Controls(strSortLbl).ForeColor = CLR_DKBLU
5470          Select Case strCalled
              Case "state_code"
5480            strSortAsc = "[state_code]"
5490            lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
5500            .Controls(strSortLbl).Left = lngSortLbl_Left
5510          Case "state_territory"
5520            strSortAsc = "[state_territory], [state_canada], [state_military], [state_multi], [state_name]"
5530          Case "state_canada"
5540            strSortAsc = "[state_canada], [state_military], [state_territory], [state_multi], [state_name]"
5550            lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
5560            .Controls(strSortLbl).Left = lngSortLbl_Left
5570          Case "state_military"
5580            strSortAsc = "[state_military], [state_territory], [state_canada] DESC, [state_multi], [state_name]"
5590          End Select
5600          If strSortNow = strSortAsc Then
5610            Select Case strCalled
                Case "state_code"
5620              strTmp01 = "[state_code] DESC"
5630            Case "state_territory"
5640              strTmp01 = "[state_territory] DESC, [state_canada] DESC, [state_military] DESC, [state_multi] DESC, [state_name] DESC"
5650            Case "state_canada"
5660              strTmp01 = "[state_canada] DESC, [state_military] DESC, [state_territory] DESC, [state_multi] DESC, [state_name] DESC"
5670            Case "state_military"
5680              strTmp01 = "[state_military] DESC, [state_territory] DESC, [state_canada] DESC, [state_multi] DESC, [state_name] DESC"
5690            End Select
5700            strSortNow = strTmp01
5710            .Controls(strSortLbl).Caption = strArwDn
5720            .Controls(strSortLbl).ForeColor = CLR_DKRED
5730          Else
5740            strSortNow = strSortAsc
5750          End If
5760        End If
5770      End If
5780      .Controls(strSortLbl).Visible = True
5790      Select Case strCalled
          Case "state_territory", "state_canada", "state_military"
5800        .Sort_line2a.Visible = True
5810        .Sort_line2b.Visible = True
5820      Case Else
5830        .Controls(strSortLine).Visible = True
5840      End Select
5850      .OrderBy = strSortNow
5860      .OrderByOn = True
5870    End With

        ' ** Potential Sorts:
        ' **   [state_military] DESC, [state_canada] DESC, [state_territory] DESC, [state_multi], [state_name]
        ' **   [state_code]
        ' **   [state_territory], [state_canada], [state_military], [state_multi], [state_name]
        ' **   [state_canada], [state_military], [state_territory], [state_multi], [state_name]
        ' **   [state_military], [state_territory], [state_canada] DESC, [state_multi], [state_name]

EXITP:
5880    Exit Sub

ERRH:
5890    Select Case ERR.Number
        Case Else
5900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5910    End Select
5920    Resume EXITP

End Sub

Public Function SortNow_Get() As String

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

6010    strRetVal = strSortNow

EXITP:
6020    SortNow_Get = strRetVal
6030    Exit Function

ERRH:
6040    strRetVal = vbNullString
6050    Select Case ERR.Number
        Case Else
6060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6070    End Select
6080    Resume EXITP

End Function
