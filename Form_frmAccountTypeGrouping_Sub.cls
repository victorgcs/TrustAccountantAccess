VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountTypeGrouping_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountTypeGrouping_Sub"

'VGC 04/19/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmAccountTypeGrouping}
' **   Show Accts:       H {chkShowAccountType on frmAccountTypeGrouping}
' **   Print Report:     P {cmdPrintReport on frmAccountTypeGrouping}
' **   Exit:             X {cmdClose on frmAccountTypeGrouping}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[accounttypegroup_sequence], [accounttype]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngGroupName_Left As Long, lngGroupSeq_Left As Long, lngGroupSeqLbl_Left As Long
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
150       If strCallingForm = "frmAccountTypeGrouping" Then

160         lngSortLbl_Width = .Sort_lbl.Width

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
200         lngGroupName_Left = .accounttypegroup_name.Left
210         lngGroupSeq_Left = .accounttypegroup_sequence.Left
220         lngGroupSeqLbl_Left = .accounttypegroup_sequence_lbl.Left

230       Else
240         Cancel = -1
250       End If
260     End With

270     If Cancel = -1 Then
280       Beep
290       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
300     End If

EXITP:
310     Exit Sub

ERRH:
320     Select Case ERR.Number
        Case Else
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
340     End Select
350     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .accounttype.ForeColor = CLR_DISABLED_FG
430       .accounttype.BackColor = CLR_DISABLED_BG
440       .accounttypegroup_name.ForeColor = CLR_DISABLED_FG
450       .accounttypegroup_name.BackColor = CLR_DISABLED_BG
460       .accounttypegroup_sequence.ForeColor = CLR_DISABLED_FG
470       .accounttypegroup_sequence.BackColor = CLR_DISABLED_BG
480     End With

490     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
500     Exit Sub

ERRH:
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

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
750           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountTypeGrouping.
760         End With
770       End Select
780     End If

        ' ** Alt keys.
790     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
800       Select Case intRetVal
          Case vbKeyX
810         With Me
820           intRetVal = 0
830           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountTypeGrouping.
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
950             .Parent.chkShowAccountType.SetFocus
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
1450            .Parent.chkShowAccountType.SetFocus
1460          End If
1470        End With
1480      End Select
1490    End If

        ' ** Shift keys.
1500    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1510      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1520        With Me
1530          intRetVal = 0
1540          If .CurrentRecord > 1 Then
1550            MoveRec acCmdRecordsGoToPrevious
1560            .FocusHolder.SetFocus
1570          Else
1580            DoCmd.SelectObject acForm, .Parent.Name, False
1590            .Parent.cmdClose.SetFocus
1600          End If
1610        End With
1620      End Select
1630    End If

EXITP:
1640    KeyCode = intRetVal
1650    Exit Sub

ERRH:
1660    intRetVal = 0
1670    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1680    Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Sub

Private Sub accounttype_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_lbl_DblClick"

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

Private Sub accounttype_GotFocus()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_GotFocus"

1910    With Me
1920      .accounttype.SelLength = 0
1930      .accounttype.SelStart = 9
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

Private Sub accounttype_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_KeyDown"

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
2090          .accounttypegroup_name.SetFocus
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

Private Sub accounttypegroup_name_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_name_lbl_DblClick"

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

Private Sub accounttypegroup_name_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_name_GotFocus"

2510    With Me
2520      .accounttypegroup_name.SelLength = 0
2530      .accounttypegroup_name.SelStart = 99
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

Private Sub accounttypegroup_name_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_name_KeyDown"

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
2690          .accounttypegroup_sequence.SetFocus
2700        End With
2710      End Select
2720    End If

        ' ** Shift keys.
2730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2750        With Me
2760          intRetVal = 0
2770          .accounttype.SetFocus
2780        End With
2790      End Select
2800    End If

EXITP:
2810    KeyCode = intRetVal
2820    Exit Sub

ERRH:
2830    intRetVal = 0
2840    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2850    Case Else
2860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2870    End Select
2880    Resume EXITP

End Sub

Private Sub accounttypegroup_sequence_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_sequence_lbl_DblClick"

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

Private Sub accounttypegroup_sequence_GotFocus()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_sequence_GotFocus"

3010    With Me
3020      .accounttypegroup_sequence.SelLength = 0
3030      .accounttypegroup_sequence.SelStart = 9
3040    End With

EXITP:
3050    Exit Sub

ERRH:
3060    Select Case ERR.Number
        Case Else
3070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3080    End Select
3090    Resume EXITP

End Sub

Private Sub accounttypegroup_sequence_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttypegroup_sequence_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3170        With Me
3180          intRetVal = 0
3190          lngRecsCur = RecCnt  ' ** Function: Below.
3200          If .CurrentRecord < lngRecsCur Then
3210            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3220            .FocusHolder.SetFocus
3230          Else
3240            .Parent.chkShowAccountType.SetFocus
3250          End If
3260        End With
3270      End Select
3280    End If

        ' ** Shift keys.
3290    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3300      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3310        With Me
3320          intRetVal = 0
3330          .accounttypegroup_name.SetFocus
3340        End With
3350      End Select
3360    End If

EXITP:
3370    KeyCode = intRetVal
3380    Exit Sub

ERRH:
3390    intRetVal = 0
3400    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3410    Case Else
3420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3430    End Select
3440    Resume EXITP

End Sub

Public Sub ShowAcctType(blnShow As Boolean)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAcctType"

        Dim blnSortHere As Boolean, blnResort As Boolean, blnIsDesc As Boolean
        Dim lngNewLeft As Long

3510    With Me
3520      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
3530        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
3540      End If
3550      blnSortHere = False: blnResort = False
3560      Select Case blnShow
          Case True
3570        .RecordSource = "qryAccountTypeGrouping_04"
3580        If .Sort_line.Left = .accounttypegroup_name_lbl.Left Then blnSortHere = True
3590        .accounttypegroup_name.Left = lngGroupName_Left
3600        .accounttypegroup_name_lbl.Left = lngGroupName_Left
3610        .accounttypegroup_name_lbl_line.Left = lngGroupName_Left
3620        If blnSortHere = True Then
3630          .Sort_line.Left = .accounttypegroup_name_lbl.Left
3640          .Sort_lbl.Left = ((.accounttypegroup_name_lbl.Left + .accounttypegroup_name_lbl.Width) - lngSortLbl_Width)
3650        End If
3660        blnSortHere = False
3670        If .Sort_line.Left = .accounttypegroup_sequence_lbl.Left Then blnSortHere = True
3680        .accounttypegroup_sequence.Left = lngGroupSeq_Left
3690        .accounttypegroup_sequence_lbl.Left = lngGroupSeqLbl_Left
3700        .accounttypegroup_sequence_lbl2.Left = lngGroupSeqLbl_Left
3710        .accounttypegroup_sequence_lbl_line.Left = lngGroupSeqLbl_Left
3720        If blnSortHere = True Then
3730          .Sort_line.Left = .accounttypegroup_sequence_lbl.Left
3740          .Sort_lbl.Left = ((.accounttypegroup_sequence_lbl.Left + .accounttypegroup_sequence_lbl.Width) - lngSortLbl_Width)
3750        End If
3760        blnSortHere = False
3770        If .Sort_line.Left = .accounttype_lbl.Left Then blnSortHere = True
3780        .accounttype.Visible = True
3790        .accounttype_lbl.Visible = True
3800        .accounttype_lbl2.Visible = True
3810        .accounttype_lbl_line.Visible = True
3820        If blnSortHere = True Then
3830          .Sort_line.Left = .accounttype_lbl.Left
3840          .Sort_lbl.Left = (((.accounttype_lbl.Left + .accounttype_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
3850        End If
3860        blnSortHere = False
3870      Case False
3880        If .Sort_line.Left = .accounttype_lbl.Left Then blnSortHere = True
3890        .accounttype.Visible = False
3900        .accounttype_lbl.Visible = False
3910        .accounttype_lbl2.Visible = False
3920        .accounttype_lbl_line.Visible = False
3930        If blnSortHere = True Then
3940          .Sort_line.Visible = False
3950          .Sort_lbl.Visible = False
3960          blnResort = True
3970        End If
3980        blnSortHere = False
3990        If .Sort_line.Left = .accounttypegroup_sequence_lbl.Left Then blnSortHere = True
4000        .accounttypegroup_sequence.Left = .accounttype_lbl.Left + (lngGroupSeq_Left - lngGroupSeqLbl_Left)
4010        .accounttypegroup_sequence_lbl.Left = .accounttype_lbl.Left
4020        .accounttypegroup_sequence_lbl2.Left = .accounttype_lbl.Left
4030        .accounttypegroup_sequence_lbl_line.Left = .accounttype_lbl.Left
4040        If blnSortHere = True Then
4050          .Sort_line.Left = .accounttypegroup_sequence_lbl.Left
4060          .Sort_lbl.Left = ((.accounttypegroup_sequence_lbl.Left + .accounttypegroup_sequence_lbl.Width) - lngSortLbl_Width)
4070        End If
4080        blnSortHere = False
4090        lngNewLeft = (.accounttype_lbl.Left + .accounttypegroup_sequence_lbl.Width) + (4& * lngTpp)
4100        If .Sort_line.Left = .accounttypegroup_name_lbl.Left Then blnSortHere = True
4110        .accounttypegroup_name.Left = lngNewLeft
4120        .accounttypegroup_name_lbl.Left = lngNewLeft
4130        .accounttypegroup_name_lbl_line.Left = lngNewLeft
4140        If blnSortHere = True Then
4150          .Sort_line.Left = .accounttypegroup_name_lbl.Left
4160          .Sort_lbl.Left = ((.accounttypegroup_name_lbl.Left + .accounttypegroup_name_lbl.Width) - lngSortLbl_Width)
4170        End If
4180        blnSortHere = False
4190        .RecordSource = "qryAccountTypeGrouping_03"
4200      End Select
4210      If InStr(strSortNow, " DESC") > 0 Then blnIsDesc = True Else blnIsDesc = False
4220      If blnResort = True Then
4230        SortNow "Form_Load"  ' ** Procedure: Below.
4240      ElseIf Left(strSortNow, 13) = "[accounttype]" Then
4250        SortNow "accounttype_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
4260      ElseIf Left(strSortNow, 27) = "[accounttypegroup_sequence]" Then
4270        SortNow "accounttypegroup_sequence_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
4280      ElseIf Left(strSortNow, 23) = "[accounttypegroup_name]" Then
4290        SortNow "accounttypegroup_name_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
4300      End If
4310    End With

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Public Function RecCnt() As Long

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

4410    lngRetVal = 0&

4420    With Me
4430      Set rst = .RecordsetClone
4440      With rst
4450        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
4460        Else
4470          .MoveLast
4480          lngRetVal = .RecordCount
4490        End If
4500        .Close
4510      End With
4520    End With

EXITP:
4530    Set rst = Nothing
4540    RecCnt = lngRetVal
4550    Exit Function

ERRH:
4560    lngRetVal = 0&
4570    Select Case ERR.Number
        Case Else
4580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4590    End Select
4600    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

4710    Select Case intWhere
        Case acCmdRecordsGoToFirst
4720      DoCmd.RunCommand acCmdRecordsGoToFirst
4730    Case acCmdRecordsGoToLast
4740      DoCmd.RunCommand acCmdRecordsGoToLast
4750    Case acCmdRecordsGoToPrevious
4760      DoCmd.RunCommand acCmdRecordsGoToPrevious
4770    Case acCmdRecordsGoToNext
4780      DoCmd.RunCommand acCmdRecordsGoToNext
4790    Case Else
4800      If IsMissing(varID) = False Then
4810        With Me
4820          Set rst = .RecordsetClone
4830          With rst
4840            .FindFirst "[accounttype] = '" & varID & "'"
4850            If .NoMatch = False Then
4860              Me.Bookmark = .Bookmark
4870            End If
4880            .Close
4890          End With
4900        End With
4910      End If
4920    End Select

EXITP:
4930    Set rst = Nothing
4940    Exit Sub

ERRH:
4950    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4960    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4970    Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Sub

Public Sub SortNow(strProc As String, Optional varShow As Variant, Optional varIsDesc As Variant)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim blnDontReverse As Boolean, blnIsDesc As Boolean
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [accounttypegroup_sequence], [accounttype]"

5110    With Me
5120      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5130        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5140      End If
5150      Select Case IsMissing(varShow)
          Case True
5160        blnDontReverse = False
5170      Case False
5180        blnDontReverse = CBool(varShow)
5190      End Select
5200      Select Case IsMissing(varIsDesc)
          Case True
5210        blnIsDesc = False
5220      Case False
5230        blnIsDesc = CBool(varIsDesc)
5240      End Select
5250      .Controls(strSortLbl).Visible = False
5260      .Controls(strSortLine).Visible = False
5270      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
5280      If strProc = "Form_Load" Then
5290        strCalled = "accounttypegroup_sequence"
5300        strSortNow = strSortOrig
5310        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5320        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5330        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5340        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5350        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5360        .Controls(strSortLbl).Top = lngSortLbl_Top
5370        .Controls(strSortLbl).Left = lngSortLbl_Left
5380        .Controls(strSortLine).Top = lngSortLine_Top
5390        .Controls(strSortLine).Left = lngSortLine_Left
5400        .Controls(strSortLine).Width = lngSortLine_Width
5410        .Controls(strSortLbl).Caption = strArwUp
5420        .Controls(strSortLbl).ForeColor = CLR_DKBLU
5430      Else
5440        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
5450        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5460        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5470        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5480        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5490        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5500        .Controls(strSortLbl).Top = lngSortLbl_Top
5510        .Controls(strSortLbl).Left = lngSortLbl_Left
5520        .Controls(strSortLine).Top = lngSortLine_Top
5530        .Controls(strSortLine).Left = lngSortLine_Left
5540        .Controls(strSortLine).Width = lngSortLine_Width
5550        If strCalled = "accounttypegroup_sequence" Then
5560          strSortAsc = strSortOrig
5570          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
5580            strSortNow = "[accounttypegroup_sequence] DESC, [accounttype] DESC"
5590            .Controls(strSortLbl).Caption = strArwDn
5600            .Controls(strSortLbl).ForeColor = CLR_DKRED
5610          Else
5620            strSortNow = strSortAsc
5630            .Controls(strSortLbl).Caption = strArwUp
5640            .Controls(strSortLbl).ForeColor = CLR_DKBLU
5650          End If
5660        Else
5670          .Controls(strSortLbl).Caption = strArwUp
5680          .Controls(strSortLbl).ForeColor = CLR_DKBLU
5690          Select Case strCalled
              Case "accounttype"
5700            strSortAsc = "[accounttype]"
5710            lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
5720            .Controls(strSortLbl).Left = lngSortLbl_Left
5730          Case "accounttypegroup_name"
5740            strSortAsc = "[accounttypegroup_name]" & strStdAsc
5750          End Select
5760          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
5770            intCnt = CharCnt(strSortAsc, ",") + 1
5780            Select Case intCnt
                Case 1
5790              strTmp01 = strSortAsc & " DESC"
5800            Case 2
5810              intPos01 = InStr(strSortAsc, ",")
5820              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5830              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
5840            Case 3
5850              intPos01 = InStr(strSortAsc, ",")
5860              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5870              strTmp02 = Mid(strSortAsc, intPos01)
5880              intPos01 = InStr(2, strTmp02, ",")
5890              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5900              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
5910            End Select
5920            strSortNow = strTmp01
5930            .Controls(strSortLbl).Caption = strArwDn
5940            .Controls(strSortLbl).ForeColor = CLR_DKRED
5950          Else
5960            strSortNow = strSortAsc
5970          End If
5980        End If
5990      End If
6000      .Controls(strSortLbl).Visible = True
6010      .Controls(strSortLine).Visible = True
6020      .OrderBy = strSortNow
6030      .OrderByOn = True
6040    End With

        ' ** Potential Sorts:
        ' **   [accounttype]
        ' **   [accounttypegroup_name], [accounttypegroup_sequence], [accounttype]
        ' **   [accounttypegroup_sequence], [accounttype]

EXITP:
6050    Exit Sub

ERRH:
6060    Select Case ERR.Number
        Case Else
6070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6080    End Select
6090    Resume EXITP

End Sub

Public Function SortNow_Get() As String

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

6110    strRetVal = strSortNow

EXITP:
6120    SortNow_Get = strRetVal
6130    Exit Function

ERRH:
6140    strRetVal = strSortOrig
6150    Select Case ERR.Number
        Case Else
6160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6170    End Select
6180    Resume EXITP

End Function
