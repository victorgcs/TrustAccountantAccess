VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssetTypeGrouping_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssetTypeGrouping_Sub"

'VGC 05/04/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmAssetTypeGrouping}
' **   Show Types:       H {chkShowAssetType on frmAssetTypeGrouping}
' **   Print Report:     P {cmdPrintReport on frmAssetTypeGrouping}
' **   Exit:             X {cmdClose on frmAssetTypeGrouping}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[assettypegroup_sequence], [assettype]"
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
150       If strCallingForm = "frmAssetTypeGrouping" Then

160         lngSortLbl_Width = .Sort_lbl.Width

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         lngGroupName_Left = .assettypegroup_name.Left
210         lngGroupSeq_Left = .assettypegroup_sequence.Left
220         lngGroupSeqLbl_Left = .assettypegroup_sequence_lbl.Left

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
420       .assettype.ForeColor = CLR_DISABLED_FG
430       .assettype.BackColor = CLR_DISABLED_BG
440       .assettypegroup_name.ForeColor = CLR_DISABLED_FG
450       .assettypegroup_name.BackColor = CLR_DISABLED_BG
460       .assettypegroup_sequence.ForeColor = CLR_DISABLED_FG
470       .assettypegroup_sequence.BackColor = CLR_DISABLED_BG
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
750           .Parent.cmdClose_Click  ' ** Form Procedure: frmAssetTypeGrouping.
760         End With
770       End Select
780     End If

        ' ** Alt keys.
790     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
800       Select Case intRetVal
          Case vbKeyX
810         With Me
820           intRetVal = 0
830           .Parent.cmdClose_Click  ' ** Form Procedure: frmAssetTypeGrouping.
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
950             .Parent.chkShowAssetType.SetFocus
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
1450            .Parent.chkShowAssetType.SetFocus
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

Private Sub assettype_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_lbl_DblClick"

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

Private Sub assettype_GotFocus()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_GotFocus"

1910    With Me
1920      .assettype.SelLength = 0
1930      .assettype.SelStart = 9
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

Private Sub assettype_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_KeyDown"

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
2090          .assettypegroup_name.SetFocus
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

Private Sub assettypegroup_name_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_name_lbl_DblClick"

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

Private Sub assettypegroup_name_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_name_GotFocus"

2510    With Me
2520      .assettypegroup_name.SelLength = 0
2530      .assettypegroup_name.SelStart = 99
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

Private Sub assettypegroup_name_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_name_KeyDown"

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
2690          .assettypegroup_sequence.SetFocus
2700        End With
2710      End Select
2720    End If

        ' ** Shift keys.
2730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2750        With Me
2760          intRetVal = 0
2770          .assettype.SetFocus
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

Private Sub assettypegroup_sequence_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_sequence_lbl_DblClick"

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

Private Sub assettypegroup_sequence_GotFocus()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_sequence_GotFocus"

3010    With Me
3020      .assettypegroup_sequence.SelLength = 0
3030      .assettypegroup_sequence.SelStart = 9
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

Private Sub assettypegroup_sequence_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_sequence_KeyDown"

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
3240            DoCmd.SelectObject acForm, .Parent.Name, False
3250            .Parent.chkShowAssetType.SetFocus
3260          End If
3270        End With
3280      End Select
3290    End If

        ' ** Shift keys.
3300    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3310      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3320        With Me
3330          intRetVal = 0
3340          .assettypegroup_name.SetFocus
3350        End With
3360      End Select
3370    End If

EXITP:
3380    KeyCode = intRetVal
3390    Exit Sub

ERRH:
3400    intRetVal = 0
3410    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3420    Case Else
3430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3440    End Select
3450    Resume EXITP

End Sub

Public Sub ShowAssetType(blnShow As Boolean)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAssetType"

        Dim blnSortHere As Boolean, blnResort As Boolean, blnIsDesc As Boolean
        Dim lngNewLeft As Long

3510    With Me
3520      blnSortHere = False: blnResort = False
3530      Select Case blnShow
          Case True
3540        .RecordSource = "qryAssetTypeGrouping_04"
3550        If .Sort_line.Left = .assettypegroup_name_lbl.Left Then blnSortHere = True
3560        .assettypegroup_name.Left = lngGroupName_Left
3570        .assettypegroup_name_lbl.Left = lngGroupName_Left
3580        .assettypegroup_name_lbl_line.Left = lngGroupName_Left
3590        If blnSortHere = True Then
3600          .Sort_line.Left = .assettypegroup_name_lbl.Left
3610          .Sort_lbl.Left = ((.assettypegroup_name_lbl.Left + .assettypegroup_name_lbl.Width) - lngSortLbl_Width)
3620        End If
3630        blnSortHere = False
3640        If .Sort_line.Left = .assettypegroup_sequence_lbl.Left Then blnSortHere = True
3650        .assettypegroup_sequence.Left = lngGroupSeq_Left
3660        .assettypegroup_sequence_lbl.Left = lngGroupSeqLbl_Left
3670        .assettypegroup_sequence_lbl2.Left = lngGroupSeqLbl_Left
3680        .assettypegroup_sequence_lbl_line.Left = lngGroupSeqLbl_Left
3690        If blnSortHere = True Then
3700          .Sort_line.Left = .assettypegroup_sequence_lbl.Left
3710          .Sort_lbl.Left = ((.assettypegroup_sequence_lbl.Left + .assettypegroup_sequence_lbl.Width) - lngSortLbl_Width)
3720        End If
3730        blnSortHere = False
3740        .assettype.Visible = True
3750        .assettype_lbl.Visible = True
3760        .assettype_lbl2.Visible = True
3770        .assettype_lbl_line.Visible = True
3780      Case False
3790        If .Sort_line.Left = .assettype_lbl.Left Then blnSortHere = True
3800        .assettype.Visible = False
3810        .assettype_lbl.Visible = False
3820        .assettype_lbl2.Visible = False
3830        .assettype_lbl_line.Visible = False
3840        If blnSortHere = True Then
3850          .Sort_line.Visible = False
3860          .Sort_lbl.Visible = False
3870          blnResort = True
3880        End If
3890        blnSortHere = False
3900        If .Sort_line.Left = .assettypegroup_sequence_lbl.Left Then blnSortHere = True
3910        .assettypegroup_sequence.Left = .assettype_lbl.Left + (lngGroupSeq_Left - lngGroupSeqLbl_Left)
3920        .assettypegroup_sequence_lbl.Left = .assettype_lbl.Left
3930        .assettypegroup_sequence_lbl2.Left = .assettype_lbl.Left
3940        .assettypegroup_sequence_lbl_line.Left = .assettype_lbl.Left
3950        If blnSortHere = True Then
3960          .Sort_line.Left = .assettypegroup_sequence_lbl.Left
3970          .Sort_lbl.Left = ((.assettypegroup_sequence_lbl.Left + .assettypegroup_sequence_lbl.Width) - lngSortLbl_Width)
3980        End If
3990        blnSortHere = False
4000        lngNewLeft = (.assettype_lbl.Left + .assettypegroup_sequence_lbl.Width) + (4& * lngTpp)
4010        If .Sort_line.Left = .assettypegroup_name_lbl.Left Then blnSortHere = True
4020        .assettypegroup_name.Left = lngNewLeft
4030        .assettypegroup_name_lbl.Left = lngNewLeft
4040        .assettypegroup_name_lbl_line.Left = lngNewLeft
4050        .RecordSource = "qryAssetTypeGrouping_03"
4060        If blnSortHere = True Then
4070          .Sort_line.Left = .assettypegroup_name_lbl.Left
4080          .Sort_lbl.Left = ((.assettypegroup_name_lbl.Left + .assettypegroup_name_lbl.Width) - lngSortLbl_Width)
4090        End If
4100        blnSortHere = False
4110      End Select
4120      If InStr(strSortNow, " DESC") > 0 Then blnIsDesc = True Else blnIsDesc = False
4130      If blnResort = True Then
4140        SortNow "Form_Load"  ' ** Procedure: Below.
4150      ElseIf Left(strSortNow, 11) = "[assettype]" Then
4160        SortNow "assettype_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
4170      ElseIf Left(strSortNow, 25) = "[assettypegroup_sequence]" Then
4180        SortNow "assettypegroup_sequence_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
4190      ElseIf Left(strSortNow, 21) = "[assettypegroup_name]" Then
4200        SortNow "assettypegroup_name_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
4210      End If
4220    End With

EXITP:
4230    Exit Sub

ERRH:
4240    Select Case ERR.Number
        Case Else
4250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4260    End Select
4270    Resume EXITP

End Sub

Public Function RecCnt() As Long

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

4310    lngRetVal = 0&

4320    With Me
4330      Set rst = .RecordsetClone
4340      With rst
4350        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
4360        Else
4370          .MoveLast
4380          lngRetVal = .RecordCount
4390        End If
4400        .Close
4410      End With
4420    End With

EXITP:
4430    Set rst = Nothing
4440    RecCnt = lngRetVal
4450    Exit Function

ERRH:
4460    lngRetVal = 0&
4470    Select Case ERR.Number
        Case Else
4480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4490    End Select
4500    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

4610    Select Case intWhere
        Case acCmdRecordsGoToFirst
4620      DoCmd.RunCommand acCmdRecordsGoToFirst
4630    Case acCmdRecordsGoToLast
4640      DoCmd.RunCommand acCmdRecordsGoToLast
4650    Case acCmdRecordsGoToPrevious
4660      DoCmd.RunCommand acCmdRecordsGoToPrevious
4670    Case acCmdRecordsGoToNext
4680      DoCmd.RunCommand acCmdRecordsGoToNext
4690    Case Else
4700      If IsMissing(varID) = False Then
4710        With Me
4720          Set rst = .RecordsetClone
4730          With rst
4740            .FindFirst "[assettype] = '" & varID & "'"
4750            If .NoMatch = False Then
4760              Me.Bookmark = .Bookmark
4770            End If
4780            .Close
4790          End With
4800        End With
4810      End If
4820    End Select

EXITP:
4830    Set rst = Nothing
4840    Exit Sub

ERRH:
4850    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4860    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4870    Case Else
4880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4890    End Select
4900    Resume EXITP

End Sub

Public Sub SortNow(strProc As String, Optional varShow As Variant, Optional varIsDesc As Variant)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim blnDontReverse As Boolean, blnIsDesc As Boolean
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [assettypegroup_sequence], [assettype]"

5010    With Me
5020      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5030        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5040      End If
5050      Select Case IsMissing(varShow)
          Case True
5060        blnDontReverse = False
5070      Case False
5080        blnDontReverse = CBool(varShow)
5090      End Select
5100      Select Case IsMissing(varIsDesc)
          Case True
5110        blnIsDesc = False
5120      Case False
5130        blnIsDesc = CBool(varIsDesc)
5140      End Select
5150      .Controls(strSortLbl).Visible = False
5160      .Controls(strSortLine).Visible = False
5170      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
5180      If strProc = "Form_Load" Then
5190        strCalled = "assettypegroup_sequence"
5200        strSortNow = strSortOrig
5210        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5220        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5230        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5240        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5250        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5260        .Controls(strSortLbl).Top = lngSortLbl_Top
5270        .Controls(strSortLbl).Left = lngSortLbl_Left
5280        .Controls(strSortLine).Top = lngSortLine_Top
5290        .Controls(strSortLine).Left = lngSortLine_Left
5300        .Controls(strSortLine).Width = lngSortLine_Width
5310        .Controls(strSortLbl).Caption = strArwUp
5320        .Controls(strSortLbl).ForeColor = CLR_DKBLU
5330      Else
5340        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
5350        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5360        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5370        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5380        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5390        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5400        .Controls(strSortLbl).Top = lngSortLbl_Top
5410        .Controls(strSortLbl).Left = lngSortLbl_Left
5420        .Controls(strSortLine).Top = lngSortLine_Top
5430        .Controls(strSortLine).Left = lngSortLine_Left
5440        .Controls(strSortLine).Width = lngSortLine_Width
5450        If strCalled = "assettypegroup_sequence" Then
5460          strSortAsc = strSortOrig
5470          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
5480            strSortNow = "[assettypegroup_sequence] DESC, [assettype] DESC"
5490            .Controls(strSortLbl).Caption = strArwDn
5500            .Controls(strSortLbl).ForeColor = CLR_DKRED
5510          Else
5520            strSortNow = strSortAsc
5530            .Controls(strSortLbl).Caption = strArwUp
5540            .Controls(strSortLbl).ForeColor = CLR_DKBLU
5550          End If
5560        Else
5570          .Controls(strSortLbl).Caption = strArwUp
5580          .Controls(strSortLbl).ForeColor = CLR_DKBLU
5590          Select Case strCalled
              Case "assettype"
5600            strSortAsc = "[assettype]" & strStdAsc
5610            lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
5620            .Controls(strSortLbl).Left = lngSortLbl_Left
5630          Case "assettypegroup_name"
5640            strSortAsc = "[assettypegroup_name]" & strStdAsc
5650          End Select
5660          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
5670            intCnt = CharCnt(strSortAsc, ",") + 1
5680            Select Case intCnt
                Case 1
5690              strTmp01 = strSortAsc & " DESC"
5700            Case 2
5710              intPos01 = InStr(strSortAsc, ",")
5720              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5730              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
5740            Case 3
5750              intPos01 = InStr(strSortAsc, ",")
5760              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5770              strTmp02 = Mid(strSortAsc, intPos01)
5780              intPos01 = InStr(2, strSortAsc, ",")
5790              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5800              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
5810            End Select
5820            strSortNow = strTmp01
5830            .Controls(strSortLbl).Caption = strArwDn
5840            .Controls(strSortLbl).ForeColor = CLR_DKRED
5850          Else
5860            strSortNow = strSortAsc
5870          End If
5880        End If
5890      End If
5900      .Controls(strSortLbl).Visible = True
5910      .Controls(strSortLine).Visible = True
5920      .OrderBy = strSortNow
5930      .OrderByOn = True
5940    End With

        ' ** Potential Sorts:
        ' **   [assettype]
        ' **   [assettypegroup_name], [assettype]
        ' **   [assettypegroup_sequence], [assettype]

EXITP:
5950    Exit Sub

ERRH:
5960    Select Case ERR.Number
        Case Else
5970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5980    End Select
5990    Resume EXITP

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
6040    strRetVal = strSortOrig
6050    Select Case ERR.Number
        Case Else
6060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6070    End Select
6080    Resume EXITP

End Function
