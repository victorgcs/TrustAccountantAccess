VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLoadTimer_Mgmt_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLoadTimer_Mgmt_Sub"

'VGC 10/11/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmLoadTimer_Mgmt}
' **   Clear:            E {cmdClear on frmLoadTimer_Mgmt}
' **   Latest Only:      O {chkLatestOnly on frmLoadTimer_Mgmt}
' **   Elapsed Time:     T {xadload_time}
' **   Exit:             X {cmdClose on frmLoadTimer_Mgmt}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[frm_name1], [frm_name2], [xadload_datemodified] DESC"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Left As Long, lngSortLine_Width As Long

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
150       If strCallingForm = "frmLoadTimer_Mgmt" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
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

420       .dbs_name1.ForeColor = CLR_DISABLED_FG
430       .dbs_name1.BackColor = CLR_DISABLED_BG
440       .frm_name1.ForeColor = CLR_DISABLED_FG
450       .frm_name1.BackColor = CLR_DISABLED_BG
460       .frm_name2.ForeColor = CLR_DISABLED_FG
470       .frm_name2.BackColor = CLR_DISABLED_BG
480       .xadload_start.ForeColor = CLR_DISABLED_FG
490       .xadload_start.BackColor = CLR_DISABLED_BG
500       .xadload_end.ForeColor = CLR_DISABLED_FG
510       .xadload_end.BackColor = CLR_DISABLED_BG
520       .xadload_access_version.ForeColor = CLR_DISABLED_FG
530       .xadload_access_version.BackColor = CLR_DISABLED_BG
540       .xadload_datemodified.ForeColor = CLR_DISABLED_FG
550       .xadload_datemodified.BackColor = CLR_DISABLED_BG

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
        ' **   Elapsed Time:     T {xadload_time}
        ' **   Exit:             X {cmdClose}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyUp
770         With Me
780           intRetVal = 0
790           lngRecsCur = .RecCnt  ' ** Function: Below.
800           If lngRecsCur > 0 Then
810             If .CurrentRecord > 1 Then
820               MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
830             End If
840           End If
850         End With
860       Case vbKeyDown
870         With Me
880           intRetVal = 0
890           lngRecsCur = .RecCnt  ' ** Function: Below.
900           If lngRecsCur > 0 Then
910             If .CurrentRecord < lngRecsCur Then
920               MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
930             End If
940           End If
950         End With
960       End Select
970     End If

        ' ** Alt keys.
980     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
990       Select Case intRetVal
          Case vbKeyT
1000        With Me
1010          intRetVal = 0
1020          lngRecsCur = .RecCnt  ' ** Function: Below.
1030          If lngRecsCur > 0 Then
1040            .xadload_time.SetFocus
1050          Else
1060            Beep
1070          End If
1080        End With
1090      Case vbKeyX
1100        With Me
1110          intRetVal = 0
1120          .Parent.cmdClose_Click  ' ** Form Procedure: frmLoadTimer_Mgmt.
1130        End With
1140      End Select
1150    End If

        ' ** Ctrl keys.
1160    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1170      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1180        With Me
1190          intRetVal = 0
1200          DoCmd.SelectObject acForm, .Parent.Name, False
1210          .Parent.chkLatestOnly.SetFocus
1220        End With
1230      End Select
1240    End If

        ' ** Ctrl-Shift keys.
1250    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1260      Select Case intRetVal
          Case vbKeyF
1270        With Me
1280          intRetVal = 0
1290          DoCmd.SelectObject acForm, .Parent.Name, False
1300          .Parent.FocusHolder.SetFocus
1310        End With
1320      Case vbKeyTab, vbKeyReturn
1330        With Me
1340          intRetVal = 0
1350          DoCmd.SelectObject acForm, .Parent.Name, False
1360          .Parent.cmdClose.SetFocus
1370        End With
1380      End Select
1390    End If

EXITP:
1400    KeyCode = intRetVal
1410    Exit Sub

ERRH:
1420    intRetVal = 0
1430    Select Case ERR.Number
        Case Else
1440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1450    End Select
1460    Resume EXITP

End Sub

Private Sub dbs_name1_lbl_DblClick(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "dbs_name1_lbl_DblClick"

1510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1520    Exit Sub

ERRH:
1530    Select Case ERR.Number
        Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Sub

Private Sub dbs_name1_GotFocus()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "dbs_name1_GotFocus"

1610    With Me
1620      .dbs_name1.SelLength = 0
1630      .dbs_name1.SelStart = 19
1640    End With

EXITP:
1650    Exit Sub

ERRH:
1660    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
1670    Case Else
1680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1690    End Select
1700    Resume EXITP

End Sub

Private Sub frm_name1_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "frm_name1_lbl_DblClick"

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

Private Sub frm_name1_GotFocus()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "frm_name1_GotFocus"

1910    With Me
1920      .frm_name1.SelLength = 0
1930      .frm_name1.SelStart = 99
1940    End With

EXITP:
1950    Exit Sub

ERRH:
1960    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
1970    Case Else
1980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1990    End Select
2000    Resume EXITP

End Sub

Private Sub frm_name2_lbl_DblClick(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "frm_name2_lbl_DblClick"

2110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub frm_name2_GotFocus()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "frm_name2_GotFocus"

2210    With Me
2220      .frm_name2.SelLength = 0
2230      .frm_name2.SelStart = 99
2240    End With

EXITP:
2250    Exit Sub

ERRH:
2260    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
2270    Case Else
2280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2290    End Select
2300    Resume EXITP

End Sub

Private Sub xadload_start_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "xadload_start_lbl_DblClick"

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

Private Sub xadload_start_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "xadload_start_GotFocus"

2510    With Me
2520      .xadload_start.SelLength = 0
2530      .xadload_start.SelStart = 19
2540    End With

EXITP:
2550    Exit Sub

ERRH:
2560    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
2570    Case Else
2580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2590    End Select
2600    Resume EXITP

End Sub

Private Sub xadload_end_lbl_DblClick(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "xadload_end_lbl_DblClick"

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

Private Sub xadload_end_GotFocus()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "xadload_end_GotFocus"

2810    With Me
2820      .xadload_end.SelLength = 0
2830      .xadload_end.SelStart = 19
2840    End With

EXITP:
2850    Exit Sub

ERRH:
2860    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
2870    Case Else
2880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2890    End Select
2900    Resume EXITP

End Sub

Private Sub xadload_time_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "xadload_time_lbl_DblClick"

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

Private Sub xadload_time_GotFocus()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "xadload_time_GotFocus"

3110    With Me
3120      .xadload_time.SelLength = 0
3130      .xadload_time.SelStart = 19
3140    End With

EXITP:
3150    Exit Sub

ERRH:
3160    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
3170    Case Else
3180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3190    End Select
3200    Resume EXITP

End Sub

Private Sub xadload_access_version_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "xadload_access_version_lbl_DblClick"

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

Private Sub xadload_access_version_GotFocus()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "xadload_access_version_GotFocus"

3410    With Me
3420      .xadload_access_version.SelLength = 0
3430      .xadload_access_version.SelStart = 99
3440    End With

EXITP:
3450    Exit Sub

ERRH:
3460    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
3470    Case Else
3480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3490    End Select
3500    Resume EXITP

End Sub

Private Sub xadload_datemodified_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "xadload_datemodified_lbl_DblClick"

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

Public Function RecCnt() As Long

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

3710    lngRetVal = 0&

3720    With Me
3730      Set rst = .RecordsetClone
3740      With rst
3750        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
3760        Else
3770          .MoveLast
3780          lngRetVal = .RecordCount
3790        End If
3800        .Close
3810      End With
3820    End With

EXITP:
3830    Set rst = Nothing
3840    RecCnt = lngRetVal
3850    Exit Function

ERRH:
3860    lngRetVal = 0&
3870    Select Case ERR.Number
        Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

4010    Select Case intWhere
        Case acCmdRecordsGoToFirst
4020      DoCmd.RunCommand acCmdRecordsGoToFirst
4030    Case acCmdRecordsGoToLast
4040      DoCmd.RunCommand acCmdRecordsGoToLast
4050    Case acCmdRecordsGoToPrevious
4060      DoCmd.RunCommand acCmdRecordsGoToPrevious
4070    Case acCmdRecordsGoToNext
4080      DoCmd.RunCommand acCmdRecordsGoToNext
4090    Case Else
4100      If IsMissing(varID) = False Then
4110        With Me
4120          Set rst = .RecordsetClone
4130          With rst
4140            .FindFirst "[xadload_id] = " & CStr(varID)
4150            If .NoMatch = False Then
4160              Me.Bookmark = .Bookmark
4170            End If
4180            .Close
4190          End With
4200        End With
4210      End If
4220    End Select

EXITP:
4230    Set rst = Nothing
4240    Exit Sub

ERRH:
4250    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4260    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4270    Case Else
4280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4290    End Select
4300    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [frm_name1], [frm_name2], [xadload_datemodified] DESC"

4410    With Me
4420      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
4430        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4440      End If
4450      .Controls(strSortLbl).Visible = False
4460      .Controls(strSortLine).Visible = False
4470      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the the right side of the form.
4480      If strProc = "Form_Load" Then
4490        strCalled = "frm_name1"
4500        strSortNow = strSortOrig
4510        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
4520        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4530        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4540        .Controls(strSortLbl).Left = lngSortLbl_Left
4550        .Controls(strSortLine).Left = lngSortLine_Left
4560        .Controls(strSortLine).Width = lngSortLine_Width
4570        .Controls(strSortLbl).Caption = strArwUp
4580        .Controls(strSortLbl).ForeColor = CLR_DKBLU
4590      Else
4600        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
4610        lngSortLbl_Left = (.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width
4620        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4630        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4640        .Controls(strSortLbl).Left = lngSortLbl_Left
4650        .Controls(strSortLine).Left = lngSortLine_Left
4660        .Controls(strSortLine).Width = lngSortLine_Width
4670        If strCalled = "frm_name1" Then
4680          strSortAsc = strSortOrig
4690          If strSortNow = strSortAsc Then
4700            strSortNow = "[JournalType_Order] DESC"
4710            .Controls(strSortLbl).Caption = strArwDn
4720            .Controls(strSortLbl).ForeColor = CLR_DKRED
4730          Else
4740            strSortNow = strSortAsc
4750            .Controls(strSortLbl).Caption = strArwUp
4760            .Controls(strSortLbl).ForeColor = CLR_DKBLU
4770          End If
4780        Else
4790          .Controls(strSortLbl).Caption = strArwUp
4800          .Controls(strSortLbl).ForeColor = CLR_DKBLU
4810          Select Case strCalled
              Case "dbs_name1"
4820            strSortAsc = "[dbs_name1]" & strStdAsc
4830          Case "frm_name2"
4840            strSortAsc = "[frm_name2]" & strStdAsc
4850          Case "xadload_start"
4860            strSortAsc = "[xadload_start]" & strStdAsc
4870          Case "xadload_end"
4880            strSortAsc = "[xadload_end]" & strStdAsc
4890          Case "xadload_time"
4900            strSortAsc = "[xadload_time]" & strStdAsc
4910            lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
4920            .Controls(strSortLbl).Left = lngSortLbl_Left
4930          Case "xadload_access_version"
4940            strSortAsc = "[xadload_access_version]" & strStdAsc
4950            lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
4960            .Controls(strSortLbl).Left = lngSortLbl_Left
4970          Case "xadload_datemodified"
4980            strSortAsc = "[xadload_datemodified]" & strStdAsc
4990          End Select
5000          If strSortNow = strSortAsc Then
5010            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
5020            Select Case intCnt
                Case 1
5030              strTmp01 = strSortAsc & " DESC"
5040            Case 2
5050              intPos01 = InStr(strSortAsc, ",")
5060              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5070              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
5080            Case 3
5090              intPos01 = InStr(strSortAsc, ",")
5100              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5110              strTmp02 = Mid(strSortAsc, intPos01)
5120              intPos01 = InStr(5, strTmp02, ",")
5130              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5140              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
5150            Case 4
5160              intPos01 = InStr(strSortAsc, ",")
5170              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5180              strTmp02 = Mid(strSortAsc, intPos01)
5190              intPos01 = InStr(5, strTmp02, ",")
5200              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5210              strTmp02 = Mid(strTmp02, intPos01)
5220              intPos01 = InStr(5, strTmp02, ",")
5230              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5240              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
5250            End Select
5260            strSortNow = strTmp01
5270            .Controls(strSortLbl).Caption = strArwDn
5280            .Controls(strSortLbl).ForeColor = CLR_DKRED
5290          Else
5300            strSortNow = strSortAsc
5310          End If
5320        End If
5330      End If
5340      .Controls(strSortLbl).Visible = True
5350      .Controls(strSortLine).Visible = True
5360      .OrderBy = strSortNow
5370      .OrderByOn = True
5380    End With

EXITP:
5390    Exit Sub

ERRH:
5400    Select Case ERR.Number
        Case Else
5410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5420    End Select
5430    Resume EXITP

End Sub

Public Sub ColHdrsSet(blnAble As Boolean)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "ColHdrsSet"

5510    With Me
5520      .FocusHolder2.SetFocus
5530      Select Case blnAble
          Case True
5540        .dbs_name1_lbl.ForeColor = CLR_DKGRY2
5550        .frm_name1_lbl.ForeColor = CLR_DKGRY2
5560        .frm_name2_lbl.ForeColor = CLR_DKGRY2
5570        .xadload_start_lbl.ForeColor = CLR_DKGRY2
5580        .xadload_end_lbl.ForeColor = CLR_DKGRY2
5590        .xadload_time_lbl.ForeColor = CLR_DKGRY2
5600        .xadload_access_version_lbl.ForeColor = CLR_DKGRY2
5610        .xadload_datemodified_lbl.ForeColor = CLR_DKGRY2
5620        .dbs_name1_lbl_line.BorderColor = CLR_DKGRY
5630        .frm_name1_lbl_line.BorderColor = CLR_DKGRY
5640        .frm_name2_lbl_line.BorderColor = CLR_DKGRY
5650        .xadload_start_lbl_line.BorderColor = CLR_DKGRY
5660        .xadload_end_lbl_line.BorderColor = CLR_DKGRY
5670        .xadload_time_lbl_line.BorderColor = CLR_DKGRY
5680        .xadload_access_version_lbl_line.BorderColor = CLR_DKGRY
5690        .xadload_datemodified_lbl_line.BorderColor = CLR_DKGRY
5700        .dbs_name1_lbl_dim_hi.Visible = False
5710        .frm_name1_lbl_dim_hi.Visible = False
5720        .frm_name2_lbl_dim_hi.Visible = False
5730        .xadload_start_lbl_dim_hi.Visible = False
5740        .xadload_end_lbl_dim_hi.Visible = False
5750        .xadload_time_lbl_dim_hi.Visible = False
5760        .xadload_access_version_lbl_dim_hi.Visible = False
5770        .xadload_datemodified_lbl_dim_hi.Visible = False
5780        .dbs_name1_lbl_line_dim_hi.Visible = False
5790        .frm_name1_lbl_line_dim_hi.Visible = False
5800        .frm_name2_lbl_line_dim_hi.Visible = False
5810        .xadload_start_lbl_line_dim_hi.Visible = False
5820        .xadload_end_lbl_line_dim_hi.Visible = False
5830        .xadload_time_lbl_line_dim_hi.Visible = False
5840        .xadload_access_version_lbl_line_dim_hi.Visible = False
5850        .xadload_datemodified_lbl_line_dim_hi.Visible = False
5860      Case False
5870        .dbs_name1_lbl.ForeColor = WIN_CLR_DISF
5880        .frm_name1_lbl.ForeColor = WIN_CLR_DISF
5890        .frm_name2_lbl.ForeColor = WIN_CLR_DISF
5900        .xadload_start_lbl.ForeColor = WIN_CLR_DISF
5910        .xadload_end_lbl.ForeColor = WIN_CLR_DISF
5920        .xadload_time_lbl.ForeColor = WIN_CLR_DISF
5930        .xadload_access_version_lbl.ForeColor = WIN_CLR_DISF
5940        .xadload_datemodified_lbl.ForeColor = WIN_CLR_DISF
5950        .dbs_name1_lbl_line.BorderColor = WIN_CLR_DISR
5960        .frm_name1_lbl_line.BorderColor = WIN_CLR_DISR
5970        .frm_name2_lbl_line.BorderColor = WIN_CLR_DISR
5980        .xadload_start_lbl_line.BorderColor = WIN_CLR_DISR
5990        .xadload_end_lbl_line.BorderColor = WIN_CLR_DISR
6000        .xadload_time_lbl_line.BorderColor = WIN_CLR_DISR
6010        .xadload_access_version_lbl_line.BorderColor = WIN_CLR_DISR
6020        .xadload_datemodified_lbl_line.BorderColor = WIN_CLR_DISR
6030        .dbs_name1_lbl_dim_hi.Visible = True
6040        .frm_name1_lbl_dim_hi.Visible = True
6050        .frm_name2_lbl_dim_hi.Visible = True
6060        .xadload_start_lbl_dim_hi.Visible = True
6070        .xadload_end_lbl_dim_hi.Visible = True
6080        .xadload_time_lbl_dim_hi.Visible = True
6090        .xadload_access_version_lbl_dim_hi.Visible = True
6100        .xadload_datemodified_lbl_dim_hi.Visible = True
6110        .dbs_name1_lbl_line_dim_hi.Visible = True
6120        .frm_name1_lbl_line_dim_hi.Visible = True
6130        .frm_name2_lbl_line_dim_hi.Visible = True
6140        .xadload_start_lbl_line_dim_hi.Visible = True
6150        .xadload_end_lbl_line_dim_hi.Visible = True
6160        .xadload_time_lbl_line_dim_hi.Visible = True
6170        .xadload_access_version_lbl_line_dim_hi.Visible = True
6180        .xadload_datemodified_lbl_line_dim_hi.Visible = True
6190      End Select
6200    End With

EXITP:
6210    Exit Sub

ERRH:
6220    Select Case ERR.Number
        Case Else
6230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6240    End Select
6250    Resume EXITP

End Sub
