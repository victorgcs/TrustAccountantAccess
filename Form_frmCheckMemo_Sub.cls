VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckMemo_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckMemo_Sub"

'VGC 07/01/2015: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd on frmCheckMemo}
' **   Close:            C {cmdClose on frmCheckMemo}
' **   Delete:           D {cmdDelete on frmCheckMemo}
' **   Memo:             M {ChkMemo_Memo}
' **   Exit:             X {cmdClose on frmCheckMemo}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private Const strSortOrig As String = "[ChkMemo_Memo]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngRecsCur As Long, lngTpp As Long
Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmCheckMemo" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
170         lngSortLbl_Width = .Sort_lbl.Width

180       Else
190         Cancel = -1
200       End If
210     End With

220     If Cancel = -1 Then
230       Beep
240       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
250     End If

EXITP:
260     Exit Sub

ERRH:
270     Select Case ERR.Number
        Case Else
280       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
290     End Select
300     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality.
        ' **   Memo:             M {ChkMemo_Memo}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality.
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyDown
570         intRetVal = 0
580         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
590       Case vbKeyUp
600         intRetVal = 0
610         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
620       Case vbKeyEscape
630         With Me
640           intRetVal = 0
650           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountComments.
660         End With
670       End Select
680     End If

        ' ** Alt keys.
690     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
700       Select Case intRetVal
          Case vbKeyM
710         With Me
720           intRetVal = 0
730           .ChkMemo_Memo.SetFocus
740         End With
750       Case vbKeyX
760         With Me
770           intRetVal = 0
780           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountComments.
790         End With
800       End Select
810     End If

        ' ** Ctrl keys.
820     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
830       Select Case intRetVal
          Case vbKeyS
840         intRetVal = 0
850         cmdSave_Click  ' ** Procedure: Below.
860       Case vbKeyTab
870         With Me
880           intRetVal = 0
890           DoCmd.SelectObject acForm, .Parent.Name, False
900           If .Parent.cmdAdd.Enabled = True Then
910             .Parent.cmdAdd.SetFocus
920           ElseIf .Parent.cmdDelete.Enabled = True Then
930             .Parent.cmdDelete.SetFocus
940           Else
950             .Parent.cmdClose.SetFocus
960           End If
970         End With
980       Case vbKeyPageUp, vbKeyHome
990         intRetVal = 0
1000        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1010      Case vbKeyPageDown, vbKeyEnd
1020        intRetVal = 0
1030        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1040      End Select
1050    End If

        ' ** Ctrl-Shift keys.
1060    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1070      Select Case intRetVal
          Case vbKeyTab
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
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
1190    Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

1310    With Me
1320      .Username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1330      .ChkMemo_DateModified = Now()
1340    End With

EXITP:
1350    Exit Sub

ERRH:
1360    Select Case ERR.Number
        Case Else
1370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1380    End Select
1390    Resume EXITP

End Sub

Private Sub cmdSave_Click()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1410    DoCmd.RunCommand acCmdSaveRecord
1420    DoEvents

EXITP:
1430    Exit Sub

ERRH:
1440    Select Case ERR.Number
        Case Else
1450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1460    End Select
1470    Resume EXITP

End Sub

Private Sub ChkMemoType_Type_lbl_DblClick(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkMemoType_Type_lbl_DblClick"

1510    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
1520    Exit Sub

ERRH:
1530    Select Case ERR.Number
        Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Sub

Private Sub ChkMemoType_Type_GotFocus()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkMemoType_Type_GotFocus"

1610    With Me
1620      .ChkMemoType_Type.SelLength = 0
1630      .ChkMemoType_Type.SelStart = 99
1640    End With

EXITP:
1650    Exit Sub

ERRH:
1660    Select Case ERR.Number
        Case Else
1670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1680    End Select
1690    Resume EXITP

End Sub

Private Sub ChkMemoType_Type_KeyDown(KeyCode As Integer, Shift As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkMemoType_Type_KeyDown"

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
1790          .ChkMemo_Memo.SetFocus
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
1890            .ChkMemo_Memo.SetFocus
1900          Else
1910            DoCmd.SelectObject acForm, .Parent.Name, False
1920            .Parent.cmdClose.SetFocus
1930          End If
1940        End With
1950      End Select
1960    End If

EXITP:
1970    KeyCode = intRetVal
1980    Exit Sub

ERRH:
1990    intRetVal = 0
2000    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2010    Case Else
2020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2030    End Select
2040    Resume EXITP

End Sub

Private Sub ChkMemo_Memo_lbl_DblClick(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkMemo_Memo_lbl_DblClick"

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

Public Sub ChkMemo_Memo_GotFocus()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkMemo_Memo_GotFocus"

2210    With Me
2220      .ChkMemo_Memo.SelLength = 0
2230      .ChkMemo_Memo.SelStart = 99
2240    End With

EXITP:
2250    Exit Sub

ERRH:
2260    Select Case ERR.Number
        Case Else
2270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2280    End Select
2290    Resume EXITP

End Sub

Private Sub ChkMemo_Memo_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkMemo_Memo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2320    intShiftDown = (Shift And acShiftMask) > 0
2330    intAltDown = (Shift And acAltMask) > 0
2340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2370        With Me
2380          intRetVal = 0
2390          lngRecsCur = RecCnt  ' ** Function: Below.
2400          If .CurrentRecord < lngRecsCur Then
2410            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2420            .ChkMemoType_Type.SetFocus
2430          Else
2440            DoCmd.SelectObject acForm, .Parent.Name, False
2450            If .Parent.cmdAdd.Enabled = True Then
2460              .Parent.cmdAdd.SetFocus
2470            ElseIf .Parent.cmdDelete.Enabled = True Then
2480              .Parent.cmdDelete.SetFocus
2490            Else
2500              .Parent.cmdClose.SetFocus
2510            End If
2520          End If
2530        End With
2540      End Select
2550    End If

        ' ** Shift keys.
2560    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2570      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2580        With Me
2590          intRetVal = 0
2600          .ChkMemoType_Type.SetFocus
2610        End With
2620      End Select
2630    End If

EXITP:
2640    KeyCode = intRetVal
2650    Exit Sub

ERRH:
2660    intRetVal = 0
2670    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2680    Case Else
2690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2700    End Select
2710    Resume EXITP

End Sub

Public Function RecCnt() As Long

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

2810    lngRetVal = 0&

2820    With Me
2830      Set rst = .RecordsetClone
2840      With rst
2850        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
2860        Else
2870          .MoveLast
2880          lngRetVal = .RecordCount
2890        End If
2900        .Close
2910      End With
2920    End With

EXITP:
2930    Set rst = Nothing
2940    RecCnt = lngRetVal
2950    Exit Function

ERRH:
2960    lngRetVal = 0&
2970    Select Case ERR.Number
        Case Else
2980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2990    End Select
3000    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

3110    Select Case intWhere
        Case acCmdRecordsGoToFirst
3120      DoCmd.RunCommand acCmdRecordsGoToFirst
3130    Case acCmdRecordsGoToLast
3140      DoCmd.RunCommand acCmdRecordsGoToLast
3150    Case acCmdRecordsGoToPrevious
3160      DoCmd.RunCommand acCmdRecordsGoToPrevious
3170    Case acCmdRecordsGoToNext
3180      DoCmd.RunCommand acCmdRecordsGoToNext
3190    Case Else
3200      If IsMissing(varID) = False Then
3210        With Me
3220          Set rst = .RecordsetClone
3230          With rst
3240            .FindFirst "[ChkMemo_ID] = " & CStr(varID)
3250            If .NoMatch = False Then
3260              Me.Bookmark = .Bookmark
3270            End If
3280            .Close
3290          End With
3300        End With
3310      End If
3320    End Select

EXITP:
3330    Set rst = Nothing
3340    Exit Sub

ERRH:
3350    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
3360    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
3370    Case Else
3380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3390    End Select
3400    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer

        Const strStdAsc As String = ", [ChkMemo_Memo]"

3510    With Me
3520      .Controls(strSortLbl).Visible = False
3530      .Controls(strSortLine).Visible = False
3540      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
3550      If strProc = "Form_Load" Then
3560        strCalled = "ChkMemo_Memo"
3570        strSortNow = strSortOrig
3580        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - (lngSortLbl_Width))
3590        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
3600        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
3610        .Controls(strSortLbl).Left = lngSortLbl_Left
3620        .Controls(strSortLbl).Caption = strArwUp
3630        .Controls(strSortLine).Left = lngSortLine_Left
3640        .Controls(strSortLine).Width = lngSortLine_Width
3650        .Controls(strSortLbl).ForeColor = CLR_DKBLU
3660      Else
3670        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
3680        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - (lngSortLbl_Width))
3690        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
3700        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
3710        .Controls(strSortLbl).Left = lngSortLbl_Left
3720        If strCalled = "ChkMemo_Memo" Then
3730          strSortAsc = strSortOrig
3740          If strSortNow = strSortAsc Then
3750            strSortNow = "[ChkMemo_Memo] DESC"
3760            .Controls(strSortLbl).Caption = strArwDn
3770            .Controls(strSortLbl).ForeColor = CLR_DKRED
3780          Else
3790            strSortNow = strSortAsc
3800            .Controls(strSortLbl).Caption = strArwUp
3810            .Controls(strSortLbl).ForeColor = CLR_DKBLU
3820          End If
3830          .Controls(strSortLine).Left = lngSortLine_Left
3840          .Controls(strSortLine).Width = lngSortLine_Width
3850        Else
3860          .Controls(strSortLbl).Caption = strArwUp
3870          .Controls(strSortLbl).ForeColor = CLR_DKBLU
3880          Select Case strCalled
              Case "ChkMemoType_Type"
3890            strSortAsc = "[ChkMemoType_Type]" & strStdAsc
3900          End Select
3910          .Controls(strSortLine).Left = lngSortLine_Left
3920          .Controls(strSortLine).Width = lngSortLine_Width
3930          If strSortNow = strSortAsc Then
3940            intPos01 = InStr(strSortAsc, ",")
3950            If intPos01 = 0 Then
                  ' ** [accounttype_description] DESC.
3960              strSortNow = strSortNow & " DESC"
3970            Else
                  ' ** [accounttypegroup_name] DESC, [accounttype] DESC.
3980              strSortNow = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
3990            End If
4000            .Controls(strSortLbl).Caption = strArwDn
4010            .Controls(strSortLbl).ForeColor = CLR_DKRED
4020          Else
4030            strSortNow = strSortAsc
4040          End If
4050        End If
4060      End If
4070      .Controls(strSortLbl).Visible = True
4080      .Controls(strSortLine).Visible = True
4090      .OrderBy = strSortNow
4100      .OrderByOn = True
4110    End With

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub
