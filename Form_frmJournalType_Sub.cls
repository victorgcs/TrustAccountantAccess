VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournalType_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournalType_Sub"

'VGC 04/19/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmJournalType}
' **   Print Report:     P {cmdPrintReport on frmJournalType}
' **   Exit:             X {cmdClose on frmJournalType}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[JournalType_Order]"
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
150       If strCallingForm = "frmJournalType" Then

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
420       .journaltype.ForeColor = CLR_DISABLED_FG
430       .journaltype.BackColor = CLR_DISABLED_BG
440       .JournalType_Description.ForeColor = CLR_DISABLED_FG
450       .JournalType_Description.BackColor = CLR_DISABLED_BG
460       .JournalType_Order.ForeColor = CLR_DISABLED_FG
470       .JournalType_Order.BackColor = CLR_DISABLED_BG
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
750           .Parent.cmdClose_Click  ' ** Form Procedure: frmJournalType.
760         End With
770       End Select
780     End If

        ' ** Alt keys.
790     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
800       Select Case intRetVal
          Case vbKeyX
810         With Me
820           intRetVal = 0
830           .Parent.cmdClose_Click  ' ** Form Procedure: frmJournalType.
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
950             .Parent.cmdClose.SetFocus
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
1480              .Parent.cmdClose.SetFocus
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

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_lbl_DblClick"

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

Private Sub journaltype_GotFocus()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_GotFocus"

1910    With Me
1920      .journaltype.SelLength = 0
1930      .journaltype.SelStart = 19
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

Private Sub JournalType_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_KeyDown"

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
2090          .JournalType_Description.SetFocus
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

Private Sub JournalType_Description_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_Description_lbl_DblClick"

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

Private Sub JournalType_description_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_description_GotFocus"

2510    With Me
2520      .JournalType_Description.SelLength = 0
2530      .JournalType_Description.SelStart = 99
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

Private Sub JournalType_description_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_description_KeyDown"

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
2690          .JournalType_Order.SetFocus
2700        End With
2710      End Select
2720    End If

        ' ** Shift keys.
2730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2750        With Me
2760          intRetVal = 0
2770          .journaltype.SetFocus
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

Private Sub JournalType_Order_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_Order_lbl_DblClick"

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

Private Sub JournalType_Order_GotFocus()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_Order_GotFocus"

3010    With Me
3020      .JournalType_Order.SelLength = 0
3030      .JournalType_Order.SelStart = 9
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

Private Sub JournalType_Order_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_Order_KeyDown"

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
3190          lngRecsCur = RecCnt  ' ** Function: Below
3200          If .CurrentRecord < lngRecsCur Then
3210            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3220            .FocusHolder.SetFocus
3230          Else
3240            DoCmd.SelectObject acForm, .Parent.Name
3250            If .Parent.cmdPrintReport.Enabled = True Then
3260              .Parent.cmdPrintReport.SetFocus
3270            Else
3280              .Parent.cmdClose.SetFocus
3290            End If
3300          End If
3310        End With
3320      End Select
3330    End If

        ' ** Shift keys.
3340    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3350      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3360        With Me
3370          intRetVal = 0
3380          .JournalType_Description.SetFocus
3390        End With
3400      End Select
3410    End If

EXITP:
3420    KeyCode = intRetVal
3430    Exit Sub

ERRH:
3440    intRetVal = 0
3450    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3460    Case Else
3470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3480    End Select
3490    Resume EXITP

End Sub

Public Function RecCnt() As Long

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

3510    lngRetVal = 0&

3520    With Me
3530      Set rst = .RecordsetClone
3540      With rst
3550        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
3560        Else
3570          .MoveLast
3580          lngRetVal = .RecordCount
3590        End If
3600        .Close
3610      End With
3620    End With

EXITP:
3630    Set rst = Nothing
3640    RecCnt = lngRetVal
3650    Exit Function

ERRH:
3660    lngRetVal = 0&
3670    Select Case ERR.Number
        Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

3810    Select Case intWhere
        Case acCmdRecordsGoToFirst
3820      DoCmd.RunCommand acCmdRecordsGoToFirst
3830    Case acCmdRecordsGoToLast
3840      DoCmd.RunCommand acCmdRecordsGoToLast
3850    Case acCmdRecordsGoToPrevious
3860      DoCmd.RunCommand acCmdRecordsGoToPrevious
3870    Case acCmdRecordsGoToNext
3880      DoCmd.RunCommand acCmdRecordsGoToNext
3890    Case Else
3900      If IsMissing(varID) = False Then
3910        With Me
3920          Set rst = .RecordsetClone
3930          With rst
3940            .FindFirst "[JournalType_ID] = " & CStr(varID)
3950            If .NoMatch = False Then
3960              Me.Bookmark = .Bookmark
3970            End If
3980            .Close
3990          End With
4000        End With
4010      End If
4020    End Select

EXITP:
4030    Set rst = Nothing
4040    Exit Sub

ERRH:
4050    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4060    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4070    Case Else
4080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4090    End Select
4100    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String

        'Const strStdAsc As String = ", [JournalType_Order]"

4210    With Me
4220      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
4230        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4240      End If
4250      .Controls(strSortLbl).Visible = False
4260      .Controls(strSortLine).Visible = False
4270      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the the right side of the form.
4280      If strProc = "Form_Load" Then
4290        strCalled = "JournalType_Order"
4300        strSortNow = strSortOrig
4310        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4320        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (3& * lngTpp))
4330        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4340        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4350        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4360        .Controls(strSortLbl).Top = lngSortLbl_Top
4370        .Controls(strSortLbl).Left = lngSortLbl_Left
4380        .Controls(strSortLine).Top = lngSortLine_Top
4390        .Controls(strSortLine).Left = lngSortLine_Left
4400        .Controls(strSortLine).Width = lngSortLine_Width
4410        .Controls(strSortLbl).Caption = strArwUp
4420        .Controls(strSortLbl).ForeColor = CLR_DKBLU
4430      Else
4440        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
4450        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4460        lngSortLbl_Left = (.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width
4470        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4480        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4490        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4500        .Controls(strSortLbl).Top = lngSortLbl_Top
4510        .Controls(strSortLbl).Left = lngSortLbl_Left
4520        .Controls(strSortLine).Top = lngSortLine_Top
4530        .Controls(strSortLine).Left = lngSortLine_Left
4540        .Controls(strSortLine).Width = lngSortLine_Width
4550        If strCalled = "JournalType_Order" Then
4560          lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
4570          .Controls(strSortLbl).Left = lngSortLbl_Left
4580          strSortAsc = strSortOrig
4590          If strSortNow = strSortAsc Then
4600            strSortNow = "[JournalType_Order] DESC"
4610            .Controls(strSortLbl).Caption = strArwDn
4620            .Controls(strSortLbl).ForeColor = CLR_DKRED
4630          Else
4640            strSortNow = strSortAsc
4650            .Controls(strSortLbl).Caption = strArwUp
4660            .Controls(strSortLbl).ForeColor = CLR_DKBLU
4670          End If
4680        Else
4690          .Controls(strSortLbl).Caption = strArwUp
4700          .Controls(strSortLbl).ForeColor = CLR_DKBLU
4710          Select Case strCalled
              Case "JournalType"
4720            strSortAsc = "[JournalType]"
4730          Case "JournalType_Description"
4740            strSortAsc = "[JournalType_Description]"
4750          End Select
4760          If strSortNow = strSortAsc Then
4770            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
4780            Select Case intCnt
                Case 1
4790              strTmp01 = strSortAsc & " DESC"
4800            Case 2
4810              intPos01 = InStr(strSortAsc, ",")
4820              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
4830              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
4840            End Select
4850            strSortNow = strTmp01
4860            .Controls(strSortLbl).Caption = strArwDn
4870            .Controls(strSortLbl).ForeColor = CLR_DKRED
4880          Else
4890            strSortNow = strSortAsc
4900          End If
4910        End If
4920      End If
4930      .Controls(strSortLbl).Visible = True
4940      .Controls(strSortLine).Visible = True
4950      .OrderBy = strSortNow
4960      .OrderByOn = True
4970    End With

        ' ** Potential Sorts:
        ' **   [JournalType_Order]
        ' **   [JournalType]
        ' **   [JournalType_Description]

EXITP:
4980    Exit Sub

ERRH:
4990    Select Case ERR.Number
        Case Else
5000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5010    End Select
5020    Resume EXITP

End Sub

Public Function SortNow_Get() As String

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

5110    strRetVal = strSortNow

EXITP:
5120    SortNow_Get = strRetVal
5130    Exit Function

ERRH:
5140    strRetVal = vbNullString
5150    Select Case ERR.Number
        Case Else
5160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5170    End Select
5180    Resume EXITP

End Function
