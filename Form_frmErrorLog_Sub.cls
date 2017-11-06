VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmErrorLog_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmErrorLog_Sub"

'VGC 04/17/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmErrorLog}
' **   Excel:            E {cmdExcel on frmErrorLog}
' **   Print:            P {cmdPrint on frmErrorLog}
' **   Clear All:        R {cmdClear on frmErrorLog}
' **   Preview:          W {cmdPreview on frmErrorLog}
' **   Exit:             X {cmdClose on frmErrorLog}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[ErrLog_Date] DESC"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private blnFoundErrLineNum As Boolean
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field
        Dim lngOffset As Long
        Dim blnRetVal As Boolean

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmErrorLog" Then

160         blnFoundErrLineNum = False

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

190         blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

210         .detail_height_line.Top = 0&  ' ** Just to get it out of the way.
220         lngSortLbl_Width = .Sort_lbl.Width
230         .DevMsg_lbl.Top = 0&

240         Set dbs = CurrentDb
250         With dbs
260           Set tdf = .TableDefs("tblErrorLog")
270           With tdf
280             For Each fld In .Fields
290               If fld.Name = "ErrLog_LineNum" Then
300                 blnFoundErrLineNum = True
310                 Exit For
320               End If
330             Next
340           End With
350           .Close
360         End With

370         If blnFoundErrLineNum = True Then
380           With Me
390             .RecordSource = "qryErrLog_01a"
400             .ErrLog_LineNum.Visible = True
410             .ErrLog_LineNum_lbl.Visible = True
420             .ErrLog_LineNum_lbl_line.Visible = True
430             lngOffset = (.ErrLog_LineNum.Width + 60&)
440             .Width = .Width + lngOffset
450             .ErrLog_FuncSub.Left = .ErrLog_FuncSub.Left + lngOffset
460             .ErrLog_FuncSub_lbl.Left = .ErrLog_FuncSub_lbl.Left + lngOffset
470             .ErrLog_FuncSub_lbl_line.Left = .ErrLog_FuncSub_lbl.Left
480             .ErrLog_Form.Left = .ErrLog_Form.Left + lngOffset
490             .ErrLog_Form_lbl.Left = .ErrLog_Form_lbl.Left + lngOffset
500             .ErrLog_Form_lbl_line.Left = .ErrLog_Form_lbl.Left
510           End With
520         End If

530       Else
540         Cancel = -1
550       End If
560     End With

570     If Cancel = -1 Then
580       Beep
590       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
600     End If

EXITP:
610     Set fld = Nothing
620     Set tdf = Nothing
630     Set dbs = Nothing
640     Exit Sub

ERRH:
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub Form_Load()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

710     With Me
720       lngRecsCur = RecCnt  ' ** Function: Below.
730       If lngRecsCur = 0& Then
740         .Parent.cmdPreview.Enabled = False
750         .Parent.cmdPrint.Enabled = False
760         .Parent.cmdExcel.Enabled = False
770         .Parent.cmdClear.Enabled = False
780       Else
790         .ErrLog_Date.ForeColor = CLR_DISABLED_FG
800         .ErrLog_Date.BackColor = CLR_DISABLED_BG
810         .ErrLog_ErrNum.ForeColor = CLR_DISABLED_FG
820         .ErrLog_ErrNum.BackColor = CLR_DISABLED_BG
830         .ErrLog_Message_Disp.ForeColor = CLR_DISABLED_FG
840         .ErrLog_Message_Disp.BackColor = CLR_DISABLED_BG
850         .ErrLog_LineNum.ForeColor = CLR_DISABLED_FG
860         .ErrLog_LineNum.BackColor = CLR_DISABLED_BG
870         .ErrLog_Form.ForeColor = CLR_DISABLED_FG
880         .ErrLog_Form.BackColor = CLR_DISABLED_BG
890         .ErrLog_FuncSub.ForeColor = CLR_DISABLED_FG
900         .ErrLog_FuncSub.BackColor = CLR_DISABLED_BG
910       End If
920       SortNow THIS_PROC  ' ** Procedure: Below.
930     End With

EXITP:
940     Exit Sub

ERRH:
950     Select Case ERR.Number
        Case Else
960       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
970     End Select
980     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strControl As String
        Dim intRetVal As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

1050  On Error Resume Next
1060    strControl = Screen.ActiveControl.Name
1070  On Error GoTo ERRH

        ' ** Plain keys.
1080    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1090      Select Case intRetVal
          Case vbKeyEscape
1100        With Me
1110          intRetVal = 0
1120          .Parent.cmdClose_Click  ' ** Form Procedure: frmErrorLog.
1130        End With
1140      Case vbKeyUp
1150        If strControl <> "ErrLog_Message_Disp" Then
1160          intRetVal = 0
1170          MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1180        End If
1190      Case vbKeyDown
1200        If strControl <> "ErrLog_Message_Disp" Then
1210          intRetVal = 0
1220          MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1230        End If
1240      End Select
1250    End If

        ' ** Alt keys.
1260    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1270      Select Case intRetVal
          Case vbKeyX
1280        With Me
1290          intRetVal = 0
1300          .Parent.cmdClose_Click  ' ** Form Procedure: frmErrorLog.
1310        End With
1320      End Select
1330    End If

        ' ** Ctrl keys.
1340    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1350      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1360        With Me
1370          intRetVal = 0
1380          DoCmd.SelectObject acForm, .Parent.Name, False
1390          If .Parent.cmdPreview.Enabled = True Then
1400            .Parent.cmdPreview.SetFocus
1410          ElseIf .Parent.cmdClear.Enabled = True Then
1420            .Parent.cmdClear.SetFocus
1430          Else
1440            .Parent.cmdClose.SetFocus
1450          End If
1460        End With
1470      Case vbKeyPageUp, vbKeyHome
1480        intRetVal = 0
1490        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1500      Case vbKeyPageDown, vbKeyEnd
1510        intRetVal = 0
1520        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1530      End Select
1540    End If

        ' ** Ctrl-Shift keys.
1550    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1560      Select Case intRetVal
          Case vbKeyF
1570        With Me
1580          intRetVal = 0
1590          DoCmd.SelectObject acForm, .Parent.Name, False
1600          .Parent.FocusHolder.SetFocus
1610        End With
1620      Case vbKeyTab, vbKeyReturn
1630        With Me
1640          intRetVal = 0
1650          DoCmd.SelectObject acForm, .Parent.Name, False
1660          .Parent.cmdClose.SetFocus
1670        End With
1680      End Select
1690    End If

EXITP:
1700    KeyCode = intRetVal
1710    Exit Sub

ERRH:
1720    intRetVal = 0
1730    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1740    Case Else
1750      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1760    End Select
1770    Resume EXITP

End Sub

Private Sub ErrLog_Date_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "ErrLog_Date_lbl_DblClick"

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

Private Sub ErrLog_Date_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "ErrLog_Date_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1920    intShiftDown = (Shift And acShiftMask) > 0
1930    intAltDown = (Shift And acAltMask) > 0
1940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1970        With Me
1980          intRetVal = 0
1990          .ErrLog_ErrNum.SetFocus
2000        End With
2010      End Select
2020    End If

        ' ** Shift keys.
2030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2050        With Me
2060          intRetVal = 0
2070          If .CurrentRecord > 1 Then
2080            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2090            .ErrLog_FuncSub.SetFocus
2100          Else
2110            DoCmd.SelectObject acForm, .Parent.Name, False
2120            .Parent.cmdClose.SetFocus
2130          End If
2140        End With
2150      End Select
2160    End If

EXITP:
2170    KeyCode = intRetVal
2180    Exit Sub

ERRH:
2190    intRetVal = 0
2200    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2210    Case Else
2220      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2230    End Select
2240    Resume EXITP

End Sub

Private Sub ErrLog_Date_GotFocus()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "ErrLog_Date_GotFocus"

2310    With Me
2320      .ErrLog_Date.SelLength = 0
2330      .ErrLog_Date.SelStart = 99
2340    End With

EXITP:
2350    Exit Sub

ERRH:
2360    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore
2370    Case Else
2380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2390    End Select
2400    Resume EXITP

End Sub

Private Sub ErrLog_ErrNum_lbl_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "ErrLog_ErrNum_lbl_DblClick"

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

Private Sub ErrLog_Message_Disp_lbl_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "ErrLog_Message_Disp_lbl_DblClick"

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

Private Sub ErrLog_LineNum_lbl_DblClick(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "ErrLog_LineNum_lbl_DblClick"

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

Private Sub ErrLog_Form_lbl_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "ErrLog_Form_lbl_DblClick"

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

Private Sub ErrLog_FuncSub_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "ErrLog_FuncSub_lbl_DblClick"

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

Private Sub ErrLog_FuncSub_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "ErrLog_FuncSub_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3020    intShiftDown = (Shift And acShiftMask) > 0
3030    intAltDown = (Shift And acAltMask) > 0
3040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3070        With Me
3080          intRetVal = 0
3090          lngRecsCur = RecCnt  ' ** Function: Below.
3100          If .CurrentRecord < lngRecsCur Then
3110            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3120            .FocusHolder.SetFocus
3130            DoEvents
3140            .ErrLog_Date.SetFocus
3150          Else
3160            DoCmd.SelectObject acForm, .Parent.Name, False
3170            If .Parent.cmdPreview.Enabled = True Then
3180              .Parent.cmdPreview.SetFocus
3190            Else
3200              .Parent.opgMsgLines.SetFocus
3210            End If
3220          End If
3230        End With
3240      End Select
3250    End If

        ' ** Shift keys.
3260    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3270      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3280        With Me
3290          intRetVal = 0
3300          .ErrLog_Form.SetFocus
3310        End With
3320      End Select
3330    End If

EXITP:
3340    KeyCode = intRetVal
3350    Exit Sub

ERRH:
3360    intRetVal = 0
3370    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3380    Case Else
3390      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
3400    End Select
3410    Resume EXITP

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
3940            .FindFirst "[ErrLog_ID] = " & CStr(varID)
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
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [ErrLog_Date] DESC"

4210    With Me
4220      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
4230        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4240      End If
4250      .Controls(strSortLbl).Visible = False
4260      .Controls(strSortLine).Visible = False
4270      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
4280      If strProc = "Form_Load" Then
4290        strCalled = "ErrLog_Date"
4300        strSortNow = strSortOrig
4310        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
4320        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4330        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4340        .Controls(strSortLbl).Left = lngSortLbl_Left
4350        .Controls(strSortLine).Left = lngSortLine_Left
4360        .Controls(strSortLine).Width = lngSortLine_Width
4370        .Controls(strSortLbl).Caption = strArwDn
4380        .Controls(strSortLbl).ForeColor = CLR_DKRED
4390      Else
4400        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
4410        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
4420        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4430        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4440        .Controls(strSortLbl).Left = lngSortLbl_Left
4450        .Controls(strSortLine).Left = lngSortLine_Left
4460        .Controls(strSortLine).Width = lngSortLine_Width
4470        If strCalled = "ErrLog_Date" Then
4480          strSortAsc = "[ErrLog_Date] DESC"
4490          If strSortNow = strSortAsc Then
4500            strSortNow = "[ErrLog_Date]"
4510            .Controls(strSortLbl).Caption = strArwUp
4520            .Controls(strSortLbl).ForeColor = CLR_DKBLU
4530          Else
4540            strSortNow = strSortOrig
4550            .Controls(strSortLbl).Caption = strArwDn
4560            .Controls(strSortLbl).ForeColor = CLR_DKRED
4570          End If
4580        Else
4590          .Controls(strSortLbl).Caption = strArwUp
4600          .Controls(strSortLbl).ForeColor = CLR_DKBLU
4610          Select Case strCalled
              Case "ErrLog_ErrNum"
4620            strSortAsc = "[ErrLog_ErrNum]" & strStdAsc
4630          Case "ErrLog_Message_Disp"
4640            strSortAsc = "[ErrLog_Message_Disp]" & strStdAsc
4650          Case "ErrLog_LineNum"
4660            strSortAsc = "[ErrLog_LineNum]" & strStdAsc
4670          Case "ErrLog_Form"
4680            strSortAsc = "[ErrLog_Form], [ErrLog_FuncSub], [ErrLog_ErrNum]" & strStdAsc
4690          Case "ErrLog_FuncSub"
4700            strSortAsc = "[ErrLog_FuncSub], [ErrLog_Form], [ErrLog_ErrNum]" & strStdAsc
4710          End Select
4720          If strSortNow = strSortAsc Then
4730            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
4740            Select Case intCnt
                Case 1
4750              strTmp01 = strSortAsc & " DESC"
4760            Case 2
4770              intPos01 = InStr(strSortAsc, ",")
4780              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
4790              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
4800            Case 3
4810              intPos01 = InStr(strSortAsc, ",")
4820              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
4830              strTmp02 = Mid(strSortAsc, intPos01)
4840              intPos01 = InStr(2, strTmp02, ",")
4850              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
4860              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
4870            Case 4
4880              intPos01 = InStr(strSortAsc, ",")
4890              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
4900              strTmp02 = Mid(strSortAsc, intPos01)
4910              intPos01 = InStr(2, strTmp02, ",")
4920              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
4930              strTmp02 = Mid(strTmp02, intPos01)
4940              intPos01 = InStr(2, strTmp02, ",")
4950              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
4960              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
4970            End Select
4980            strSortNow = strTmp01
4990            .Controls(strSortLbl).Caption = strArwDn
5000            .Controls(strSortLbl).ForeColor = CLR_DKRED
5010          Else
5020            strSortNow = strSortAsc
5030          End If
5040        End If
5050      End If
5060      .Controls(strSortLbl).Visible = True
5070      .Controls(strSortLine).Visible = True
5080      .OrderBy = strSortNow
5090      .OrderByOn = True
5100    End With

        ' ** Potential Sorts:
        ' **   [ErrLog_Date] DESC
        ' **   [ErrLog_ErrNum], [ErrLog_Date] DESC
        ' **   [ErrLog_Message_Disp], [ErrLog_Date] DESC
        ' **   [ErrLog_LineNum], [ErrLog_Date] DESC
        ' **   [ErrLog_Form], [ErrLog_FuncSub], [ErrLog_ErrNum], [ErrLog_Date] DESC
        ' **   [ErrLog_FuncSub], [ErrLog_Form], [ErrLog_ErrNum], [ErrLog_Date] DESC

EXITP:
5110    Exit Sub

ERRH:
5120    Select Case ERR.Number
        Case Else
5130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5140    End Select
5150    Resume EXITP

End Sub

Public Function SortNow_Get() As String

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

5210    strRetVal = strSortNow

EXITP:
5220    SortNow_Get = strRetVal
5230    Exit Function

ERRH:
5240    strRetVal = vbNullString
5250    Select Case ERR.Number
        Case Else
5260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5270    End Select
5280    Resume EXITP

End Function
