VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountComments_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountComments_Sub"

'VGC 04/15/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmAccountComments}
' **   Include Archive:  H {chkIncludeArchive on frmAccountComments}
' **   Include Currency: L {chkIncludeCurrency on frmAccountComments}
' **   Update:           U {cmdUpdate on frmAccountComments}
' **   Exit:             X {cmdClose on frmAccountComments}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[transdate] DESC, [journaltype], [journalno] DESC"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngForm_Width As Long, lngCurrID_Left As Long, lngCurrID_Width As Long, lngCurrIDLbl_Offset As Long

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
150       If strCallingForm = "frmAccountComments" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         .form_width_line.Left = 0&  ' ** Just get it out of the way.

200         lngSortLbl_Width = .Sort_lbl.Width
210         lngForm_Width = .Width
220         lngCurrID_Left = .curr_id.Left
230         lngCurrID_Width = (.curr_id.Left + .curr_id.Width) - (.Cost.Left + .Cost.Width)
240         lngCurrIDLbl_Offset = .curr_id.Left - .curr_id_lbl.Left

250       Else
260         Cancel = -1
270       End If
280     End With

290     If Cancel = -1 Then
300       Beep
310       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
320     End If

EXITP:
330     Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     SortNow THIS_PROC  ' ** Procedure: Below

420     With Me
430       .transdate.ForeColor = CLR_DISABLED_FG
440       .transdate.BackColor = CLR_DISABLED_BG
450       .journaltype.ForeColor = CLR_DISABLED_FG
460       .journaltype.BackColor = CLR_DISABLED_BG
470       .ICash.ForeColor = CLR_DISABLED_FG
480       .ICash.BackColor = CLR_DISABLED_BG
490       .PCash.ForeColor = CLR_DISABLED_FG
500       .PCash.BackColor = CLR_DISABLED_BG
510       .Cost.ForeColor = CLR_DISABLED_FG
520       .Cost.BackColor = CLR_DISABLED_BG
530       .curr_id.ForeColor = CLR_DISABLED_FG
540       .curr_id.BackColor = CLR_DISABLED_BG
550       .Comment.SetFocus
560     End With

EXITP:
570     Exit Sub

ERRH:
580     Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
590     Case Else
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

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyDown
770         intRetVal = 0
780         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
790       Case vbKeyUp
800         intRetVal = 0
810         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
820       Case vbKeyEscape
830         With Me
840           intRetVal = 0
850           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountComments.
860         End With
870       End Select
880     End If

        ' ** Alt keys.
890     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
900       Select Case intRetVal
          Case vbKeyX
910         With Me
920           intRetVal = 0
930           .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountComments.
940         End With
950       End Select
960     End If

        ' ** Ctrl keys.
970     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
980       Select Case intRetVal
          Case vbKeyS
990         intRetVal = 0
1000        cmdSave_Click  ' ** Procedure: Below.
1010      Case vbKeyTab
1020        With Me
1030          intRetVal = 0
1040          DoCmd.SelectObject acForm, .Parent.Name, False
1050          If .Parent.chkIncludeArchive.Enabled = True Then
1060            .Parent.chkIncludeArchive.SetFocus
1070          ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
1080            .Parent.chkIncludeCurrency.SetFocus
1090          ElseIf .Parent.cmdUpdate.Enabled = True Then
1100            .Parent.cmdUpdate.SetFocus
1110          Else
1120            .Parent.cmdClose.SetFocus
1130          End If
1140        End With
1150      Case vbKeyPageUp, vbKeyHome
1160        intRetVal = 0
1170        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1180      Case vbKeyPageDown, vbKeyEnd
1190        intRetVal = 0
1200        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1210      End Select
1220    End If

        ' ** Ctrl-Shift keys.
1230    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1240      Select Case intRetVal
          Case vbKeyF
1250        With Me
1260          intRetVal = 0
1270          DoCmd.SelectObject acForm, .Parent.Name, False
1280          .Parent.FocusHolder.SetFocus
1290        End With
1300      Case vbKeyTab
1310        With Me
1320          intRetVal = 0
1330          DoCmd.SelectObject acForm, .Parent.Name, False
1340          .Parent.cmdClose.SetFocus
1350        End With
1360      End Select
1370    End If

EXITP:
1380    KeyCode = intRetVal
1390    Exit Sub

ERRH:
1400    intRetVal = 0
1410    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
1420    Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Public Sub cmdSave_Click()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1510    DoCmd.RunCommand acCmdSaveRecord
1520    DoEvents

EXITP:
1530    Exit Sub

ERRH:
1540    Select Case ERR.Number
        Case Else
1550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1560    End Select
1570    Resume EXITP

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

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

1710    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub comment_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "comment_lbl_DblClick"

1810    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub Comment_AfterUpdate()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Comment_AfterUpdate"

1910    cmdSave_Click  ' ** Procedure: Above.
1920    Forms("frmAccountComments").chkUpdated = True
1930    Forms("frmAccountComments").cmdUpdate.Enabled = True

EXITP:
1940    Exit Sub

ERRH:
1950    Select Case ERR.Number
        Case Else
1960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1970    End Select
1980    Resume EXITP

End Sub

Private Sub Comment_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Comment_KeyDown"

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
2090          lngRecsCur = RecCnt  ' ** Function: Below.
2100          If .CurrentRecord < lngRecsCur Then
2110            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2120          Else
2130            DoCmd.SelectObject acForm, .Parent.Name, False
2140            If .Parent.chkIncludeArchive.Enabled = True Then
2150              .Parent.chkIncludeArchive.SetFocus
2160            ElseIf .Parent.chkIncludeCurrency.Enabled = True Then
2170              .Parent.chkIncludeCurrency.SetFocus
2180            ElseIf .Parent.cmdUpdate.Enabled = True Then
2190              .Parent.cmdUpdate.SetFocus
2200            Else
2210              .Parent.cmdClose.SetFocus
2220            End If
2230          End If
2240        End With
2250      End Select
2260    End If

        ' ** Shift keys.
2270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2280      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2290        With Me
2300          intRetVal = 0
2310          If .CurrentRecord > 1 Then
2320            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2330          Else
2340            DoCmd.SelectObject acForm, .Parent.Name, False
2350            .Parent.cmdClose.SetFocus
2360          End If
2370        End With
2380      End Select
2390    End If

EXITP:
2400    KeyCode = intRetVal
2410    Exit Sub

ERRH:
2420    intRetVal = 0
2430    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2440    Case Else
2450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2460    End Select
2470    Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "ICash_lbl_DblClick"

2510    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub icash_str_Enter()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_str_Enter"

2610    With Me
2620      If .curr_id = 150 Then
2630        .icash_str.BackColor = CLR_DISABLED_BG
2640      Else
2650        .icash_str.BackColor = CLR_LTGRN
2660      End If
2670    End With

EXITP:
2680    Exit Sub

ERRH:
2690    Select Case ERR.Number
        Case Else
2700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2710    End Select
2720    Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "PCash_lbl_DblClick"

2810    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub pcash_str_Enter()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_str_Enter"

2910    With Me
2920      If .curr_id = 150 Then
2930        .pcash_str.BackColor = CLR_DISABLED_BG
2940      Else
2950        .pcash_str.BackColor = CLR_LTGRN
2960      End If
2970    End With

EXITP:
2980    Exit Sub

ERRH:
2990    Select Case ERR.Number
        Case Else
3000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3010    End Select
3020    Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Cost_lbl_DblClick"

3110    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub cost_str_Enter()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_str_Enter"

3210    With Me
3220      If .curr_id = 150 Then
3230        .cost_str.BackColor = CLR_DISABLED_BG
3240      Else
3250        .cost_str.BackColor = CLR_LTGRN
3260      End If
3270    End With

EXITP:
3280    Exit Sub

ERRH:
3290    Select Case ERR.Number
        Case Else
3300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3310    End Select
3320    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

3410    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub curr_id_Enter()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_Enter"

3510    With Me
3520      If .curr_id = 150 Then
3530        .curr_id.BackColor = CLR_DISABLED_BG
3540      Else
3550        .curr_id.BackColor = CLR_LTGRN
3560      End If
3570    End With

EXITP:
3580    Exit Sub

ERRH:
3590    Select Case ERR.Number
        Case Else
3600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3610    End Select
3620    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnInclude As Boolean)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim blnSortHere As Boolean, blnResort As Boolean

3710    With Me
3720      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
3730        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
3740      End If
3750      blnSortHere = False: blnResort = False
3760      Select Case blnInclude
          Case True
3770        .Width = lngForm_Width
3780        .curr_id.Left = lngCurrID_Left
3790        .curr_id_lbl.Left = (.curr_id.Left - lngCurrIDLbl_Offset)
3800        .curr_id_lbl_line.Left = .curr_id.Left
3810        .curr_id_box.Left = .curr_id.Left
3820        .curr_id_forex_bg.Left = (.curr_id.Left - lngTpp)
3830        .curr_id.Visible = True
3840        .curr_id_lbl.Visible = True
3850        .curr_id_lbl_line.Visible = True
3860        .curr_id_box.Visible = True
3870        .curr_id_forex_bg.Visible = True
3880        .IsArch_bg1.Width = .Width
3890        .IsArch_bg2.Width = .Width
3900      Case False
3910        If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
3920        .curr_id.Visible = False
3930        .curr_id_lbl.Visible = False
3940        .curr_id_lbl_line.Visible = False
3950        .curr_id_box.Visible = False
3960        .curr_id_forex_bg.Visible = False
3970        .curr_id.Left = 0&
3980        .curr_id_lbl.Left = 0&
3990        .curr_id_lbl_line.Left = 0&
4000        .curr_id_box.Left = 0&
4010        .curr_id_forex_bg.Left = 0&
4020        If blnSortHere = True Then
4030          .Sort_lbl.Visible = False
4040          .Sort_line.Visible = False
4050          blnResort = True
4060        End If
4070        blnSortHere = False
4080        .IsArch_bg1.Width = ((lngForm_Width - lngCurrID_Width) + (1& * lngTpp))
4090        .IsArch_bg2.Width = ((lngForm_Width - lngCurrID_Width) + (1& * lngTpp))
4100        If blnResort = True Then
4110          SortNow "Form_Load"  ' ** Procedure: Below.
4120        End If
4130        DoEvents
4140        .Width = ((lngForm_Width - lngCurrID_Width) + (1& * lngTpp))
4150      End Select
4160      DoEvents
4170    End With

EXITP:
4180    Exit Sub

ERRH:
4190    Select Case ERR.Number
        Case Else
4200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4210    End Select
4220    Resume EXITP

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
4740            .FindFirst "[journalno] = " & CStr(varID)
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

Public Sub SortNow(strProc As String)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [journaltype], [journalno]"

5010    With Me
5020      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5030        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5040      End If
5050      .Controls(strSortLbl).Visible = False
5060      .Controls(strSortLine).Visible = False
5070      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push of the right side of the form.
5080      If strProc = "Form_Load" Then
5090        strCalled = "transdate"
5100        strSortNow = strSortOrig
5110        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5120        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5130        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5140        .Controls(strSortLbl).Left = lngSortLbl_Left
5150        .Controls(strSortLine).Left = lngSortLine_Left
5160        .Controls(strSortLine).Width = lngSortLine_Width
5170        .Controls(strSortLbl).Caption = strArwDn   ' ** Original is descending.
5180        .Controls(strSortLbl).ForeColor = CLR_DKRED
5190      Else
5200        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
5210        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5220        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5230        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5240        .Controls(strSortLbl).Left = lngSortLbl_Left
5250        .Controls(strSortLine).Left = lngSortLine_Left
5260        .Controls(strSortLine).Width = lngSortLine_Width
5270        If strCalled = "transdate" Then
5280          strSortAsc = "[transdate], [journaltype], [journalno]"  ' ** All are ascending.
5290          If strSortNow = strSortAsc Then
5300            strSortNow = strSortOrig   ' ** Original is descending: "[transdate] DESC, [journaltype] ASC, [journalno] DESC"
5310            .Controls(strSortLbl).Caption = strArwDn
5320            .Controls(strSortLbl).ForeColor = CLR_DKRED
5330          Else
5340            strSortNow = strSortAsc
5350            .Controls(strSortLbl).Caption = strArwUp
5360            .Controls(strSortLbl).ForeColor = CLR_DKBLU
5370          End If
5380        Else
5390          .Controls(strSortLbl).Caption = strArwUp
5400          .Controls(strSortLbl).ForeColor = CLR_DKBLU
5410          Select Case strCalled
              Case "journaltype"
5420            strSortAsc = "[journaltype], [transdate], [journalno]"
5430          Case "Comment"
5440            strSortAsc = "[Comment]" & strStdAsc
5450          Case "ICash"
5460            strSortAsc = "[ICash]" & strStdAsc
5470          Case "PCash"
5480            strSortAsc = "[PCash]" & strStdAsc
5490          Case "Cost"
5500            strSortAsc = "[Cost]" & strStdAsc
5510          Case "curr_id"
5520            strSortAsc = "[curr_code]" & strStdAsc
5530          End Select
5540          If strSortNow = strSortAsc Then
5550            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
5560            Select Case intCnt
                Case 1
5570              strTmp01 = strSortAsc & " DESC"
5580            Case 2
5590              intPos01 = InStr(strSortAsc, ",")
5600              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5610              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
5620            Case 3
5630              intPos01 = InStr(strSortAsc, ",")
5640              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5650              strTmp02 = Mid(strSortAsc, intPos01)
5660              intPos01 = InStr(2, strTmp02, ",")
5670              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5680              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
5690            Case 4
5700              intPos01 = InStr(strSortAsc, ",")
5710              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5720              strTmp02 = Mid(strSortAsc, intPos01)
5730              intPos01 = InStr(2, strTmp02, ",")
5740              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5750              strTmp02 = Mid(strTmp02, intPos01)
5760              intPos01 = InStr(2, strTmp02, ",")
5770              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5780              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
5790            Case 5
5800              intPos01 = InStr(strSortAsc, ",")
5810              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5820              strTmp02 = Mid(strSortAsc, intPos01)
5830              intPos01 = InStr(2, strTmp02, ",")
5840              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5850              strTmp02 = Mid(strTmp02, intPos01)
5860              intPos01 = InStr(2, strTmp02, ",")
5870              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5880              strTmp02 = Mid(strTmp02, intPos01)
5890              intPos01 = InStr(2, strTmp02, ",")
5900              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5910              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
5920            End Select
5930            strSortNow = strTmp01
5940            .Controls(strSortLbl).Caption = strArwDn
5950            .Controls(strSortLbl).ForeColor = CLR_DKRED
5960          Else
5970            strSortNow = strSortAsc
5980          End If
5990        End If
6000      End If
6010      .Controls(strSortLbl).Visible = True
6020      .Controls(strSortLine).Visible = True
6030      .OrderBy = strSortNow
6040      .OrderByOn = True
6050    End With

EXITP:
6060    Exit Sub

ERRH:
6070    Select Case ERR.Number
        Case Else
6080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6090    End Select
6100    Resume EXITP

End Sub
