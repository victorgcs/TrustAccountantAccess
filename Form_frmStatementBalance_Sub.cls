VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStatementBalance_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmStatementBalance_Sub"

'VGC 03/19/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Total Account:    A {AccountValue on frmStatementBalance}
' **   Close:            C {cmdClose on frmStatementBalance}
' **   Statement Date:   D {balance_date}
' **   Edit:             E {cmdXAdgfx_image1}
' **   Income Cash:      I {icash on frmStatementBalance}
' **   Last Stmt Date:   L {cmdLastStmtDate on frmStatementBalance}
' **   Total Market:     M {TotalMarketValue on frmStatementBalance}
' **   Principal Cash:   N {pcash on frmStatementBalance}
' **   Cost:             O {cost on frmStatementBalance}
' **   Print Report:     P {cmdPrintReport on frmStatementBalance}
' **   Update:           U {cmdUpdate on frmStatementBalance}
' **   Exit:             X {cmdClose on frmStatementBalance}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmStatementBalance}
' **   Next Record:      N {MoveRec on frmStatementBalance}
' **   Previous Record:  P {MoveRec on frmStatementBalance}

' ** NOTE: Lock image is changed via queries in ChangeDate() on frmStatementBalance.

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[accountno], [balance_date] DESC"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngRecsCur As Long, lngTpp As Long, lngBalID As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmStatementBalance" Then

160         lngBalID = 0&

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180         lngSortLbl_Width = .Sort_lbl.Width

190         CLR_DISABLED_FG = CLR_DKGRY
200         CLR_DISABLED_BG = CLR_LTTEAL

210       Else
220         Cancel = -1
230       End If
240     End With

250     If Cancel = -1 Then
260       Beep
270       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
280     End If

EXITP:
290     Exit Sub

ERRH:
300     Select Case ERR.Number
        Case Else
310       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
320     End Select
330     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       lngRecsCur = RecCnt  ' ** Function: Below
430       If lngRecsCur > 0& Then
440         .balance_date.ForeColor = CLR_DISABLED_FG
450         .balance_date.BackColor = CLR_DISABLED_BG
460         .ICash.ForeColor = CLR_DISABLED_FG
470         .ICash.BackColor = CLR_DISABLED_BG
480         .PCash.ForeColor = CLR_DISABLED_FG
490         .PCash.BackColor = CLR_DISABLED_BG
500         .Cost.ForeColor = CLR_DISABLED_FG
510         .Cost.BackColor = CLR_DISABLED_BG
520         .TotalMarketValue.ForeColor = CLR_DISABLED_FG
530         .TotalMarketValue.BackColor = CLR_DISABLED_BG
540         .AccountValue.ForeColor = CLR_DISABLED_FG
550         .AccountValue.BackColor = CLR_DISABLED_BG
560         SortNow THIS_PROC
570       End If
580     End With

EXITP:
590     Exit Sub

ERRH:
600     Select Case ERR.Number
        Case Else
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End Select
630     Resume EXITP

End Sub

Private Sub Form_Current()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

710     With Me
720       .ctlCurrentRecord = .SelTop
730     End With

EXITP:
740     Exit Sub

ERRH:
750     Select Case ERR.Number
        Case Else
760       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770     End Select
780     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyDown
870         intRetVal = 0
880         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
890       Case vbKeyUp
900         intRetVal = 0
910         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
920       Case vbKeyEscape
930         With Me
940           intRetVal = 0
950           .Parent.cmdClose_Click  ' ** Form Procedure: frmStatementBalance.
960         End With
970       End Select
980     End If

        ' ** Alt keys.
990     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1000      Select Case intRetVal
          Case vbKeyD
1010        With Me
1020          intRetVal = 0
1030          .balance_date.SetFocus
1040        End With
1050      Case vbKeyX
1060        With Me
1070          intRetVal = 0
1080          .Parent.cmdClose_Click  ' ** Form Procedure: frmStatementBalance.
1090        End With
1100      End Select
1110    End If

        ' ** Ctrl keys.
1120    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1130      Select Case intRetVal
          Case vbKeyN
1140        With Me
1150          intRetVal = 0
1160          .Parent.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmStatementBalance.
1170        End With
1180      Case vbKeyP
1190        With Me
1200          intRetVal = 0
1210          .Parent.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmStatementBalance.
1220        End With
1230      Case vbKeyS
1240        With Me
1250          intRetVal = 0
1260          .Parent.cmdSave_Click  ' ** Form Procedure: frmStatementBalance.
1270        End With
1280      Case vbKeyTab, vbKeyReturn
1290        With Me
1300          intRetVal = 0
1310          DoCmd.SelectObject acForm, .Parent.Name, False
1320          If .Parent.cmdPrintReport.Enabled = True Then
1330            .Parent.cmdPrintReport.SetFocus
1340          ElseIf .Parent.cmdUpdate.Enabled = True Then
1350            .Parent.cmdUpdate.SetFocus
1360          ElseIf .Parent.cmbAccounts.Enabled = True And .Parent.cmbAccounts.Visible = True Then
1370            .Parent.cmbAccounts.SetFocus
1380          Else
1390            .Parent.cmdClose.SetFocus
1400          End If
1410        End With
1420      Case vbKeyPageUp, vbKeyHome
1430        intRetVal = 0
1440        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1450      Case vbKeyPageDown, vbKeyEnd
1460        intRetVal = 0
1470        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1480      End Select
1490    End If

        ' ** Ctrl-Shift keys.
1500    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1510      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1520        With Me
1530          intRetVal = 0
1540          DoCmd.SelectObject acForm, .Parent.Name, False
1550          .Parent.AccountValue.SetFocus
1560        End With
1570      End Select
1580    End If

EXITP:
1590    KeyCode = intRetVal
1600    Exit Sub

ERRH:
1610    intRetVal = 0
1620    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
1630    Case Else
1640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1650    End Select
1660    Resume EXITP

End Sub

Private Sub Form_Click()
' ** GetLineNumber() in modMiscFuncs.
' ** ctlBack1.ControlSource: =IIf([SelTop]=[ctlCurrentLine],String(255,"Û"),Null)
' ** ctlBack2.ControlSource:

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Click"

1710    With Me
1720      .ctlCurrentRecord = .SelTop
1730    End With

EXITP:
1740    Exit Sub

ERRH:
1750    Select Case ERR.Number
        Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Private Sub cmdXAdgfx_image1_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdXAdgfx_image1_Click"

1810    With Me
1820      If .bal_lock = True Then
1830        lngBalID = .bal_id
1840        .Parent.ChangeDate .balance_date, lngBalID  ' ** Form Procedure: frmStatementBalance.
1850        DoEvents
1860        MoveRec 0, lngBalID  ' ** Procedure: Below.
1870      End If
1880    End With

EXITP:
1890    Exit Sub

ERRH:
1900    Select Case ERR.Number
        Case Else
1910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1920    End Select
1930    Resume EXITP

End Sub

Private Sub cmdXAdgfx_image1_DblClick(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdXAdgfx_image1_DblClick"

2010    cmdXAdgfx_image1_Click  ' ** Procedure: Above.

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub cmdXAdgfx_image1_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdXAdgfx_image1_KeyDown"

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
          Case vbKeyTab
2170        With Me
2180          intRetVal = 0
2190          .balance_date.SetFocus
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab
2250        With Me
2260          intRetVal = 0
2270          If .CurrentRecord > 1 Then
2280            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2290            .AccountValue.SetFocus
2300          Else
2310            DoCmd.SelectObject acForm, .Parent.Name, False
2320            .Parent.AccountValue.SetFocus
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
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2410    Case Else
2420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2430    End Select
2440    Resume EXITP

End Sub

Private Sub cmdXAdgfx_image1_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdXAdgfx_image1_GotFocus"

2510    With Me
2520      .ctlBack1.Visible = True
2530      .ctlBack2.Visible = True
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

Private Sub cmdXAdgfx_image1_LostFocus()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdXAdgfx_image1_LostFocus"

2610    With Me
2620      .ctlBack1.Visible = False
2630      .ctlBack2.Visible = False
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

Private Sub balance_date_lbl_DblClick(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "balance_date_lbl_DblClick"

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

Private Sub balance_date_lbl2_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "balance_date_lbl2_DblClick"

2810    balance_date_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

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

Private Sub pcash_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

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

Private Sub cost_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

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

Private Sub TotalMarketValue_lbl_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "TotalMarketValue_lbl_DblClick"

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

Private Sub TotalMarketValue_lbl2_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "TotalMarketValue_lbl2_DblClick"

3310    TotalMarketValue_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub AccountValue_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "AccountValue_lbl_DblClick"

3410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub AccountValue_lbl2_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "AccountValue_lbl2_DblClick"

3510    AccountValue_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub AccountValue_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "AccountValue_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3620    intShiftDown = (Shift And acShiftMask) > 0
3630    intAltDown = (Shift And acAltMask) > 0
3640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3670        With Me
3680          intRetVal = 0
3690          lngRecsCur = RecCnt  ' ** Function: Below.
3700          If .CurrentRecord < lngRecsCur Then
3710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3720            .cmdXAdgfx_image1.SetFocus
3730          Else
3740            DoCmd.SelectObject acForm, .Parent.Name, False
3750            If .Parent.cmdPrintReport.Enabled = True Then
3760              .Parent.cmdPrintReport.SetFocus
3770            ElseIf .Parent.cmdUpdate.Enabled = True Then
3780              .Parent.cmdUpdate.SetFocus
3790            ElseIf .Parent.cmbAccounts.Enabled = True And .Parent.cmbAccounts.Visible = True Then
3800              .Parent.cmbAccounts.SetFocus
3810            Else
3820              .Parent.cmdClose.SetFocus
3830            End If
3840          End If
3850        End With
3860      End Select
3870    End If

        ' ** Shift keys.
3880    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3890      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3900        With Me
3910          intRetVal = 0
3920          .TotalMarketValue.SetFocus
3930        End With
3940      End Select
3950    End If

EXITP:
3960    KeyCode = intRetVal
3970    Exit Sub

ERRH:
3980    intRetVal = 0
3990    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
4000    Case Else
4010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4020    End Select
4030    Resume EXITP

End Sub

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
4540            .FindFirst "[bal_id] = " & CStr(varID)
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
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
4660    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4670    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4680    Case Else
4690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4700    End Select
4710    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [accountno], [balance_date] DESC"

4810    With Me
4820      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
4830        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4840      End If
4850      .Controls(strSortLbl).Visible = False
4860      .Controls(strSortLine).Visible = False
4870      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
4880      If strProc = "Form_Load" Then
4890        strCalled = "balance_date"
4900        strSortNow = strSortOrig
4910        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4920        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (6& * lngTpp))
4930        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
4940        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
4950        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
4960        .Controls(strSortLbl).Top = lngSortLbl_Top
4970        .Controls(strSortLbl).Left = lngSortLbl_Left
4980        .Controls(strSortLine).Top = lngSortLine_Top
4990        .Controls(strSortLine).Left = lngSortLine_Left
5000        .Controls(strSortLine).Width = lngSortLine_Width
5010        .Controls(strSortLbl).Caption = strArwDn  ' ** Balance_Date Descending.
5020        .Controls(strSortLbl).ForeColor = CLR_DKRED
5030      Else
5040        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
5050        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5060        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
5070        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
5080        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
5090        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
5100        .Controls(strSortLbl).Top = lngSortLbl_Top
5110        .Controls(strSortLbl).Left = lngSortLbl_Left
5120        .Controls(strSortLine).Top = lngSortLine_Top
5130        .Controls(strSortLine).Left = lngSortLine_Left
5140        .Controls(strSortLine).Width = lngSortLine_Width
5150        If strCalled = "balance_date" Then
5160          strSortAsc = "[accountno] DESC, [balance_date]"  ' ** accountno isn't visible, and all are the same.
5170          If strSortNow = strSortOrig Then
5180            strSortNow = strSortAsc
5190            lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
5200            .Controls(strSortLbl).Left = lngSortLbl_Left
5210            .Controls(strSortLbl).Caption = strArwUp
5220            .Controls(strSortLbl).ForeColor = CLR_DKBLU
5230          Else
5240            strSortNow = strSortOrig
5250            .Controls(strSortLbl).Caption = strArwDn
5260            .Controls(strSortLbl).ForeColor = CLR_DKRED
5270          End If
5280        Else
5290          .Controls(strSortLbl).Caption = strArwUp
5300          .Controls(strSortLbl).ForeColor = CLR_DKBLU
5310          Select Case strCalled
              Case "icash"
5320            strSortAsc = "[icash]" & strStdAsc
5330          Case "pcash"
5340            strSortAsc = "[pcash]" & strStdAsc
5350          Case "cost"
5360            strSortAsc = "[cost]" & strStdAsc
5370          Case "TotalMarketValue"
5380            strSortAsc = "[TotalMarketValue]" & strStdAsc
5390          Case "AccountValue"
5400            strSortAsc = "[AccountValue]" & strStdAsc
5410          End Select
5420          If strSortNow = strSortAsc Then
5430            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
5440            Select Case intCnt
                Case 1
5450              strTmp01 = strSortAsc & " DESC"
5460            Case 2
5470              intPos01 = InStr(strSortAsc, ",")
5480              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5490              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
5500            Case 3
5510              intPos01 = InStr(strSortAsc, ",")
5520              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5530              strTmp02 = Mid(strSortAsc, intPos01)
5540              intPos01 = InStr(2, strTmp02, ",")
5550              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5560              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
5570            Case 4
5580              intPos01 = InStr(strSortAsc, ",")
5590              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
5600              strTmp02 = Mid(strSortAsc, intPos01)
5610              intPos01 = InStr(2, strTmp02, ",")
5620              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5630              strTmp02 = Mid(strTmp02, intPos01)
5640              intPos01 = InStr(2, strTmp02, ",")
5650              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
5660              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
5670            End Select
5680            strSortNow = strTmp01
5690            .Controls(strSortLbl).Caption = strArwDn
5700            .Controls(strSortLbl).ForeColor = CLR_DKRED
5710          Else
5720            strSortNow = strSortAsc
5730          End If
5740        End If
5750      End If
5760      .Controls(strSortLbl).Visible = True
5770      .Controls(strSortLine).Visible = True
5780      .OrderBy = strSortNow
5790      .OrderByOn = True
5800    End With

EXITP:
5810    Exit Sub

ERRH:
5820    Select Case ERR.Number
        Case Else
5830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5840    End Select
5850    Resume EXITP

End Sub
