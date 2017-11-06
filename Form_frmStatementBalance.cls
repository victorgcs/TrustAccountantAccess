VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStatementBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmStatementBalance"

'VGC 03/19/2017: CHANGES!

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Total Account:    A {AccountValue}
' **   Close:            C {cmdClose}
' **   Statement Date:   D {balance_date on frmStatementBalance_Sub}
' **   Edit:             E {cmdXAdgfx_image1 on frmStatementBalance_Sub}
' **   Income Cash:      I {icash}
' **   Last Stmt Date:   L {cmdLastStmtDate}
' **   Total Market:     M {TotalMarketValue}
' **   Principal Cash:   N {pcash}
' **   Cost:             O {cost}
' **   Print Report:     P {cmdPrintReport}
' **   Update:           U {cmdUpdate}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}
' **   Next Record:      N {MoveRec}
' **   Previous Record:  P {MoveRec}

' ** NOTE: Lock image is changed via queries in ChangeDate(), below.

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Array: arr_varAcct().
Private lngAccts As Long, arr_varAcct As Variant
Private Const A_NUM As Integer = 0
Private Const A_DAT As Integer = 1

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean
Private blnRecAdded As Boolean, blnMoveSomewhere As Boolean
Private strAccountNo As String, strLastDate As String, blnOpenFocus As Boolean
Private lngGfxID_Open As Long, lngGfxID_Closed As Long, lngGfxID_OpenFocus As Long, lngGfxID_ClosedFocus As Long
Private lngRecsCur As Long
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strFilter As String, strOrderBy As String
        Dim lngRecs As Long
        Dim blnMsg As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long

110     blnMsg = False

120     With Me

130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then
150           strCallingForm = Trim(.OpenArgs)
160           intPos01 = InStr(strCallingForm, "~")
170           If intPos01 > 0 Then

180             CLR_DISABLED_FG = CLR_DKGRY
190             CLR_DISABLED_BG = CLR_LTTEAL

200             strAccountNo = Mid(strCallingForm, (intPos01 + 1))
210             strCallingForm = Left(strCallingForm, (intPos01 - 1))
220             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
230             blnRecAdded = False: blnMoveSomewhere = False

240             If IsLoaded("frmStatementParameters", acForm) = True Then  ' ** Module Function: modFileUtilities.
250               .chkLegalName = Forms("frmStatementParameters").chkLegalName
260             Else
270               .chkLegalName = False
280             End If

290             .cmdLastStmtDate.Visible = True
300             .cmdLastStmtDate.Enabled = True
310             .cmdAdd.Visible = True
320             .cmdAdd.Enabled = True

330           Else
340             Cancel = -1
350           End If
360         Else
370           Cancel = -1
380         End If
390       Else
400         Cancel = -1
410       End If

420       If Cancel = 0 Then
430         Set dbs = CurrentDb
440         With dbs
              ' ** Balance table, grouped by accountno, with Max(balance_date).
450           Set qdf = .QueryDefs("qryStatementBalance_01")
460           Set rst = qdf.OpenRecordset
470           With rst
480             If .BOF = True And .EOF = True Then
                  ' ** Shouldn't happen.
490               Cancel = -1
500               Beep
510               MsgBox "There are no past balances.", vbCritical + vbOKOnly, "Balance Table Is Empty"
520               blnMsg = True
530             Else
540               .MoveLast
550               lngAccts = .RecordCount
560               .MoveFirst
570               arr_varAcct = .GetRows(lngAccts)
                  ' *************************************************
                  ' ** Array: arr_varHide()
                  ' **
                  ' **   Field  Element  Name            Constant
                  ' **   =====  =======  ==============  ==========
                  ' **     1       0     accountno       A_NUM
                  ' **     2       1     balance_date    A_DAT
                  ' **
                  ' *************************************************
580             End If
590             .Close
600           End With
610           .Close
620         End With
630       End If  ' ** Cancel.

640       If Cancel = 0 Then

650         SetNav_Access2007 Me  ' ** Module Procedure: modXAccess_07_10_Funcs.
            'SystemColor_Move Me  ' ** Module Procedure: modSystemColorFuncs.

660         strFilter = vbNullString: strLastDate = vbNullString
670         strTmp01 = vbNullString: strTmp02 = vbNullString
680         If strAccountNo = "ALL" Then
690           .cmbAccounts.Visible = True
700           .accountno.Visible = False
710           .cmbShortName.Visible = True
720           .shortname.Visible = False
730           lngRecs = .cmbAccounts.ListCount
740           For lngX = 0& To (lngRecs - 1&)
750             If .cmbAccounts.Column(2, lngX) = "X" Then
760               strTmp01 = .cmbAccounts.Column(0, lngX)
770               strTmp02 = .cmbAccounts.Column(1, lngX)
780               Exit For
790             End If
800           Next
810           .cmbAccounts = strTmp01
820           .cmbShortName = strTmp02
830         Else
840           .accountno.Visible = True
850           .cmbAccounts.Visible = False
860           .shortname.Visible = True
870           .cmbShortName.Visible = False
880           strTmp01 = strAccountNo
890         End If

900         For lngX = 0& To (lngAccts - 1&)
910           If arr_varAcct(A_NUM, lngX) = strTmp01 Then
920             strLastDate = Format(arr_varAcct(A_DAT, lngX), "mm/dd/yyyy")
930             Exit For
940           End If
950         Next

960         If strLastDate <> vbNullString Then
970           .LastDate = CDate(strLastDate)
980           strFilter = "[accountno] = '" & strTmp01 & "' And [balance_date] = #" & strLastDate & "#"
990         ElseIf strLastDate = vbNullString And strAccountNo = "ALL" Then
              ' ** Allow them to choose another.
1000          MsgBox "No previous balance exists for this account.", vbInformation + vbOKOnly, "No Balance Found"
1010          .balance_date.Enabled = False
1020          .TotalMarketValue.Enabled = False
1030          .AccountValue.Enabled = False
1040          .ICash.Enabled = False
1050          .PCash.Enabled = False
1060          .Cost.Enabled = False
1070          .frmStatementBalance_Sub.Enabled = False
1080        Else
1090          Cancel = -1
1100          DoCmd.Hourglass False
1110          Beep
1120          MsgBox "No previous balance exists for this account.", vbInformation + vbOKOnly, "No Balance Found"
1130          blnMsg = True
1140        End If

1150        If Cancel = 0 Then

1160          lngGfxID_Open = DLookup("[xadgfx_id]", "tblXAdmin_Graphics", "[xadgfx_name] = 'LockOpenDes'")
1170          lngGfxID_Closed = DLookup("[xadgfx_id]", "tblXAdmin_Graphics", "[xadgfx_name] = 'LockClosedDes'")
1180          lngGfxID_OpenFocus = DLookup("[xadgfx_id]", "tblXAdmin_Graphics", "[xadgfx_name] = 'LockOpenDesFoc'")
1190          lngGfxID_ClosedFocus = DLookup("[xadgfx_id]", "tblXAdmin_Graphics", "[xadgfx_name] = 'LockClosedDesFoc'")

1200          If strFilter <> vbNullString Then
1210            .Filter = strFilter
1220            .FilterOn = True
1230          End If

1240          strOrderBy = "[balance_date] DESC"
1250          .OrderBy = strOrderBy
1260          .OrderByOn = True

1270          .cmdUpdate.Enabled = False

1280        End If
1290      End If  ' ** Cancel.

1300    End With

1310    If Cancel = -1 And blnMsg = False Then
1320      Beep
1330      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1340    End If

EXITP:
1350    Set rst = Nothing
1360    Set qdf = Nothing
1370    Set dbs = Nothing
1380    Exit Sub

ERRH:
1390    Cancel = -1
1400    Select Case ERR.Number
        Case Else
1410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1420    End Select
1430    Resume EXITP

End Sub

Private Sub Form_Load()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1510    With Me
1520      blnOpenFocus = True
1530      .accountno.ForeColor = CLR_DISABLED_FG
1540      .accountno.BackColor = CLR_DISABLED_BG
1550      .shortname.ForeColor = CLR_DISABLED_FG
1560      .shortname.BackColor = CLR_DISABLED_BG
1570      .last_balance_date.ForeColor = CLR_DISABLED_FG
1580      .last_balance_date.BackColor = CLR_DISABLED_BG
1590      .balance_date.ForeColor = CLR_DISABLED_FG
1600      .balance_date.BackColor = CLR_DISABLED_BG
1610      DoCmd.Hourglass False
1620      .ICash.SetFocus
1630    End With

EXITP:
1640    Exit Sub

ERRH:
1650    Select Case ERR.Number
        Case Else
1660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1670    End Select
1680    Resume EXITP

End Sub

Private Sub Form_Current()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

1710    With Me
1720      If .balance_date = .last_balance_date Then
1730        .last_balance_date_lbl3.Visible = True
1740        .last_balance_date_lbl4.Visible = True
1750        .balance_date_lbl3.Visible = True
1760        .balance_date_lbl4.Visible = True
1770      Else
1780        .last_balance_date_lbl3.Visible = False
1790        .last_balance_date_lbl4.Visible = False
1800        .balance_date_lbl3.Visible = False
1810        .balance_date_lbl4.Visible = False
1820      End If
1830    End With

EXITP:
1840    Exit Sub

ERRH:
1850    Select Case ERR.Number
        Case Else
1860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1870    End Select
1880    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

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
          Case vbKeyEscape
1970        intRetVal = 0
1980        cmdClose_Click  ' ** Procedure: Below
1990      End Select
2000    End If

        ' ** Alt keys.
2010    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2020      Select Case intRetVal
          Case vbKeyD
2030        With Me
2040          intRetVal = 0
2050          .frmStatementBalance_Sub.SetFocus
2060          .frmStatementBalance_Sub.Form.balance_date.SetFocus
2070        End With
2080      Case vbKeyX
2090        intRetVal = 0
2100        cmdClose_Click  ' ** Procedure: Below.
2110      End Select
2120    End If

        ' ** Ctrl keys.
2130    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2140      Select Case intRetVal
          Case vbKeyN
2150        intRetVal = 0
2160        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2170      Case vbKeyP
2180        intRetVal = 0
2190        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2200      Case vbKeyS
2210        intRetVal = 0
2220        cmdSave_Click  ' ** Procedure: Below.
2230      End Select
2240    End If

        ' ** Ctrl-Shift keys.
2250    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2260      Select Case intRetVal
          Case vbKeyF
2270        With Me
2280          intRetVal = 0
2290          .FocusHolder.SetFocus
2300        End With
2310      End Select
2320    End If

EXITP:
2330    KeyCode = intRetVal
2340    Exit Sub

ERRH:
2350    intRetVal = 0
2360    Select Case ERR.Number
        Case Else
2370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2380    End Select
2390    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

2410    With Me
2420  On Error Resume Next
2430      .bal_username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2440      .bal_datemodified = Now()
2450  On Error GoTo ERRH
2460    End With

EXITP:
2470    Exit Sub

ERRH:
2480    Select Case ERR.Number
        Case 2115  ' ** The macro or function set to the BeforeUpdate or ValidationRule property for this
          ' ** Not sure why!  ' ** field is preventing Microsoft Access from saving the data in the field.
2490    Case Else
2500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2510    End Select
2520    Resume EXITP

End Sub

Private Sub Form_Timer()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2610    With Me
2620      .TimerInterval = 0&
2630      If gblnSetFocus = True Then
2640        gblnSetFocus = False
2650        strFormName = vbNullString
2660  On Error Resume Next
2670        Do Until strFormName = THIS_NAME
2680          DoCmd.SelectObject acForm, THIS_NAME, False
2690          strFormName = Screen.ActiveForm.Name
2700        Loop
2710  On Error GoTo ERRH
2720      ElseIf blnMoveSomewhere = True Then
2730        blnMoveSomewhere = False
2740        .ICash.SetFocus
2750      End If
2760      If gblnGoToReport = True Then
2770        DoCmd.Hourglass True  ' ** Make sure it's still running.
2780        DoEvents
2790        Select Case blnGoingToReport
            Case True
2800          gblnGoToReport = False
2810          blnGoingToReport = False
2820          .GoToReport_arw_pr_img.Visible = False
2830          DoCmd.Hourglass False
2840          DoEvents
2850        Case False
2860          blnGoingToReport = True
2870          .GoToReport_arw_pr_img.Visible = True
2880          Beep
2890          DoEvents
2900          .TimerInterval = (GTR_WAIT * 2&)
2910        End Select
2920      End If
2930    End With

EXITP:
2940    Exit Sub

ERRH:
2950    Select Case ERR.Number
        Case Else
2960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970    End Select
2980    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControl As String
        Dim dblDataErr As Double, strErrDesc As String

3010    With Me
3020      dblDataErr = DataErr
3030      strErrDesc = ERR.description
3040  On Error Resume Next
3050      strControl = .ActiveControl.Name
3060  On Error GoTo ERRH
3070      Select Case dblDataErr
          Case 2113, 2279
            ' ** 2113: The value you entered isn't valid for this field.
            ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3080        If strControl = "balance_date" Or strControl = "balance_date_entry" Then
3090          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, _
                ("Invalid Date" & Space(50) & "01")
3100          Response = acDataErrContinue
3110        Else
3120          Response = acDataErrDisplay
3130        End If
3140      Case 3314  ' ** The field '|' cannot contain a Null value because the Required property for this field is set to True.
3150        If strControl = "balance_date" Or strControl = "balance_date_entry" Then
3160          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, _
                ("Entry Required" & Space(50) & "02")
3170          Response = acDataErrContinue
3180        Else
3190          Response = acDataErrDisplay
3200        End If
3210      Case Else
3220        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3230        zErrorHandler THIS_NAME, THAT_PROC, dblDataErr, That_Erl, strErrDesc  ' ** Module Function: modErrorHandler.
3240      End Select
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

Private Sub Form_Unload(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3410    If Reports.Count > 0 Then
3420      Do While Reports.Count > 0
3430        DoCmd.Close acReport, Reports(0).Name
3440      Loop
3450    End If

3460    Select Case strCallingForm
        Case "frmStatementParameters"
3470      DoCmd.SelectObject acForm, strCallingForm, False
3480    End Select

EXITP:
3490    Exit Sub

ERRH:
3500    Select Case ERR.Number
        Case Else
3510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3520    End Select
3530    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3610    With Me
3620      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3630        Select Case blnPrintReport_Focus
            Case True
3640          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3650          .cmdPrintReport_raised_img.Visible = False
3660        Case False
3670          .cmdPrintReport_raised_img.Visible = True
3680          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3690        End Select
3700        .cmdPrintReport_raised_focus_img.Visible = False
3710        .cmdPrintReport_raised_focus_dots_img.Visible = False
3720        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3730        .cmdPrintReport_raised_img_dis.Visible = False
3740      End If
3750    End With

EXITP:
3760    Exit Sub

ERRH:
3770    Select Case ERR.Number
        Case Else
3780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3790    End Select
3800    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Public Sub cmdClose_Click()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4010    With Me
4020      Select Case .chkUpdated
          Case True
4030        If MsgBox("You have made changes, but have not confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") = vbYes Then
4040          DoCmd.Close acForm, THIS_NAME
4050        End If
4060      Case False
4070        DoCmd.Close acForm, THIS_NAME
4080      End Select
4090    End With

EXITP:
4100    Exit Sub

ERRH:
4110    Select Case ERR.Number
        Case Else
4120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4130    End Select
4140    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4220    intShiftDown = (Shift And acShiftMask) > 0
4230    intAltDown = (Shift And acAltMask) > 0
4240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4260      Select Case intRetVal
          Case vbKeyTab
4270        With Me
4280          intRetVal = 0
4290          If .balance_date.Visible = True And .balance_date.Enabled = True Then
4300            .balance_date.SetFocus
4310          ElseIf .balance_date_entry.Visible = True And .balance_date_entry.Enabled = True Then
4320            .balance_date_entry.SetFocus
4330          End If
4340        End With
4350      End Select
4360    End If

        ' ** Shift keys.
4370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4380      Select Case intRetVal
          Case vbKeyTab
4390        With Me
4400          intRetVal = 0
4410          If .cmbShortName.Enabled = True And .cmbShortName.Visible = True Then
4420            .cmbShortName.SetFocus
4430          ElseIf .cmdUpdate.Enabled = True Then
4440            .cmdUpdate.SetFocus
4450          ElseIf .cmdPrintReport.Enabled = True Then
4460            .cmdPrintReport.SetFocus
4470          Else
4480            lngRecsCur = .frmStatementBalance_Sub.Form.RecCnt  ' ** Form Function: frmStatementBalance_Sub.
4490            If lngRecsCur > 0& Then
4500              .frmStatementBalance_Sub.SetFocus
4510              .frmStatementBalance_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmStatementBalance_Sub.
4520              .frmStatementBalance_Sub.Form.AccountValue.SetFocus
4530            Else
4540              .AccountValue.SetFocus
4550            End If
4560          End If
4570        End With
4580      End Select
4590    End If

EXITP:
4600    KeyCode = intRetVal
4610    Exit Sub

ERRH:
4620    intRetVal = 0
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

4710    With Me
4720      blnPrintReport_Focus = True
4730      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
4740      .cmdPrintReport_raised_img.Visible = False
4750      .cmdPrintReport_raised_focus_img.Visible = False
4760      .cmdPrintReport_raised_focus_dots_img.Visible = False
4770      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4780      .cmdPrintReport_raised_img_dis.Visible = False
4790    End With

EXITP:
4800    Exit Sub

ERRH:
4810    Select Case ERR.Number
        Case Else
4820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4830    End Select
4840    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

4910    With Me
4920      blnPrintReport_MouseDown = True
4930      .cmdPrintReport_sunken_focus_dots_img.Visible = True
4940      .cmdPrintReport_raised_img.Visible = False
4950      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4960      .cmdPrintReport_raised_focus_img.Visible = False
4970      .cmdPrintReport_raised_focus_dots_img.Visible = False
4980      .cmdPrintReport_raised_img_dis.Visible = False
4990    End With

EXITP:
5000    Exit Sub

ERRH:
5010    Select Case ERR.Number
        Case Else
5020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5030    End Select
5040    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

5110    With Me
5120      strDocName = "rptStatement_Balance"
5130      If gblnDev_Debug Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
5140        DoCmd.OpenReport strDocName, acViewPreview, "[accountno] = '" & .accountno & "'"
5150        DoCmd.Maximize
5160        DoCmd.RunCommand acCmdFitToWindow
5170        .Visible = False
5180      Else
            '##GTR_Ref: rptStatement_Balance
5190        DoCmd.OpenReport strDocName, acViewNormal, "[accountno] = '" & .accountno & "'"
5200      End If
5210    End With

EXITP:
5220    Exit Sub

ERRH:
5230    Me.Visible = True
5240    DoCmd.Hourglass False
5250    Select Case ERR.Number
        Case Else
5260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5270    End Select
5280    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5320    intShiftDown = (Shift And acShiftMask) > 0
5330    intAltDown = (Shift And acAltMask) > 0
5340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab
5370        With Me
5380          intRetVal = 0
5390          If .cmdUpdate.Enabled = True Then
5400            .cmdUpdate.SetFocus
5410          ElseIf .cmbAccounts.Enabled = True And .cmbAccounts.Visible = True Then
5420            .cmbAccounts.SetFocus
5430          Else
5440            .cmdClose.SetFocus
5450          End If
5460        End With
5470      End Select
5480    End If

        ' ** Shift keys.
5490    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5500      Select Case intRetVal
          Case vbKeyTab
5510        With Me
5520          intRetVal = 0
5530          lngRecsCur = .frmStatementBalance_Sub.Form.RecCnt  ' ** Form Function: frmStatementBalance_Sub.
5540          If lngRecsCur > 0& Then
5550            .frmStatementBalance_Sub.SetFocus
5560            .frmStatementBalance_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmStatementBalance_Sub.
5570            .frmStatementBalance_Sub.Form.AccountValue.SetFocus
5580          Else
5590            .AccountValue.SetFocus
5600          End If
5610        End With
5620      End Select
5630    End If

EXITP:
5640    KeyCode = intRetVal
5650    Exit Sub

ERRH:
5660    intRetVal = 0
5670    Select Case ERR.Number
        Case Else
5680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5690    End Select
5700    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

5810    With Me
5820      If blnPrintReport_MouseDown = False Then
5830        Select Case blnPrintReport_Focus
            Case True
5840          .cmdPrintReport_raised_focus_dots_img.Visible = True
5850          .cmdPrintReport_raised_focus_img.Visible = False
5860        Case False
5870          .cmdPrintReport_raised_focus_img.Visible = True
5880          .cmdPrintReport_raised_focus_dots_img.Visible = False
5890        End Select
5900        .cmdPrintReport_raised_img.Visible = False
5910        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
5930        .cmdPrintReport_raised_img_dis.Visible = False
5940      End If
5950    End With

EXITP:
5960    Exit Sub

ERRH:
5970    Select Case ERR.Number
        Case Else
5980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5990    End Select
6000    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

6110    With Me
6120      .cmdPrintReport_raised_focus_dots_img.Visible = True
6130      .cmdPrintReport_raised_img.Visible = False
6140      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6150      .cmdPrintReport_raised_focus_img.Visible = False
6160      .cmdPrintReport_sunken_focus_dots_img.Visible = False
6170      .cmdPrintReport_raised_img_dis.Visible = False
6180      blnPrintReport_MouseDown = False
6190    End With

EXITP:
6200    Exit Sub

ERRH:
6210    Select Case ERR.Number
        Case Else
6220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6230    End Select
6240    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

6310    With Me
6320      .cmdPrintReport_raised_img.Visible = True
6330      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6340      .cmdPrintReport_raised_focus_img.Visible = False
6350      .cmdPrintReport_raised_focus_dots_img.Visible = False
6360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
6370      .cmdPrintReport_raised_img_dis.Visible = False
6380      blnPrintReport_Focus = False
6390    End With

EXITP:
6400    Exit Sub

ERRH:
6410    Select Case ERR.Number
        Case Else
6420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6430    End Select
6440    Resume EXITP

End Sub

Private Sub cmdAdd_Click()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim datNewDate As Date
        Dim lngMaxBalID_Then As Long, lngMaxBalID_Now As Long

6510    With Me

6520      strAccountNo = .accountno
6530      datNewDate = Date
6540      lngMaxBalID_Then = DMax("[bal_id]", "tmpBalance")

6550      Set dbs = CurrentDb
6560      With dbs

            ' ** Update tmpBalance, for bal_lock = True.
6570        Set qdf = .QueryDefs("qryStatementBalance_15")
6580        qdf.Execute

            ' ** Append qryStatementBalance_14e (Union of qryStatementBalance_14b (qryStatementBalance_14a
            ' ** (Account, linked to qryStatementBalance_01 (Balance table, grouped by accountno, with
            ' ** Max(balance_date)), with tblXAdmin_Graphics, for non-latest balance_date, lock closed, desert,
            ' ** by specified [actno]; Cartesian), with tblXAdmin_Graphics, lock closed focus, desert; Cartesian),
            ' ** qryStatementBalance_14d (qryStatementBalance_14c (Account, linked to qryStatementBalance_01
            ' ** (Balance table, grouped by accountno, with Max(balance_date)), with tblXAdmin_Graphics, for non-latest
            ' ** balance_date, lock open, desert, by specified [actno]; Cartesian), with tblXAdmin_Graphics, lock open
            ' ** focus, desert; Cartesian)), with qryStatementBalance_06b (tblForm_Graphics, just frmStatementBalance,
            ' ** for cmdPrintReport), to tmpBalance, by specified [actno]; Cartesian.
6590        Set qdf = .QueryDefs("qryStatementBalance_16")
6600        With qdf.Parameters
6610          ![actno] = strAccountNo
6620        End With
6630        qdf.Execute

6640        .Close
6650      End With  ' ** dbs.
6660      DoEvents

6670      lngMaxBalID_Now = DMax("[bal_id]", "tmpBalance")

6680      .Requery
6690      DoEvents
6700      .frmStatementBalance_Sub.Form.Requery
6710      DoEvents

6720      ChangeDate datNewDate, lngMaxBalID_Now  ' ** Procedure: Below.

6730      .FocusHolder.SetFocus
6740      .balance_date_entry.Enabled = True
6750      .balance_date_entry.Locked = False
6760      .balance_date_entry.Visible = True
6770      .balance_date.Visible = False
6780      .balance_date_entry.SetFocus
6790      DoEvents
6800      .balance_date_entry.SelLength = 0

6810      blnRecAdded = True
6820      .chkUpdated = True
6830      .cmdUpdate.Enabled = True

6840    End With  ' ** Me.

EXITP:
6850    Set qdf = Nothing
6860    Set dbs = Nothing
6870    Exit Sub

ERRH:
6880    Select Case ERR.Number
        Case Else
6890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6900    End Select
6910    Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

7010    DoCmd.Hourglass True

7020    cmdSave_Click  ' ** Procedure: Below.
7030    DoEvents

7040    Set dbs = CurrentDb
7050    With dbs

          ' ** Update Balance table from tmpBalance.
7060      Set qdf = .QueryDefs("qryStatementBalance_10")
7070      qdf.Execute dbFailOnError

7080      If blnRecAdded = True Then
            ' ** Append tmpBalance, not in Balance, to Balance.
7090        Set qdf = .QueryDefs("qryStatementBalance_17")
7100        qdf.Execute dbFailOnError
7110        blnRecAdded = False
7120      End If

7130      .Close
7140    End With

7150    With Me
7160      .chkUpdated = False
7170      .cmdClose.SetFocus
7180      DoEvents
7190      .cmdUpdate.Enabled = False
7200      If .balance_date_entry.Visible = True Then
7210        .balance_date.Visible = True
7220        .balance_date_entry.Visible = False
7230      End If
7240      DoCmd.Hourglass False
7250      MsgBox "Updates completed.", vbInformation + vbOKOnly, ("Update Successful" & Space(40))
7260    End With

EXITP:
7270    Set qdf = Nothing
7280    Set dbs = Nothing
7290    Exit Sub

ERRH:
7300    DoCmd.Hourglass False
7310    Select Case ERR.Number
        Case Else
7320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7330    End Select
7340    Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7420    intShiftDown = (Shift And acShiftMask) > 0
7430    intAltDown = (Shift And acAltMask) > 0
7440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7460      Select Case intRetVal
          Case vbKeyTab
7470        With Me
7480          intRetVal = 0
7490          If .cmbAccounts.Enabled = True And .cmbAccounts.Visible = True Then
7500            .cmbAccounts.SetFocus
7510          Else
7520            .cmdClose.SetFocus
7530          End If
7540        End With
7550      End Select
7560    End If

        ' ** Shift keys.
7570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7580      Select Case intRetVal
          Case vbKeyTab
7590        With Me
7600          intRetVal = 0
7610          If .cmdPrintReport.Enabled = True Then
7620            .cmdPrintReport.SetFocus
7630          Else
7640            lngRecsCur = .frmStatementBalance_Sub.Form.RecCnt  ' ** Form Function: frmStatementBalance_Sub.
7650            If lngRecsCur > 0& Then
7660              .frmStatementBalance_Sub.SetFocus
7670              .frmStatementBalance_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmStatementBalance_Sub.
7680              .frmStatementBalance_Sub.Form.AccountValue.SetFocus
7690            Else
7700              .AccountValue.SetFocus
7710            End If
7720          End If
7730        End With
7740      End Select
7750    End If

EXITP:
7760    KeyCode = intRetVal
7770    Exit Sub

ERRH:
7780    intRetVal = 0
7790    Select Case ERR.Number
        Case Else
7800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7810    End Select
7820    Resume EXITP

End Sub

Private Sub cmdLastStmtDate_Click()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLastStmtDate_Click"

        Dim strDocName As String

7910    strDocName = "frmStatementBalance_Date"
7920    DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME

EXITP:
7930    Exit Sub

ERRH:
7940    Select Case ERR.Number
        Case Else
7950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7960    End Select
7970    Resume EXITP

End Sub

Public Sub cmdSave_Click()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

8010    DoCmd.RunCommand acCmdSelectRecord
8020    DoCmd.RunCommand acCmdSaveRecord

EXITP:
8030    Exit Sub

ERRH:
8040    THAT_PROC = THIS_PROC
8050    That_Erl = Erl
8060    Select Case ERR.Number
        Case Else
8070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8080    End Select
8090    Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strFilter As String
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim lngX As Long

8110    With Me
8120      If IsNull(.cmbAccounts) = False Then

8130        If .Dirty = True Then
8140          cmdSave_Click  ' ** Procedure: Above.
8150        End If

            ' ** If the current accountno is not on the latest
            ' ** balance_date, reset it before moving on.
8160        If .balance_date <> .last_balance_date Then
8170          Set dbs = CurrentDb
8180          With dbs
                ' ** Update tmpBalance, by specified [actno], [gfxid], [datbal], [ballok].
8190            Set qdf = .QueryDefs("qryStatementBalance_09")
8200            With qdf.Parameters
8210              ![actno] = Me.accountno
8220              ![gfxid] = lngGfxID_Closed
8230              ![datbal] = Me.balance_date
8240              ![ballok] = True
8250            End With
8260            qdf.Execute
                ' ** Then reset the latest date to unlocked.
8270            Set qdf = .QueryDefs("qryStatementBalance_09")
8280            With qdf.Parameters
8290              ![actno] = Me.accountno
8300              ![gfxid] = lngGfxID_Open
8310              ![datbal] = Me.last_balance_date
8320              ![ballok] = False
8330            End With
8340            qdf.Execute
8350            .Close
8360          End With
8370        End If

8380        If .cmbAccounts.Column(2) = "X" Then
              ' ** Only reset if the account has records in the Balance table.

8390          If .frmStatementBalance_Sub.Enabled = False Then
8400            .balance_date.Enabled = True
8410            .TotalMarketValue.Enabled = True
8420            .AccountValue.Enabled = True
8430            .ICash.Enabled = True
8440            .PCash.Enabled = True
8450            .Cost.Enabled = True
8460            .frmStatementBalance_Sub.Enabled = True
8470          End If

8480          strTmp01 = .cmbAccounts.Column(0)
8490          strTmp02 = .cmbAccounts.Column(1)
8500          For lngX = 0& To (lngAccts - 1&)
8510            If arr_varAcct(A_NUM, lngX) = strTmp01 Then
8520              strTmp03 = Format(arr_varAcct(A_DAT, lngX), "mm/dd/yyyy")
8530              Exit For
8540            End If
8550          Next

8560          strFilter = "[accountno] = '" & strTmp01 & "' And [balance_date] = #" & strTmp03 & "#"
8570          .Filter = strFilter
8580          .FilterOn = True
              '.Refresh  CAUSES A LOCKUP, WITH ERROR 2486:  You can't carry out this action at the present time.

              ' ** Keep both combo boxes synchronized.
8590          .cmbShortName = strTmp02

8600        Else
8610          If .frmStatementBalance_Sub.Enabled = False Then
8620            .balance_date.Enabled = False
8630            .TotalMarketValue.Enabled = False
8640            .AccountValue.Enabled = False
8650            .ICash.Enabled = False
8660            .PCash.Enabled = False
8670            .Cost.Enabled = False
8680            .frmStatementBalance_Sub.Enabled = False
8690          End If
8700          MsgBox "No previous balance exists for this account.", vbInformation + vbOKOnly, "No Balance Found"
8710        End If
8720      End If
8730    End With

EXITP:
8740    Set qdf = Nothing
8750    Set dbs = Nothing
8760    Exit Sub

ERRH:
8770    Select Case ERR.Number
        Case Else
8780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8790    End Select
8800    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8920    intShiftDown = (Shift And acShiftMask) > 0
8930    intAltDown = (Shift And acAltMask) > 0
8940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8970        With Me
8980          intRetVal = 0
8990          .cmbShortName.SetFocus
9000        End With
9010      End Select
9020    End If

        ' ** Shift keys.
9030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9050        With Me
9060          intRetVal = 0
9070          If .cmdUpdate.Enabled = True Then
9080            .cmdUpdate.SetFocus
9090          ElseIf .cmdPrintReport.Enabled = True Then
9100            .cmdPrintReport.SetFocus
9110          Else
9120            lngRecsCur = .frmStatementBalance_Sub.Form.RecCnt  ' ** Form Function: frmStatementBalance_Sub.
9130            If lngRecsCur > 0& Then
9140              .frmStatementBalance_Sub.SetFocus
9150              .frmStatementBalance_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmStatementBalance_Sub.
9160              .frmStatementBalance_Sub.Form.AccountValue.SetFocus
9170            Else
9180              .AccountValue.SetFocus
9190            End If
9200          End If
9210        End With
9220      End Select
9230    End If

EXITP:
9240    KeyCode = intRetVal
9250    Exit Sub

ERRH:
9260    intRetVal = 0
9270    Select Case ERR.Number
        Case Else
9280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9290    End Select
9300    Resume EXITP

End Sub

Private Sub cmbShortName_AfterUpdate()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbShortName_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strFilter As String
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim lngX As Long

9410    With Me
9420      If IsNull(.cmbShortName) = False Then

9430        If .Dirty = True Then
9440          cmdSave_Click  ' ** Procedure: Above.
9450        End If

            ' ** If the current accountno is not on the latest
            ' ** balance_date, reset it before moving on.
9460        If .balance_date <> .last_balance_date Then
9470          Set dbs = CurrentDb
9480          With dbs
                ' ** Update tmpBalance, by specified [actno], [gfxid], [datbal], [ballok].
9490            Set qdf = .QueryDefs("qryStatementBalance_09")
9500            With qdf.Parameters
9510              ![actno] = Me.accountno
9520              ![gfxid] = lngGfxID_Closed
9530              ![datbal] = Me.balance_date
9540              ![ballok] = True
9550            End With
9560            qdf.Execute
                ' ** Then reset the latest date to unlocked.
9570            Set qdf = .QueryDefs("qryStatementBalance_09")
9580            With qdf.Parameters
9590              ![actno] = Me.accountno
9600              ![gfxid] = lngGfxID_Open
9610              ![datbal] = Me.last_balance_date
9620              ![ballok] = False
9630            End With
9640            qdf.Execute
9650            .Close
9660          End With
9670        End If

9680        If .cmbShortName.Column(2) = "X" Then
              ' ** Only reset if the account has records in the Balance table.

9690          If .frmStatementBalance_Sub.Enabled = False Then
9700            .balance_date.Enabled = True
9710            .TotalMarketValue.Enabled = True
9720            .AccountValue.Enabled = True
9730            .ICash.Enabled = True
9740            .PCash.Enabled = True
9750            .Cost.Enabled = True
9760            .frmStatementBalance_Sub.Enabled = True
9770          End If

9780          strTmp01 = .cmbShortName.Column(1)
9790          strTmp02 = .cmbShortName.Column(0)
9800          For lngX = 0& To (lngAccts - 1&)
9810            If arr_varAcct(A_NUM, lngX) = strTmp01 Then
9820              strTmp03 = Format(arr_varAcct(A_DAT, lngX), "mm/dd/yyyy")
9830              Exit For
9840            End If
9850          Next

9860          strFilter = "[accountno] = '" & strTmp01 & "' And [balance_date] = #" & strTmp03 & "#"
9870          .Filter = strFilter
9880          .FilterOn = True
              '.Refresh  CAUSES A LOCKUP, WITH ERROR 2486:  You can't carry out this action at the present time.

              ' ** Keep both combo boxes synchronized.
9890          .cmbAccounts = strTmp01

9900        Else
9910          If .frmStatementBalance_Sub.Enabled = False Then
9920            .balance_date.Enabled = False
9930            .TotalMarketValue.Enabled = False
9940            .AccountValue.Enabled = False
9950            .ICash.Enabled = False
9960            .PCash.Enabled = False
9970            .Cost.Enabled = False
9980            .frmStatementBalance_Sub.Enabled = False
9990          End If
10000         MsgBox "No previous balance exists for this account.", vbInformation + vbOKOnly, "No Balance Found"
10010       End If
10020     End If
10030   End With

EXITP:
10040   Set qdf = Nothing
10050   Set dbs = Nothing
10060   Exit Sub

ERRH:
10070   Select Case ERR.Number
        Case Else
10080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10090   End Select
10100   Resume EXITP

End Sub

Private Sub balance_date_entry_BeforeUpdate(Cancel As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "balance_date_entry_BeforeUpdate"

        Dim lngBalID As Long
        Dim datNewBalDate As Date
        Dim varTmp00 As Variant

10210   With Me
10220     Select Case IsNull(.balance_date_entry)
          Case True
            ' ** Let AfterUpdate() handle it.
10230     Case False
10240       Select Case IsNull(.balance_date_entry.text)
            Case True
              ' ** Let AfterUpdate() handle it.
10250       Case False
10260         Select Case IsDate(.balance_date_entry.text)
              Case True
10270           datNewBalDate = CDate(.balance_date_entry.text)
10280           lngBalID = .bal_id
10290           varTmp00 = DLookup("[bal_id]", "tmpBalance", _
                  "[balance_date] = #" & Format(datNewBalDate, "mm/dd/yyyy") & "# And " & _
                  "[bal_id] <> " & CStr(lngBalID))
10300           Select Case IsNull(varTmp00)
                Case True
                  ' ** All's well.
10310           Case False
10320             Cancel = -1
10330             Beep
10340             MsgBox "The date you've entered is already in the Balance table.", vbInformation + vbOKOnly, "Invalid Entry"
10350           End Select
10360         Case False
10370           If .balance_date_entry.text = vbNullString Then
                  ' ** Let AfterUpdate() handle it.
10380           Else
10390             Cancel = -1
10400             Beep
10410             MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, _
                    ("Invalid Date" & Space(50) & "03")
10420           End If
10430         End Select
10440       End Select
10450     End Select
10460   End With

EXITP:
10470   Exit Sub

ERRH:
10480   THAT_PROC = THIS_PROC
10490   That_Erl = Erl
10500   Select Case ERR.Number
        Case Else
10510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10520   End Select
10530   Resume EXITP

End Sub

Private Sub balance_date_entry_AfterUpdate()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "balance_date_entry_AfterUpdate"

10610   With Me
10620     If IsNull(.balance_date_entry) = True Then
10630       .balance_date_entry = Date
10640     End If
10650     .balance_date = .balance_date_entry
10660     cmdSave_Click  ' ** Procedure: Above.
10670   End With

EXITP:
10680   Exit Sub

ERRH:
10690   THAT_PROC = THIS_PROC
10700   That_Erl = Erl
10710   Select Case ERR.Number
        Case Else
10720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10730   End Select
10740   Resume EXITP

End Sub

Private Sub balance_date_entry_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "balance_date_entry_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim intRetVal As Integer

10810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10820   intShiftDown = (Shift And acShiftMask) > 0
10830   intAltDown = (Shift And acAltMask) > 0
10840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         intRetVal = 0
              ' ** I just want to get out of the damn field!
10890 On Error Resume Next
10900         strTmp01 = .balance_date_entry.text
10910         If ERR.Number <> 0 Then
10920 On Error GoTo ERRH
10930           strTmp02 = .balance_date_entry
10940         Else
10950 On Error GoTo ERRH
10960         End If
10970         If strTmp01 <> vbNullString Then
10980           Select Case IsDate(strTmp01)
                Case True
10990 On Error Resume Next
11000             .ICash.SetFocus
11010             If ERR.Number <> 0 Then
11020 On Error GoTo ERRH
11030               blnMoveSomewhere = True
11040               .TimerInterval = 500&
11050             Else
11060 On Error GoTo ERRH
11070             End If
11080           Case False
11090             Beep
11100             MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, _
                    ("Invalid Date" & Space(50) & "04")
11110             .balance_date_entry.Undo
11120           End Select
11130         Else
11140           Select Case IsDate(strTmp02)
                Case True
11150 On Error Resume Next
11160             .ICash.SetFocus
11170             If ERR.Number <> 0 Then
11180 On Error GoTo ERRH
11190               blnMoveSomewhere = True
11200               .TimerInterval = 500&
11210             Else
11220 On Error GoTo ERRH
11230             End If
11240           Case False
11250             Beep
11260             MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, _
                    ("Invalid Date" & Space(50) & "05")
11270             .balance_date_entry.Undo
11280           End Select
11290         End If
11300       End With
11310     End Select
11320   End If

        ' ** Shift keys.
11330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11350       With Me
11360         intRetVal = 0
11370         .cmdClose.SetFocus
11380       End With
11390     End Select
11400   End If

EXITP:
11410   KeyCode = intRetVal
11420   Exit Sub

ERRH:
11430   intRetVal = 0
11440   Select Case ERR.Number
        Case Else
11450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11460   End Select
11470   Resume EXITP

End Sub

Private Sub balance_date_GotFocus()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "balance_date_GotFocus"

11510   With Me
11520 On Error Resume Next
11530     .balance_date.SelLength = 0
11540 On Error GoTo ERRH
11550   End With

EXITP:
11560   Exit Sub

ERRH:
11570   Select Case ERR.Number
        Case Else
11580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11590   End Select
11600   Resume EXITP

End Sub

Private Sub balance_date_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "balance_date_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11720   intShiftDown = (Shift And acShiftMask) > 0
11730   intAltDown = (Shift And acAltMask) > 0
11740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11770       With Me
11780         intRetVal = 0
11790         .ICash.SetFocus
11800       End With
11810     End Select
11820   End If

        ' ** Shift keys.
11830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11850       With Me
11860         intRetVal = 0
11870         .cmdClose.SetFocus
11880       End With
11890     End Select
11900   End If

EXITP:
11910   KeyCode = intRetVal
11920   Exit Sub

ERRH:
11930   intRetVal = 0
11940   Select Case ERR.Number
        Case Else
11950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11960   End Select
11970   Resume EXITP

End Sub

Private Sub icash_GotFocus()

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_GotFocus"

12010   With Me
12020     If blnOpenFocus = True Then
12030       blnOpenFocus = False
12040       .ICash.SelLength = 0
12050     End If
12060   End With

EXITP:
12070   Exit Sub

ERRH:
12080   Select Case ERR.Number
        Case Else
12090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12100   End Select
12110   Resume EXITP

End Sub

Private Sub icash_Change()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "icash_Change"

12210   With Me
12220     .chkUpdated = True
12230     .cmdUpdate.Enabled = True
12240   End With

EXITP:
12250   Exit Sub

ERRH:
12260   Select Case ERR.Number
        Case Else
12270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12280   End Select
12290   Resume EXITP

End Sub

Private Sub pcash_Change()

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_Change"

12310   With Me
12320     .chkUpdated = True
12330     .cmdUpdate.Enabled = True
12340   End With

EXITP:
12350   Exit Sub

ERRH:
12360   Select Case ERR.Number
        Case Else
12370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12380   End Select
12390   Resume EXITP

End Sub

Private Sub cost_Change()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cost_Change"

12410   With Me
12420     .chkUpdated = True
12430     .cmdUpdate.Enabled = True
12440   End With

EXITP:
12450   Exit Sub

ERRH:
12460   Select Case ERR.Number
        Case Else
12470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12480   End Select
12490   Resume EXITP

End Sub

Private Sub TotalMarketValue_Change()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "TotalMarketValue_Change"

12510   With Me
12520     .chkUpdated = True
12530     .cmdUpdate.Enabled = True
12540   End With

EXITP:
12550   Exit Sub

ERRH:
12560   Select Case ERR.Number
        Case Else
12570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12580   End Select
12590   Resume EXITP

End Sub

Private Sub AccountValue_Change()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "AccountValue_Change"

12610   With Me
12620     .chkUpdated = True
12630     .cmdUpdate.Enabled = True
12640   End With

EXITP:
12650   Exit Sub

ERRH:
12660   Select Case ERR.Number
        Case Else
12670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12680   End Select
12690   Resume EXITP

End Sub

Private Sub AccountValue_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "AccountValue_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12720   intShiftDown = (Shift And acShiftMask) > 0
12730   intAltDown = (Shift And acAltMask) > 0
12740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12770       With Me
12780         intRetVal = 0
12790         lngRecsCur = .frmStatementBalance_Sub.Form.RecCnt  ' ** Form Function: frmStatementBalance_Sub.
12800         If lngRecsCur > 0 Then
12810           .frmStatementBalance_Sub.SetFocus
12820           .frmStatementBalance_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmStatementBalance_Sub.
12830           .frmStatementBalance_Sub.Form.cmdXAdgfx_image1.SetFocus
12840         Else
12850           If .cmdPrintReport.Enabled = True Then
12860             .cmdPrintReport.SetFocus
12870           ElseIf .cmdUpdate.Enabled = True Then
12880             .cmdUpdate.SetFocus
12890           ElseIf .cmbAccounts.Enabled = True Then
12900             .cmbAccounts.SetFocus
12910           Else
12920             .cmdClose.SetFocus
12930           End If
12940         End If
12950       End With
12960     End Select
12970   End If

        ' ** Shift keys.
12980   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12990     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13000       With Me
13010         intRetVal = 0
13020         .TotalMarketValue.SetFocus
13030       End With
13040     End Select
13050   End If

EXITP:
13060   KeyCode = intRetVal
13070   Exit Sub

ERRH:
13080   intRetVal = 0
13090   Select Case ERR.Number
        Case Else
13100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13110   End Select
13120   Resume EXITP

End Sub

Public Sub MoveRec(intWhere As AcCommand)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, ctl As Access.ComboBox
        Dim strAccountNo As String, strShortName As String
        Dim strFilter As String, strOrderBy As String
        Dim lngRecs As Long
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, lngTmp04 As Long
        Dim lngX As Long

13210   With Me
13220     Set ctl = .cmbAccounts
13230     If ctl.Visible = True Then

13240       strAccountNo = vbNullString
13250       lngRecs = .cmbAccounts.ListCount
13260       lngTmp04 = ctl.ListIndex

13270       Select Case intWhere
            Case acCmdRecordsGoToPrevious
13280         If lngTmp04 > 0& Then
13290           For lngX = (lngTmp04 - 1&) To 0& Step -1&
13300             If ctl.Column(2, lngX) = "X" Then
13310               strAccountNo = ctl.Column(0, lngX)
13320               strShortName = ctl.Column(1, lngX)
13330               Exit For
13340             End If
13350           Next
13360         End If
13370         If strAccountNo = vbNullString Then
                ' ** Go to the end of the list.
13380           For lngX = (lngRecs - 1&) To 0& Step -1&
13390             If ctl.Column(2, lngX) = "X" Then
13400               strAccountNo = ctl.Column(0, lngX)
13410               strShortName = ctl.Column(1, lngX)
13420               Exit For
13430             End If
13440           Next
13450         End If
13460       Case acCmdRecordsGoToNext
13470         If lngTmp04 < (lngRecs - 1&) Then
13480           For lngX = (lngTmp04 + 1&) To (lngRecs - 1&)
13490             If ctl.Column(2, lngX) = "X" Then
13500               strAccountNo = ctl.Column(0, lngX)
13510               strShortName = ctl.Column(1, lngX)
13520               Exit For
13530             End If
13540           Next
13550         End If
13560         If strAccountNo = vbNullString Then
                ' ** Go to the beginning of the list.
13570           For lngX = 0& To (lngRecs - 1&)
13580             If ctl.Column(2, lngX) = "X" Then
13590               strAccountNo = ctl.Column(0, lngX)
13600               strShortName = ctl.Column(1, lngX)
13610               Exit For
13620             End If
13630           Next
13640         End If
13650       End Select

13660       If strAccountNo <> vbNullString Then

13670         If .Dirty = True Then
13680           cmdSave_Click  ' ** Procedure: Above.
13690         End If

              ' ** If the current accountno is not on the latest
              ' ** balance_date, reset it before moving on.
13700         If .balance_date <> .last_balance_date Then
13710           Set dbs = CurrentDb
13720           With dbs
                  ' ** Update tmpBalance, by specified [actno], [gfxid], [datbal], [ballok].
13730             Set qdf = .QueryDefs("qryStatementBalance_09")
13740             With qdf.Parameters
13750               ![actno] = Me.accountno
13760               ![gfxid] = lngGfxID_Closed
13770               ![datbal] = Me.balance_date
13780               ![ballok] = True
13790             End With
13800             qdf.Execute
                  ' ** Then reset the latest date to unlocked.
13810             Set qdf = .QueryDefs("qryStatementBalance_09")
13820             With qdf.Parameters
13830               ![actno] = Me.accountno
13840               ![gfxid] = lngGfxID_Open
13850               ![datbal] = Me.last_balance_date
13860               ![ballok] = False
13870             End With
13880             qdf.Execute
13890             .Close
13900           End With
13910         End If

13920         If .frmStatementBalance_Sub.Enabled = False Then
13930           .balance_date.Enabled = True
13940           .TotalMarketValue.Enabled = True
13950           .AccountValue.Enabled = True
13960           .ICash.Enabled = True
13970           .PCash.Enabled = True
13980           .Cost.Enabled = True
13990           .frmStatementBalance_Sub.Enabled = True
14000         End If

14010         For lngX = 0& To (lngAccts - 1&)
14020           If arr_varAcct(A_NUM, lngX) = strAccountNo Then
14030             strTmp01 = strAccountNo
14040             strTmp02 = strShortName
14050             strTmp03 = Format(arr_varAcct(A_DAT, lngX), "mm/dd/yyyy")
14060             Exit For
14070           End If
14080         Next

              ' ** Keep both combo boxes synchronized.
14090         .cmbAccounts = strTmp01
14100         .cmbShortName = strTmp02

14110         strFilter = "[accountno] = '" & strTmp01 & "' And [balance_date] = #" & strTmp03 & "#"
14120         .Filter = strFilter
14130         .FilterOn = True

14140         strOrderBy = "[balance_date] DESC"
14150         .frmStatementBalance_Sub.Form.OrderBy = strOrderBy
14160         .frmStatementBalance_Sub.Form.OrderByOn = True

14170       End If  ' ** strAccountNo.

14180     End If
14190   End With

EXITP:
14200   Set ctl = Nothing
14210   Set qdf = Nothing
14220   Set dbs = Nothing
14230   Exit Sub

ERRH:
14240   Select Case ERR.Number
        Case Else
14250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14260   End Select
14270   Resume EXITP

End Sub

Public Sub ChangeDate(datNewBalDate As Date, lngBalID As Long)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "ChangeDate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strAccountNo As String, strNewBalDate As String
        Dim strFilter As String

14310   With Me

14320     strAccountNo = .accountno
14330     strNewBalDate = Format(datNewBalDate, "mm/dd/yyyy")
14340     strFilter = "[accountno] = '" & strAccountNo & "' And [balance_date] = #" & strNewBalDate & "#"

14350     Set dbs = CurrentDb
14360     With dbs
            ' ** Update tmpBalance, by specified [actno], [gfxid], [ballok].
14370       Set qdf = .QueryDefs("qryStatementBalance_08")
14380       With qdf.Parameters
14390         ![actno] = strAccountNo
14400         ![gfxid] = lngGfxID_Closed
14410         ![ballok] = True
14420       End With
14430       qdf.Execute
            ' ** Update tmpBalance, by specified [actno], [gfxid], [datbal], [ballok].
14440       Set qdf = .QueryDefs("qryStatementBalance_09")
14450       With qdf.Parameters
14460         ![actno] = strAccountNo
14470         ![gfxid] = lngGfxID_Open
14480         ![datbal] = datNewBalDate
14490         ![ballok] = False
14500       End With
14510       qdf.Execute
14520       .Close
14530     End With

14540     .Filter = strFilter
14550     .FilterOn = True
14560     DoEvents

14570   End With

EXITP:
14580   Set qdf = Nothing
14590   Set dbs = Nothing
14600   Exit Sub

ERRH:
14610   Select Case ERR.Number
        Case 2455  ' ** You entered an expression that has an invalid reference to the property Form/Report.
14620     If strAccountNo <> "ALL" Then
14630       Me.Filter = "[accountno] = '" & strAccountNo & "' And [balance_date] = #" & strLastDate & "#"
14640     Else
            ' ** This shouldn't happen, but if it does, this at least should leave something showing on the screen.
14650       Me.Filter = vbNullString
14660     End If
14670   Case Else
14680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14690   End Select
14700   Resume EXITP

End Sub
