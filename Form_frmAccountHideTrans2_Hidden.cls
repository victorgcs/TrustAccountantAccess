VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountHideTrans2_Hidden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountHideTrans2_Hidden"

'VGC 09/07/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Check Again:      A {cmdCheckAgain}
' **   Close:            C {cmdClose}
' **   Show JournalNo:   J {chkShowJournalNo}
' **   Print Report:     P {cmdPrintReport}
' **   Exit:             X {cmdClose}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

'Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngGrpNum_Left As Long, lngSharefaceSub_Left As Long, lngICashSub_Left As Long
Private lngPCashSub_Left As Long, lngCostSub_Left As Long, lngCurrID_Width As Long, lngCurrID_Left As Long
Private lngSubTotLbl_Offset As Long, lngPrintReportLbl_Left As Long

Private blnGoingToReport As Boolean, blnHasForEx_Chk As Boolean, blnUpdateTotals As Boolean
Private lngHids As Long, lngGrps As Long, datPriorPeriodMax As Date
Private lngRecsCur As Long, lngTpp As Long, blnRequery As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim strFilter As String
        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           intPos01 = InStr(Trim(.OpenArgs), "~")
150           If intPos01 > 0 Then

                ' ** THIS_NAME ~ .accountno ~ .hidden_trans ~ gdatEndDate
160             strTmp01 = Trim(.OpenArgs)
170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strTmp01 = Mid(strTmp01, (intPos01 + 1))
190             intPos01 = InStr(strTmp01, "~")
200             gstrAccountNo = Left(strTmp01, (intPos01 - 1))
210             strTmp01 = Mid(strTmp01, (intPos01 + 1))
220             intPos01 = InStr(strTmp01, "~")
230             lngHids = CLng(Left(strTmp01, (intPos01 - 1)))
240             datPriorPeriodMax = CDate(Mid(strTmp01, (intPos01 + 1)))

250             CLR_DISABLED_FG = CLR_DKGRY
260             CLR_DISABLED_BG = CLR_LTTEAL

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
270             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

280             gblnHasForExThis = False: blnUpdateTotals = False
290             .CallingForm = strCallingForm

300             lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left
310             lngGrpNum_Left = .ledghid_grpnum.Left
320             lngSharefaceSub_Left = .shareface_subtot.Left
330             lngICashSub_Left = .icash_subtot.Left
340             lngPCashSub_Left = .pcash_subtot.Left
350             lngCostSub_Left = .cost_subtot.Left
360             lngCurrID_Width = (.curr_id.Width + (4& * lngTpp))
370             lngCurrID_Left = .curr_id.Left
380             lngSubTotLbl_Offset = (.shareface_subtot.Left - (.subtot_lbl.Left + .subtot_lbl.Width))

390             If gblnGoToReport = True Then
400               .cmdPrintReport_lbl2.Left = (lngPrintReportLbl_Left + .GoToReport_arw_pr_img.Width)
410               .GoToReport_arw_pr_img.Visible = True
420               DoEvents
430             End If

440             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

450             strFilter = "[accountno] = '" & gstrAccountNo & "'"
460             .frmAccountHideTrans2_Hidden_Sub.Form.Filter = strFilter
470             .frmAccountHideTrans2_Hidden_Sub.Form.FilterOn = True

480             .PriorPeriodMax = datPriorPeriodMax
490             gdatEndDate = datPriorPeriodMax

500             gstrFormQuerySpec = THIS_NAME
510             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
520             blnRequery = False

530             chkShowJournalNo_AfterUpdate  ' ** Procedure: Below.

540             blnHasForEx_Chk = True
550             .cmdCheckAgain.Enabled = False

560             DoCmd.Hourglass True  ' ** Make sure it's still running.
570             DoEvents
580             .TimerInterval = 100&

590           Else
600             Cancel = -1
610           End If
620         Else
630           Cancel = -1
640         End If
650       Else
660         Cancel = -1
670       End If
680     End With

690     If Cancel = -1 Then
700       DoCmd.Hourglass False
710       Beep
720       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
730     End If

EXITP:
740     Exit Sub

ERRH:
750     DoCmd.Hourglass False
760     Select Case ERR.Number
        Case Else
770       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
780     End Select
790     Resume EXITP

End Sub

Private Sub Form_Load()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

810     With Me

820       .accountno.ForeColor = CLR_DISABLED_FG
830       .accountno.BackColor = CLR_DISABLED_BG
840       .shortname.ForeColor = CLR_DISABLED_FG
850       .shortname.BackColor = CLR_DISABLED_BG
860       DoEvents

870       Set dbs = CurrentDb
          ' ** qryAccountHideTrans2_Hidden_06 (qryAccountHideTrans2_Hidden_05 (tblLedgerHidden,
          ' ** linked to qryAccountHideTrans2_Hidden_04 (Union of qryAccountHideTrans2_Hidden_04a
          ' ** (Ledger, linked to MasterAsset, just ledger_HIDDEN = True, by specified
          ' ** FormRef('AcctNum')), qryAccountHideTrans2_Hidden_04b (LedgerArchive, linked
          ' ** to MasterAsset, just ledger_HIDDEN = True, by specified FormRef('AcctNum'))),
          ' ** with totdesc), with shareface_0, shareface_1, shareface_2, shareface_3,
          ' ** shareface_4, by specified FormRef('PriorPeriod')), grouped by ledghid_grpnum.
880       Set qdf = dbs.QueryDefs("qryAccountHideTrans2_Hidden_07")
890       Set rst = qdf.OpenRecordset
900       If rst.BOF = True And rst.EOF = True Then
910         lngGrps = 0&
920       Else
930         rst.MoveLast
940         lngGrps = rst.RecordCount
950       End If
960       rst.Close
970       dbs.Close

980       .ledghid_groups = lngGrps
990       DoEvents

1000      gstrFormQuerySpec = THIS_NAME

1010    End With

EXITP:
1020    Set rst = Nothing
1030    Set qdf = Nothing
1040    Set dbs = Nothing
1050    Exit Sub

ERRH:
1060    DoCmd.Hourglass False
1070    Select Case ERR.Number
        Case Else
1080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1090    End Select
1100    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1210    intShiftDown = (Shift And acShiftMask) > 0
1220    intAltDown = (Shift And acAltMask) > 0
1230    intCtrlDown = (Shift And acCtrlMask) > 0

1240    intRetVal = KeyCode

        ' ** Alt keys.
1250    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyX
1270        intRetVal = 0
1280        cmdClose_Click  ' ** Procedure: Below.
1290      End Select
1300    End If

        ' ** Ctrl-Shift keys.
1310    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1320      Select Case intRetVal
          Case vbKeyF
1330        With Me
1340          intRetVal = 0
1350          .FocusHolder.SetFocus
1360        End With
1370      End Select
1380    End If

EXITP:
1390    KeyCode = intRetVal
1400    Exit Sub

ERRH:
1410    intRetVal = 0
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Private Sub Form_Timer()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1510    With Me
1520      .TimerInterval = 0&
1530      If gblnSetFocus = True Then
1540        gblnSetFocus = False
1550        strFormName = vbNullString
1560  On Error Resume Next
1570        Do Until strFormName = THIS_NAME
1580          DoCmd.SelectObject acForm, THIS_NAME, False
1590          strFormName = Screen.ActiveForm.Name
1600        Loop
1610  On Error GoTo ERRH
1620      End If
1630      If blnRequery = True Then
1640        blnRequery = False
1650  On Error Resume Next
1660        .frmAccountHideTrans2_Hidden_Sub.Form.RecordSource = "qryAccountHideTrans2_Hidden_08"
1670        DoEvents
1680        .frmAccountHideTrans2_Hidden_Sub.Form.Requery
1690        DoEvents
1700        .frmAccountHideTrans2_Hidden_Sub.Form.OrderBy = "[ledghid_grpnum], [ledghid_ord]"
1710        .frmAccountHideTrans2_Hidden_Sub.Form.OrderByOn = True
1720        .frmAccountHideTrans2_Hidden_Sub.SetFocus
1730        .frmAccountHideTrans2_Hidden_Sub.Form.FocusHolder.SetFocus
1740        If ERR.Number <> 0 Then
1750  On Error GoTo ERRH
1760          blnRequery = True
1770          .TimerInterval = 750&
1780        Else
1790  On Error GoTo ERRH
1800        End If
1810      End If
1820      If gblnGoToReport = True Then
1830        Select Case blnGoingToReport
            Case True
1840          gblnGoToReport = False
1850          blnGoingToReport = False
1860          .GoToReport_arw_pr_img.Visible = False
1870          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1880          gblnSetFocus = True
1890          DoCmd.Hourglass False  ' ** Should already be off.
1900          .TimerInterval = 100&
1910        Case False
1920          DoCmd.Hourglass True  ' ** Make sure it's still running.
1930          DoEvents
1940          blnGoingToReport = True
1950          If .cmdPrintReport.Enabled = True Then
1960            .cmdPrintReport.SetFocus
1970          Else
1980            .cmdClose.SetFocus
1990          End If
2000          .GoToReport_arw_pr_img.Visible = True
              ' ** This may already be off.
2010          Forms(strCallingForm).GTREmblem_Set False  ' ** Form Procedure: frmAccountHideTrans2.
2020          DoCmd.SelectObject acForm, THIS_NAME, False
2030          DoEvents
2040          Beep
2050          DoCmd.Hourglass False
2060          DoEvents
2070          gblnSetFocus = True
2080          .TimerInterval = (GTR_WAIT * 2&)
2090        End Select
2100      End If
2110      If blnUpdateTotals = True Then
2120        blnUpdateTotals = False
2130        lngRecsCur = .frmAccountHideTrans2_Hidden_Sub.Form.RecCnt  ' ** Form Function: frmAccountHideTrans2_Hidden_Sub.
2140        If lngRecsCur < lngHids Then
2150          .cmdCheckAgain.Enabled = True
2160        End If
2170      End If
2180      If blnHasForEx_Chk = True Then
2190        blnHasForEx_Chk = False
2200        gblnHasForExThis = HasForEx_Acct(gstrAccountNo, "L")  ' ** Module Function: modCurrencyFuncs.
2210        Select Case gblnHasForEx
            Case True
2220          Select Case gblnHasForExThis
              Case True
2230            .frmAccountHideTrans2_Hidden_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountHideTrans2_Hidden_Sub.
2240            IncludeCurrency True  ' ** Procedure: Below.
2250          Case False
2260            .frmAccountHideTrans2_Hidden_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountHideTrans2_Hidden_Sub.
2270            IncludeCurrency False  ' ** Procedure: Below.
2280          End Select
2290        Case False
2300          .frmAccountHideTrans2_Hidden_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountHideTrans2_Hidden_Sub.
2310          IncludeCurrency False  ' ** Procedure: Below.
2320        End Select
2330        chkShowJournalNo_AfterUpdate  ' ** Procedure: Below.
2340        DoCmd.Hourglass False
2350        blnUpdateTotals = True
2360        .TimerInterval = 200&
2370      End If
2380    End With

EXITP:
2390    Exit Sub

ERRH:
2400    DoCmd.Hourglass False
2410    Select Case ERR.Number
        Case Else
2420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2430    End Select
2440    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2510    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2520    If Reports.Count > 0 Then
2530      Do While Reports.Count > 0
2540        DoCmd.Close acReport, Reports(0).Name
2550        DoEvents
2560      Loop
2570    End If

2580    Select Case strCallingForm
        Case "frmAccountHideTrans2"
2590      gblnSetFocus = True
2600      gstrReturningForm = THIS_NAME
2610      Select Case IsLoaded(strCallingForm, acForm)  ' ** Module Function: modFileUtilities.
          Case True
2620        DoCmd.SelectObject acForm, strCallingForm, False
2630        Forms(strCallingForm).TimerInterval = 250&
2640      Case False
2650        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Account"
            'NEEDS OPEN ARGS!!!
2660      End Select
2670    End Select

EXITP:
2680    Exit Sub

ERRH:
2690    Select Case ERR.Number
        Case Else
2700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2710    End Select
2720    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2810    With Me
2820      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2830        Select Case blnPrintReport_Focus
            Case True
2840          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2850          .cmdPrintReport_raised_img.Visible = False
2860        Case False
2870          .cmdPrintReport_raised_img.Visible = True
2880          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2890        End Select
2900        .cmdPrintReport_raised_focus_img.Visible = False
2910        .cmdPrintReport_raised_focus_dots_img.Visible = False
2920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2930        .cmdPrintReport_raised_img_dis.Visible = False
2940      End If
2950    End With

EXITP:
2960    Exit Sub

ERRH:
2970    Select Case ERR.Number
        Case Else
2980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2990    End Select
3000    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Public Sub cmdClose_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3210    DoCmd.Close acForm, THIS_NAME

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
3310    intShiftDown = (Shift And acShiftMask) > 0
3320    intAltDown = (Shift And acAltMask) > 0
3330    intCtrlDown = (Shift And acCtrlMask) > 0

3340    intRetVal = KeyCode

        ' ** Plain keys.
3350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3360      Select Case intRetVal
          Case vbKeyTab
3370        With Me
3380          intRetVal = 0
3390          .frmAccountHideTrans2_Hidden_Sub.SetFocus
3400          .frmAccountHideTrans2_Hidden_Sub.Form.FocusHolder.SetFocus
3410        End With
3420      End Select
3430    End If

        ' ** Shift keys.
3440    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3450      Select Case intRetVal
          Case vbKeyTab
3460        With Me
3470          intRetVal = 0
3480          If .cmdPrintReport.Enabled = True Then
3490            .cmdPrintReport.SetFocus
3500          Else
3510            .chkShowJournalNo.SetFocus
3520          End If
3530        End With
3540      End Select
3550    End If

EXITP:
3560    KeyCode = intRetVal
3570    Exit Sub

ERRH:
3580    intRetVal = 0
3590    Select Case ERR.Number
        Case Else
3600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3610    End Select
3620    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3710    With Me
3720      blnPrintReport_Focus = True
3730      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3740      .cmdPrintReport_raised_img.Visible = False
3750      .cmdPrintReport_raised_focus_img.Visible = False
3760      .cmdPrintReport_raised_focus_dots_img.Visible = False
3770      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3780      .cmdPrintReport_raised_img_dis.Visible = False
3790    End With

EXITP:
3800    Exit Sub

ERRH:
3810    Select Case ERR.Number
        Case Else
3820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3830    End Select
3840    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3910    With Me
3920      blnPrintReport_MouseDown = True
3930      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3940      .cmdPrintReport_raised_img.Visible = False
3950      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3960      .cmdPrintReport_raised_focus_img.Visible = False
3970      .cmdPrintReport_raised_focus_dots_img.Visible = False
3980      .cmdPrintReport_raised_img_dis.Visible = False
3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    Select Case ERR.Number
        Case Else
4020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4030    End Select
4040    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strSortNow As String, strSortOrig As String
        Dim intMode As Integer

4110    With Me
4120      If .hidden_trans > 0& Then
4130        strSortNow = .frmAccountHideTrans2_Hidden_Sub.Form.SortNow_Get  ' ** Form Function: frmAccountHideTrans2_Hidden_Sub.
4140        strSortOrig = .frmAccountHideTrans2_Hidden_Sub.Form.SortOrig_Get  ' ** Form Function: frmAccountHideTrans2_Hidden_Sub.
4150        If strSortNow = strSortOrig Then
4160          strDocName = "rptTransaction_Hidden"
4170        Else
4180          strDocName = "rptTransaction_Hidden_Hidden"
4190        End If
4200        If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4210          intMode = acViewPreview
4220        Else
4230          intMode = acViewNormal
4240        End If
            '##GTR_Ref: rptTransaction_Hidden
4250        DoCmd.OpenReport strDocName, intMode, , , , .accountno
4260        If intMode = acViewPreview Then
4270          DoCmd.Maximize
4280          DoCmd.RunCommand acCmdFitToWindow
4290          Forms(THIS_NAME).Visible = False
4300          Forms(strCallingForm).Visible = False
4310        End If
4320      Else
4330        Beep
4340      End If
4350    End With

EXITP:
4360    Exit Sub

ERRH:
4370    Select Case ERR.Number
        Case Else
4380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4390    End Select
4400    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
4510    intShiftDown = (Shift And acShiftMask) > 0
4520    intAltDown = (Shift And acAltMask) > 0
4530    intCtrlDown = (Shift And acCtrlMask) > 0

4540    intRetVal = KeyCode

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab
4570        With Me
4580          intRetVal = 0
4590          .cmdClose.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab
4650        With Me
4660          intRetVal = 0
4670          .chkShowJournalNo.SetFocus
4680        End With
4690      End Select
4700    End If

EXITP:
4710    KeyCode = intRetVal
4720    Exit Sub

ERRH:
4730    intRetVal = 0
4740    Select Case ERR.Number
        Case Else
4750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4760    End Select
4770    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4810    With Me
4820      If blnPrintReport_MouseDown = False Then
4830        Select Case blnPrintReport_Focus
            Case True
4840          .cmdPrintReport_raised_focus_dots_img.Visible = True
4850          .cmdPrintReport_raised_focus_img.Visible = False
4860        Case False
4870          .cmdPrintReport_raised_focus_img.Visible = True
4880          .cmdPrintReport_raised_focus_dots_img.Visible = False
4890        End Select
4900        .cmdPrintReport_raised_img.Visible = False
4910        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4930        .cmdPrintReport_raised_img_dis.Visible = False
4940      End If
4950    End With

EXITP:
4960    Exit Sub

ERRH:
4970    Select Case ERR.Number
        Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

5110    With Me
5120      .cmdPrintReport_raised_focus_dots_img.Visible = True
5130      .cmdPrintReport_raised_img.Visible = False
5140      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5150      .cmdPrintReport_raised_focus_img.Visible = False
5160      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5170      .cmdPrintReport_raised_img_dis.Visible = False
5180      blnPrintReport_MouseDown = False
5190    End With

EXITP:
5200    Exit Sub

ERRH:
5210    Select Case ERR.Number
        Case Else
5220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5230    End Select
5240    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

5310    With Me
5320      .cmdPrintReport_raised_img.Visible = True
5330      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5340      .cmdPrintReport_raised_focus_img.Visible = False
5350      .cmdPrintReport_raised_focus_dots_img.Visible = False
5360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5370      .cmdPrintReport_raised_img_dis.Visible = False
5380      blnPrintReport_Focus = False
5390    End With

EXITP:
5400    Exit Sub

ERRH:
5410    Select Case ERR.Number
        Case Else
5420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5430    End Select
5440    Resume EXITP

End Sub

Private Sub chkShowJournalNo_AfterUpdate()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowJournalNo_AfterUpdate"

5510    With Me
5520      DoCmd.Hourglass True
5530      DoEvents
5540      Select Case .chkShowJournalNo
          Case True
5550        .chkShowJournalNo_lbl.FontBold = True
5560        .frmAccountHideTrans2_Hidden_Sub.Form.ShowJournalNo_Sub True  ' ** Form Procedure: frmAccountHideTrans2_Hidden_Sub.
5570      Case False
5580        .chkShowJournalNo_lbl.FontBold = False
5590        .frmAccountHideTrans2_Hidden_Sub.Form.ShowJournalNo_Sub False  ' ** Form Procedure: frmAccountHideTrans2_Hidden_Sub.
5600      End Select
5610      .frmAccountHideTrans2_Hidden_Sub.SetFocus
5620      DoCmd.Hourglass False
5630    End With

EXITP:
5640    Exit Sub

ERRH:
5650    DoCmd.Hourglass False
5660    Select Case ERR.Number
        Case Else
5670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5680    End Select
5690    Resume EXITP

End Sub

Private Sub chkShowJournalNo_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowJournalNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
5710    intShiftDown = (Shift And acShiftMask) > 0
5720    intAltDown = (Shift And acAltMask) > 0
5730    intCtrlDown = (Shift And acCtrlMask) > 0

5740    intRetVal = KeyCode

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab
5770        With Me
5780          intRetVal = 0
5790          If .cmdPrintReport.Enabled = True Then
5800            .cmdPrintReport.SetFocus
5810          Else
5820            .cmdClose.SetFocus
5830          End If
5840        End With
5850      End Select
5860    End If

        ' ** Shift keys.
5870    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5880      Select Case intRetVal
          Case vbKeyTab
5890        With Me
5900          intRetVal = 0
5910          .frmAccountHideTrans2_Hidden_Sub.SetFocus
5920          .frmAccountHideTrans2_Hidden_Sub.Form.FocusHolder.SetFocus
5930        End With
5940      End Select
5950    End If

EXITP:
5960    KeyCode = intRetVal
5970    Exit Sub

ERRH:
5980    intRetVal = 0
5990    Select Case ERR.Number
        Case Else
6000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6010    End Select
6020    Resume EXITP

End Sub

Private Sub cmdCheckAgain_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCheckAgain_Click"

6110    HiddenCheck  ' ** Procedure: Below

EXITP:
6120    Exit Sub

ERRH:
6130    DoCmd.Hourglass False
6140    Select Case ERR.Number
        Case Else
6150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6160    End Select
6170    Resume EXITP

End Sub

Private Sub IncludeCurrency(blnShow As Boolean)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

6210    With Me
6220      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
6230        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6240      End If
6250      Select Case blnShow
          Case True
6260        .ledghid_grpnum.Left = lngGrpNum_Left
6270        .shareface_subtot.Left = lngSharefaceSub_Left
6280        .icash_subtot.Left = lngICashSub_Left
6290        .pcash_subtot.Left = lngPCashSub_Left
6300        .cost_subtot.Left = lngCostSub_Left
6310        .curr_id.Left = lngCurrID_Left
6320        .curr_id_lbl.Left = .curr_id.Left
6330        .curr_id.Visible = True
6340        .curr_id_lbl.Visible = True
6350      Case False
6360        .ledghid_grpnum.Left = (lngGrpNum_Left + lngCurrID_Width)
6370        .shareface_subtot.Left = (lngSharefaceSub_Left + lngCurrID_Width)
6380        .icash_subtot.Left = (lngICashSub_Left + lngCurrID_Width)
6390        .pcash_subtot.Left = (lngPCashSub_Left + lngCurrID_Width)
6400        .cost_subtot.Left = (lngCostSub_Left + lngCurrID_Width)
6410        .curr_id.Visible = False
6420        .curr_id_lbl.Visible = False
6430        .curr_id.Left = 0&
6440        .curr_id_lbl.Left = 0&
6450      End Select
6460      .ledghid_grpnum_lbl.Left = (.ledghid_grpnum.Left - (ledghid_grpnum_lbl.Width + lngSubTotLbl_Offset))
6470      .subtot_lbl.Left = (.shareface_subtot.Left - (.subtot_lbl.Width + lngSubTotLbl_Offset))
6480      .shareface_subtot_lbl.Left = .shareface_subtot.Left
6490      .icash_subtot_lbl.Left = .icash_subtot.Left
6500      .pcash_subtot_lbl.Left = .pcash_subtot.Left
6510      .cost_subtot_lbl.Left = .cost_subtot.Left
6520      .IsInvalid_box.Left = (.shareface_subtot.Left - lngTpp)
6530      .IsInvalid_lbl_box.Left = (.IsInvalid_box.Left - ((.IsInvalid_lbl_box.Width - .IsInvalid_box.Width) / 2))
6540      .IsInvalid_lbl1.Left = .IsInvalid_lbl_box.Left
6550      .IsInvalid_lbl2.Left = .IsInvalid_lbl_box.Left

6560    End With

EXITP:
6570    Exit Sub

ERRH:
6580    Select Case ERR.Number
        Case Else
6590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6600    End Select
6610    Resume EXITP

End Sub

Private Sub HiddenCheck()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "HiddenCheck"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngHidns As Long, arr_varHidn As Variant
        Dim lngItems As Long, arr_varItem() As Variant
        Dim lngSkips As Long, arr_varSkip() As Variant
        Dim lngHoles As Long, arr_varHole() As Variant
        Dim lngAssetNo As Long, lngMaxGrp As Long, lngRecs As Long
        Dim lngGoodNums As Long, lngBadNums As Long, lngHidType As Long
        Dim strUniqueID As String
        Dim blnContinue As Boolean, blnFound As Boolean
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long
        Dim lngX As Long, lngY As Long, lngE As Long

        ' ** Array: arr_varHidn().
        Const H_JNO   As Integer = 0
        Const H_JTYPE As Integer = 1
        Const H_ACTNO As Integer = 2
        Const H_ASTNO As Integer = 3
        Const H_HID   As Integer = 4
        Const H_TDAT  As Integer = 5
        Const H_ADAT  As Integer = 6
        Const H_PDAT  As Integer = 7
        Const H_SHRS  As Integer = 8
        Const H_ICSH  As Integer = 9
        Const H_PCSH  As Integer = 10
        Const H_COST  As Integer = 11
        Const H_SRC   As Integer = 12

        ' ** Array: arr_varItem().
        Const I_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const I_JNO1   As Integer = 0
        Const I_JNO2   As Integer = 1
        Const I_ACTNO  As Integer = 2
        Const I_ASTNO  As Integer = 3
        Const I_HELEM1 As Integer = 4
        Const I_HELEM2 As Integer = 5

        ' ** Array: arr_varSkip().
        Const S_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const S_JNO As Integer = 0

        ' ** Array: arr_varHole().
        Const O_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const O_GRP As Integer = 0

6710    With Me

6720      DoCmd.Hourglass True
6730      DoEvents

6740      blnContinue = True

          ' ** This is rather simplistic matching, and, at least for now,
          ' ** won't catch those really odd or complex ones.

6750      Set dbs = CurrentDb
6760      With dbs

            ' ** qryAccountHidden_02 (qryAccountHidden_01 (Union of qryAccountHidden_01_01
            ' ** (Ledger, just ledger_HIDDEN = True), qryAccountHidden_01_02 (LedgerArchive, just
            ' ** ledger_HIDDEN = True), all hidden entries), not in tblLedgerHidden), by specified [actno].
6770        Set qdf = .QueryDefs("qryAccountHidden_03")
6780        With qdf.Parameters
6790          ![actno] = gstrAccountNo
6800        End With
6810        Set rst = qdf.OpenRecordset
6820        If rst.BOF = True And rst.EOF = True Then
              ' ** All present and accounted for.
6830          blnContinue = False
6840          rst.Close
6850          Set rst = Nothing
6860          Set qdf = Nothing
6870        Else
              ' ** Hidden found not in this form's listing.
6880          rst.MoveLast
6890          lngHidns = rst.RecordCount
6900          rst.MoveFirst
6910          arr_varHidn = rst.GetRows(lngHidns)
              ' **************************************************
              ' ** Array: arr_varHidn()
              ' **
              ' **   Field  Element  Name             Constant
              ' **   =====  =======  ===============  ==========
              ' **     1       0     journalno        H_JNO
              ' **     2       1     journaltype      H_JTYPE
              ' **     3       2     accountno        H_ACTNO
              ' **     4       3     assetno          H_ASTNO
              ' **     5       4     ledger_HIDDEN    H_HID
              ' **     6       5     transdate        H_TDAT
              ' **     7       6     assetdate        H_ADAT
              ' **     8       7     PurchaseDate     H_PDAT
              ' **     9       8     shareface        H_SHRS
              ' **    10       9     icash            H_ICSH
              ' **    11      10     pcash            H_PCSH
              ' **    12      11     cost             H_COST
              ' **    13      12     src              H_SRC
              ' **
              ' **************************************************
6920        End If

6930        If blnContinue = True Then

6940          lngItems = 0&
6950          ReDim arr_varItem(I_ELEMS, 0)

6960          lngSkips = 0&
6970          ReDim arr_varSkip(S_ELEMS, 0)

              'ALL RIGHT, HOW DO I WANT TO DO THIS?
              '1ST, COMPARE ASSET NO'S.

              ' ************************************
              ' ** Check assetno transactions.
              ' ************************************
6980          Do While blnContinue = True

6990            lngAssetNo = 0&
7000            lngTmp01 = 0&: lngTmp02 = 0&

7010            For lngX = 0& To (lngHidns - 1&)
7020              If arr_varHidn(H_ASTNO, lngX) > 0& Then
7030                blnFound = False
7040                For lngY = 0& To (lngItems - 1&)
                      ' ** Check against what's already been done.
7050                  If arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO1, lngY) Or arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO2, lngY) Then
7060                    blnFound = True
7070                    Exit For
7080                  End If
7090                Next
7100                If blnFound = False Then
7110                  For lngY = 0& To (lngSkips - 1&)
                        ' ** Check against deadends.
7120                    If arr_varHidn(H_JNO, lngX) = arr_varSkip(S_JNO, lngY) Then
7130                      blnFound = True
7140                      Exit For
7150                    End If
7160                  Next
7170                End If
7180                If blnFound = False Then
7190                  lngAssetNo = arr_varHidn(H_ASTNO, lngX)
7200                  lngTmp01 = lngX
7210                  Exit For
7220                End If
7230              End If
7240            Next  ' ** lngX.

7250            If lngAssetNo > 0 Then
                  ' ** At least 1 assetno present.

                  ' ** Try for a simple matching pair.
7260              For lngX = 0& To (lngHidns - 1&)
7270                If lngX <> lngTmp01 Then
7280                  blnFound = False
7290                  For lngY = 0& To (lngItems - 1&)
                        ' ** Check against what's already been done.
7300                    If arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO1, lngY) Or arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO2, lngY) Then
7310                      blnFound = True
7320                      Exit For
7330                    End If
7340                  Next
7350                  If blnFound = False Then
7360                    For lngY = 0& To (lngSkips - 1&)
                          ' ** Check against deadends.
7370                      If arr_varHidn(H_JNO, lngX) = arr_varSkip(S_JNO, lngY) Then
7380                        blnFound = True
7390                        Exit For
7400                      End If
7410                    Next
7420                  End If
7430                  If blnFound = False Then
7440                    If arr_varHidn(H_ASTNO, lngX) = lngAssetNo Then
                          ' ** A 2nd, matching assetno found.
7450                      If arr_varHidn(H_JTYPE, lngTmp01) = "Deposit" And arr_varHidn(H_JTYPE, lngX) = "Withdrawn" Then
7460                        blnFound = True  ' ** Provisional!
7470                        lngTmp02 = lngX
7480                        Exit For
7490                      ElseIf arr_varHidn(H_JTYPE, lngTmp01) = "Purchase" And arr_varHidn(H_JTYPE, lngX) = "Sold" Then
7500                        blnFound = True  ' ** Provisional!
7510                        lngTmp02 = lngX
7520                        Exit For
7530                      ElseIf arr_varHidn(H_JTYPE, lngTmp01) = "Withdrawn" And arr_varHidn(H_JTYPE, lngX) = "Deposit" Then
7540                        blnFound = True  ' ** Provisional!
7550                        lngTmp02 = lngX
7560                        Exit For
7570                      ElseIf arr_varHidn(H_JTYPE, lngTmp01) = "Sold" And arr_varHidn(H_JTYPE, lngX) = "Purchase" Then
7580                        blnFound = True  ' ** Provisional!
7590                        lngTmp02 = lngX
7600                        Exit For
7610                      ElseIf arr_varHidn(H_JTYPE, lngTmp01) = "Cost Adj." Or arr_varHidn(H_JTYPE, lngX) = "Cost Adj." Then
7620                        blnFound = True  ' ** Provisional!
7630                        lngTmp02 = lngX
7640                        Exit For
7650                      ElseIf arr_varHidn(H_JTYPE, lngTmp01) = "Liability" Or arr_varHidn(H_JTYPE, lngX) = "Liability" Then
                            'YUK!

7660                      Else
                            ' ** What else?

7670                      End If
7680                    End If
7690                  End If
7700                End If
7710              Next  ' ** lngX.

7720              If blnFound = True Then
7730                If arr_varHidn(H_JTYPE, lngTmp01) = "Cost Adj." Then
                      ' ** Shareface should be Zero, so just match cost.

7740                  If Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngTmp02), 2) = 0 Then
                        ' ** We've got a match!
7750                    lngItems = lngItems + 1&
7760                    lngE = lngItems - 1&
7770                    ReDim Preserve arr_varItem(I_ELEMS, lngE)
7780                    arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
7790                    arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
7800                    arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
7810                    arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
7820                    arr_varItem(I_HELEM1, lngE) = lngTmp01
7830                    arr_varItem(I_HELEM2, lngE) = lngTmp02
7840                  Else
                        ' ** Is there another Cost Adj. in arr_varHidn()?
7850                    For lngX = 0& To (lngHidns - 1&)
7860                      If lngX <> lngTmp01 And lngX <> lngTmp02 Then
7870                        blnFound = False
7880                        For lngY = 0& To (lngItems - 1&)
                              ' ** Check against what's already been done.
7890                          If arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO1, lngY) Or arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO2, lngY) Then
7900                            blnFound = True
7910                            Exit For
7920                          End If
7930                        Next
7940                        If blnFound = False Then
7950                          For lngY = 0& To (lngSkips - 1&)
                                ' ** Check against deadends.
7960                            If arr_varHidn(H_JNO, lngX) = arr_varSkip(S_JNO, lngY) Then
7970                              blnFound = True
7980                              Exit For
7990                            End If
8000                          Next
8010                        End If
8020                        If blnFound = False Then
8030                          If arr_varHidn(H_JTYPE, lngX) = "Cost Adj." Then
8040                            If Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngX), 2) = 0 Then
                                  ' ** Got it!
8050                              blnFound = True
8060                              lngTmp02 = lngX
8070                              Exit For
8080                            End If
8090                          End If
8100                        End If
8110                      End If
8120                    Next
8130                    If blnFound = True Then
8140                      lngItems = lngItems + 1&
8150                      lngE = lngItems - 1&
8160                      ReDim Preserve arr_varItem(I_ELEMS, lngE)
8170                      arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
8180                      arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
8190                      arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
8200                      arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
8210                      arr_varItem(I_HELEM1, lngE) = lngTmp01
8220                      arr_varItem(I_HELEM2, lngE) = lngTmp02
8230                    Else
                          ' ** lngTmp01 is now an unmatched Cost Adj., and lngTmp02 is ignored.
8240                      lngSkips = lngSkips + 1&
8250                      lngE = lngSkips - 1&
8260                      ReDim Preserve arr_varSkip(S_ELEMS, lngE)
8270                      arr_varSkip(S_JNO, lngE) = arr_varHidn(H_JNO, lngTmp01)
8280                    End If
8290                  End If

8300                ElseIf arr_varHidn(H_JTYPE, lngTmp01) = "Deposit" Then
                      ' ** Other is Withdrawn.

8310                  If (Round(arr_varHidn(H_SHRS, lngTmp01), 4) = Round(arr_varHidn(H_SHRS, lngTmp02), 4)) And _
                          (Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngTmp02), 2) = 0) Then
                        ' ** Good match.
8320                    lngItems = lngItems + 1&
8330                    lngE = lngItems - 1&
8340                    ReDim Preserve arr_varItem(I_ELEMS, lngE)
8350                    arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
8360                    arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
8370                    arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
8380                    arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
8390                    arr_varItem(I_HELEM1, lngE) = lngTmp01
8400                    arr_varItem(I_HELEM2, lngE) = lngTmp02
8410                  Else
                        ' ** Keep looking.
8420                    For lngX = 0& To (lngHidns - 1&)
8430                      If lngX <> lngTmp01 And lngX <> lngTmp02 Then
8440                        blnFound = False
8450                        For lngY = 0& To (lngItems - 1&)
                              ' ** Check against what's already been done.
8460                          If arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO1, lngY) Or arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO2, lngY) Then
8470                            blnFound = True
8480                            Exit For
8490                          End If
8500                        Next
8510                        If blnFound = False Then
8520                          For lngY = 0& To (lngSkips - 1&)
                                ' ** Check against deadends.
8530                            If arr_varHidn(H_JNO, lngX) = arr_varSkip(S_JNO, lngY) Then
8540                              blnFound = True
8550                              Exit For
8560                            End If
8570                          Next
8580                        End If
8590                        If blnFound = False Then
8600                          If arr_varHidn(H_JTYPE, lngX) = "Withdrawn" Then
8610                            If (Round(arr_varHidn(H_SHRS, lngTmp01), 4) = Round(arr_varHidn(H_SHRS, lngX), 4)) And _
                                    (Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngX), 2) = 0) Then
                                  ' ** Got it!
8620                              blnFound = True
8630                              lngTmp02 = lngX
8640                              Exit For
8650                            End If
8660                          End If
8670                        End If
8680                      End If
8690                    Next
8700                    If blnFound = True Then
8710                      lngItems = lngItems + 1&
8720                      lngE = lngItems - 1&
8730                      ReDim Preserve arr_varItem(I_ELEMS, lngE)
8740                      arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
8750                      arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
8760                      arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
8770                      arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
8780                      arr_varItem(I_HELEM1, lngE) = lngTmp01
8790                      arr_varItem(I_HELEM2, lngE) = lngTmp02
8800                    Else
                          ' ** lngTmp01 is now an unmatched Deposit, and lngTmp02 is ignored.
8810                      lngSkips = lngSkips + 1&
8820                      lngE = lngSkips - 1&
8830                      ReDim Preserve arr_varSkip(S_ELEMS, lngE)
8840                      arr_varSkip(S_JNO, lngE) = arr_varHidn(H_JNO, lngTmp01)
8850                    End If
8860                  End If

8870                ElseIf arr_varHidn(H_JTYPE, lngTmp01) = "Purchase" Then
                      ' ** Other is Sold.

8880                  If (Round(arr_varHidn(H_SHRS, lngTmp01), 4) = Round(arr_varHidn(H_SHRS, lngTmp02), 4)) And _
                          (Round(arr_varHidn(H_ICSH, lngTmp01), 2) + Round(arr_varHidn(H_ICSH, lngTmp02), 2) = 0) And _
                          (Round(arr_varHidn(H_PCSH, lngTmp01), 2) + Round(arr_varHidn(H_PCSH, lngTmp02), 2) = 0) And _
                          (Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngTmp02), 2) = 0) Then
                        ' ** Good match.
8890                    lngItems = lngItems + 1&
8900                    lngE = lngItems - 1&
8910                    ReDim Preserve arr_varItem(I_ELEMS, lngE)
8920                    arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
8930                    arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
8940                    arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
8950                    arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
8960                    arr_varItem(I_HELEM1, lngE) = lngTmp01
8970                    arr_varItem(I_HELEM2, lngE) = lngTmp02
8980                  Else
                        ' ** Keep looking.
8990                    For lngX = 0& To (lngHidns - 1&)
9000                      If lngX <> lngTmp01 And lngX <> lngTmp02 Then
9010                        blnFound = False
9020                        For lngY = 0& To (lngItems - 1&)
                              ' ** Check against what's already been done.
9030                          If arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO1, lngY) Or arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO2, lngY) Then
9040                            blnFound = True
9050                            Exit For
9060                          End If
9070                        Next
9080                        If blnFound = False Then
9090                          For lngY = 0& To (lngSkips - 1&)
                                ' ** Check against deadends.
9100                            If arr_varHidn(H_JNO, lngX) = arr_varSkip(S_JNO, lngY) Then
9110                              blnFound = True
9120                              Exit For
9130                            End If
9140                          Next
9150                        End If
9160                        If blnFound = False Then
9170                          If arr_varHidn(H_JTYPE, lngX) = "Sold" Then
9180                            If (Round(arr_varHidn(H_SHRS, lngTmp01), 4) = Round(arr_varHidn(H_SHRS, lngX), 4)) And _
                                    (Round(arr_varHidn(H_ICSH, lngTmp01), 2) + Round(arr_varHidn(H_ICSH, lngX), 2) = 0) And _
                                    (Round(arr_varHidn(H_PCSH, lngTmp01), 2) + Round(arr_varHidn(H_PCSH, lngX), 2) = 0) And _
                                    (Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngX), 2) = 0) Then
                                  ' ** Got it!
9190                              blnFound = True
9200                              lngTmp02 = lngX
9210                              Exit For
9220                            End If
9230                          End If
9240                        End If
9250                      End If
9260                    Next
9270                    If blnFound = True Then
9280                      lngItems = lngItems + 1&
9290                      lngE = lngItems - 1&
9300                      ReDim Preserve arr_varItem(I_ELEMS, lngE)
9310                      arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
9320                      arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
9330                      arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
9340                      arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
9350                      arr_varItem(I_HELEM1, lngE) = lngTmp01
9360                      arr_varItem(I_HELEM2, lngE) = lngTmp02
9370                    Else
                          ' ** lngTmp01 is now an unmatched Purchase, and lngTmp02 is ignored.
9380                      lngSkips = lngSkips + 1&
9390                      lngE = lngSkips - 1&
9400                      ReDim Preserve arr_varSkip(S_ELEMS, lngE)
9410                      arr_varSkip(S_JNO, lngE) = arr_varHidn(H_JNO, lngTmp01)
9420                    End If
9430                  End If

9440                ElseIf arr_varHidn(H_JTYPE, lngTmp01) = "Withdrawn" Then
                      ' ** Other is Deposit.

9450                  If (Round(arr_varHidn(H_SHRS, lngTmp01), 4) = Round(arr_varHidn(H_SHRS, lngTmp02), 4)) And _
                          (Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngTmp02), 2) = 0) Then
                        ' ** Good match.
9460                    lngItems = lngItems + 1&
9470                    lngE = lngItems - 1&
9480                    ReDim Preserve arr_varItem(I_ELEMS, lngE)
9490                    arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
9500                    arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
9510                    arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
9520                    arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
9530                    arr_varItem(I_HELEM1, lngE) = lngTmp01
9540                    arr_varItem(I_HELEM2, lngE) = lngTmp02
9550                  Else
                        ' ** Keep looking.
9560                    For lngX = 0& To (lngHidns - 1&)
9570                      If lngX <> lngTmp01 And lngX <> lngTmp02 Then
9580                        blnFound = False
9590                        For lngY = 0& To (lngItems - 1&)
                              ' ** Check against what's already been done.
9600                          If arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO1, lngY) Or arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO2, lngY) Then
9610                            blnFound = True
9620                            Exit For
9630                          End If
9640                        Next
9650                        If blnFound = False Then
9660                          For lngY = 0& To (lngSkips - 1&)
                                ' ** Check against deadends.
9670                            If arr_varHidn(H_JNO, lngX) = arr_varSkip(S_JNO, lngY) Then
9680                              blnFound = True
9690                              Exit For
9700                            End If
9710                          Next
9720                        End If
9730                        If blnFound = False Then
9740                          If arr_varHidn(H_JTYPE, lngX) = "Deposit" Then
9750                            If (Round(arr_varHidn(H_SHRS, lngTmp01), 4) = Round(arr_varHidn(H_SHRS, lngX), 4)) And _
                                    (Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngX), 2) = 0) Then
                                  ' ** Got it!
9760                              blnFound = True
9770                              lngTmp02 = lngX
9780                              Exit For
9790                            End If
9800                          End If
9810                        End If
9820                      End If
9830                    Next
9840                    If blnFound = True Then
9850                      lngItems = lngItems + 1&
9860                      lngE = lngItems - 1&
9870                      ReDim Preserve arr_varItem(I_ELEMS, lngE)
9880                      arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
9890                      arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
9900                      arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
9910                      arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
9920                      arr_varItem(I_HELEM1, lngE) = lngTmp01
9930                      arr_varItem(I_HELEM2, lngE) = lngTmp02
9940                    Else
                          ' ** lngTmp01 is now an unmatched Withdrawn, and lngTmp02 is ignored.
9950                      lngSkips = lngSkips + 1&
9960                      lngE = lngSkips - 1&
9970                      ReDim Preserve arr_varSkip(S_ELEMS, lngE)
9980                      arr_varSkip(S_JNO, lngE) = arr_varHidn(H_JNO, lngTmp01)
9990                    End If
10000                 End If

10010               ElseIf arr_varHidn(H_JTYPE, lngTmp01) = "Sold" Then
                      ' ** Other is Purchase.

10020                 If (Round(arr_varHidn(H_SHRS, lngTmp01), 4) = Round(arr_varHidn(H_SHRS, lngTmp02), 4)) And _
                          (Round(arr_varHidn(H_ICSH, lngTmp01), 2) + Round(arr_varHidn(H_ICSH, lngTmp02), 2) = 0) And _
                          (Round(arr_varHidn(H_PCSH, lngTmp01), 2) + Round(arr_varHidn(H_PCSH, lngTmp02), 2) = 0) And _
                          (Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngTmp02), 2) = 0) Then
                        ' ** Good match.
10030                   lngItems = lngItems + 1&
10040                   lngE = lngItems - 1&
10050                   ReDim Preserve arr_varItem(I_ELEMS, lngE)
10060                   arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
10070                   arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
10080                   arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
10090                   arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
10100                   arr_varItem(I_HELEM1, lngE) = lngTmp01
10110                   arr_varItem(I_HELEM2, lngE) = lngTmp02
10120                 Else
                        ' ** Keep looking.
10130                   For lngX = 0& To (lngHidns - 1&)
10140                     If lngX <> lngTmp01 And lngX <> lngTmp02 Then
10150                       blnFound = False
10160                       For lngY = 0& To (lngItems - 1&)
                              ' ** Check against what's already been done.
10170                         If arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO1, lngY) Or arr_varHidn(H_JNO, lngX) = arr_varItem(I_JNO2, lngY) Then
10180                           blnFound = True
10190                           Exit For
10200                         End If
10210                       Next
10220                       If blnFound = False Then
10230                         For lngY = 0& To (lngSkips - 1&)
                                ' ** Check against deadends.
10240                           If arr_varHidn(H_JNO, lngX) = arr_varSkip(S_JNO, lngY) Then
10250                             blnFound = True
10260                             Exit For
10270                           End If
10280                         Next
10290                       End If
10300                       If blnFound = False Then
10310                         If arr_varHidn(H_JTYPE, lngX) = "Purchase" Then
10320                           If (Round(arr_varHidn(H_SHRS, lngTmp01), 4) = Round(arr_varHidn(H_SHRS, lngX), 4)) And _
                                    (Round(arr_varHidn(H_ICSH, lngTmp01), 2) + Round(arr_varHidn(H_ICSH, lngX), 2) = 0) And _
                                    (Round(arr_varHidn(H_PCSH, lngTmp01), 2) + Round(arr_varHidn(H_PCSH, lngX), 2) = 0) And _
                                    (Round(arr_varHidn(H_COST, lngTmp01), 2) + Round(arr_varHidn(H_COST, lngX), 2) = 0) Then
                                  ' ** Got it!
10330                             blnFound = True
10340                             lngTmp02 = lngX
10350                             Exit For
10360                           End If
10370                         End If
10380                       End If
10390                     End If
10400                   Next
10410                   If blnFound = True Then
10420                     lngItems = lngItems + 1&
10430                     lngE = lngItems - 1&
10440                     ReDim Preserve arr_varItem(I_ELEMS, lngE)
10450                     arr_varItem(I_JNO1, lngE) = arr_varHidn(H_JNO, lngTmp01)
10460                     arr_varItem(I_JNO2, lngE) = arr_varHidn(H_JNO, lngTmp02)
10470                     arr_varItem(I_ACTNO, lngE) = arr_varHidn(H_ACTNO, lngTmp01)
10480                     arr_varItem(I_ASTNO, lngE) = arr_varHidn(H_ASTNO, lngTmp01)
10490                     arr_varItem(I_HELEM1, lngE) = lngTmp01
10500                     arr_varItem(I_HELEM2, lngE) = lngTmp02
10510                   Else
                          ' ** lngTmp01 is now an unmatched Sold, and lngTmp02 is ignored.
10520                     lngSkips = lngSkips + 1&
10530                     lngE = lngSkips - 1&
10540                     ReDim Preserve arr_varSkip(S_ELEMS, lngE)
10550                     arr_varSkip(S_JNO, lngE) = arr_varHidn(H_JNO, lngTmp01)
10560                   End If
10570                 End If

10580               Else
                      'LIABILITY

10590               End If  ' ** H_JTYPE.

10600             Else
                    ' ** Only 1? Is this a Dividend or Interest?

10610             End If  ' ** ASTNO.

10620           Else
                  ' ** Loop till there are no more assetno's.
10630             blnContinue = False
10640             Exit Do
10650           End If  ' ** lngAssetNo.

10660         Loop  ' ** blnContinue.

10670         blnContinue = True

              ' ************************************
              ' ** Check non-assetno transactions.
              ' ************************************

              'Debug.Print "'lngHidns:  " & CStr(lngHidns)
              'Debug.Print "'lngItems: " & CStr(lngItems)
              'Debug.Print "'lngSkips: " & CStr(lngSkips)

10680         If lngItems > 0& Then

10690           lngHoles = 0&
10700           ReDim arr_varHole(O_ELEMS, 0)

10710           lngMaxGrp = DMax("[ledghid_grpnum]", "tblLedgerHidden", "[accountno] = '" & gstrAccountNo & "'")
10720           If lngMaxGrp >= 1000& Then
                  ' ** At least 1 group didn't get renumbered.

10730             lngTmp01 = 0&: lngTmp02 = 0&: lngTmp03 = 0&

                  ' ** qryAccountHidden_04_01 (tblLedgerHidden, just ledghid_grpnum >= 1000,
                  ' ** by specified [actno]), grouped by ledghid_grpnum, with cnt_jno.
10740             Set qdf = .QueryDefs("qryAccountHidden_04_02")
10750             With qdf.Parameters
10760               ![actno] = gstrAccountNo
10770             End With
10780             Set rst = qdf.OpenRecordset
10790             With rst
10800               .MoveLast
10810               lngBadNums = .RecordCount  ' ** Number of groups needing renumbering.
10820               .MoveFirst
10830               .Close
10840             End With
10850             Set rst = Nothing
10860             Set qdf = Nothing

                  ' ** qryAccountHidden_04_03 (tblLedgerHidden, just ledghid_grpnum < 1000,
                  ' ** by specified [actno]), grouped by ledghid_grpnum, with cnt_jno.
10870             Set qdf = .QueryDefs("qryAccountHidden_04_04")
10880             With qdf.Parameters
10890               ![actno] = gstrAccountNo
10900             End With
10910             Set rst = qdf.OpenRecordset
10920             With rst
10930               If .BOF = True And .EOF = True Then
                      ' ** All need renumbering
10940                 lngGoodNums = 0&
10950               Else
10960                 .MoveLast
10970                 lngGoodNums = .RecordCount  ' ** Number of groups correctly numbered.
10980                 .MoveFirst
10990                 lngTmp03 = 0&
11000                 For lngX = 1& To lngGoodNums
11010                   If ![ledghid_grpnum] = lngTmp03 + 1& Then
11020                     lngTmp03 = ![ledghid_grpnum]
11030                   Else
                          ' ** A hole in the sequence.
11040                     Do Until lngTmp03 + 1& = ![ledghid_grpnum]
11050                       lngHoles = lngHoles + 1&
11060                       lngE = lngHoles - 1&
11070                       ReDim Preserve arr_varHole(O_ELEMS, lngE)
11080                       arr_varHole(O_GRP, lngE) = (lngTmp03 + 1&)
11090                       lngTmp03 = lngTmp03 + 1&
11100                       If lngTmp03 + 1& = ![ledghid_grpnum] Then
11110                         lngTmp03 = ![ledghid_grpnum]
11120                         Exit Do
11130                       End If
11140                     Loop
11150                   End If
11160                   If lngX < lngGoodNums Then .MoveNext
11170                 Next  ' ** lngX.
11180               End If
11190               .Close
11200             End With
11210             Set rst = Nothing
11220             Set qdf = Nothing

11230             If lngHoles > 0& Then
                    ' ** If lngHoles = lngTmp01, great; use arr_varHole() to number groups below.
                    ' ** If lngHoles > lngTmp01, weird; just use them up.
                    ' ** If lngHoles < lngTmp01, messy; use up, then continue with new max grp.
                    ' ** tblLedgerHidden, just ledghid_grpnum >= 1000, by specified [actno].
11240               Set qdf = .QueryDefs("qryAccountHidden_04_01")
11250               With qdf.Parameters
11260                 ![actno] = gstrAccountNo
11270               End With
11280               Set rst = qdf.OpenRecordset
11290               With rst
11300                 .MoveLast
11310                 lngRecs = .RecordCount
11320                 .MoveFirst
11330                 lngTmp01 = 0&: lngTmp02 = 0&: lngTmp03 = -1&
11340                 lngMaxGrp = 0&
11350                 For lngX = 1& To lngRecs
11360                   If lngMaxGrp = 0& Then
11370                     If ![ledghid_grpnum] = lngTmp01 Or lngTmp01 = 0& Then
                            ' ** First one or same group as last record.
11380                       lngTmp01 = ![ledghid_grpnum]
11390                       For lngY = 0& To (lngHoles - 1&)
11400                         If arr_varHole(O_GRP, lngY) > 0& Then
11410                           lngTmp02 = arr_varHole(O_GRP, lngY)
11420                           lngTmp03 = lngY
11430                           Exit For
11440                         End If
11450                       Next
11460                     Else
                            ' ** New group.
11470                       arr_varHole(O_GRP, lngTmp03) = 0&  ' ** Zero-out the used one.
11480                       lngTmp03 = -1&
11490                       lngTmp01 = ![ledghid_grpnum]
11500                       For lngY = 0& To (lngHoles - 1&)
11510                         If arr_varHole(O_GRP, lngY) > 0& Then
11520                           lngTmp02 = arr_varHole(O_GRP, lngY)
11530                           lngTmp03 = lngY
11540                           Exit For
11550                         End If
11560                       Next
11570                     End If
11580                     If lngTmp02 = 0& Then
                            ' ** We've used up all the holes.
11590                       lngTmp01 = ![ledghid_grpnum]
11600                       lngMaxGrp = (lngGoodNums + lngHoles)
11610                       lngTmp02 = lngMaxGrp + 1&
11620                     End If
11630                   Else
11640                     If ![ledghid_grpnum] = lngTmp01 Then
11650                       lngTmp01 = ![ledghid_grpnum]
11660                       lngTmp02 = lngMaxGrp + 1&
11670                     Else
11680                       lngTmp01 = ![ledghid_grpnum]
11690                       lngMaxGrp = lngMaxGrp + 1&
11700                       lngTmp02 = lngMaxGrp + 1&
11710                     End If
11720                   End If
11730                   .Edit
11740                   ![ledghid_grpnum] = lngTmp02
11750                   ![ledghid_username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
11760                   ![ledghid_datemodified] = Now()
11770                   .Update
11780                   If lngX < lngRecs Then .MoveNext
11790                 Next
11800                 .Close
11810               End With
11820               Set rst = Nothing
11830               Set qdf = Nothing
11840               lngMaxGrp = lngTmp02

11850             Else

11860               lngMaxGrp = lngGoodNums  ' ** No holes, so max is record count.
                    ' ** tblLedgerHidden, just ledghid_grpnum >= 1000, by specified [actno].
11870               Set qdf = .QueryDefs("qryAccountHidden_04_01")
11880               With qdf.Parameters
11890                 ![actno] = gstrAccountNo
11900               End With
11910               Set rst = qdf.OpenRecordset
11920               With rst
11930                 .MoveLast
11940                 lngRecs = .RecordCount
11950                 .MoveFirst
11960                 lngTmp01 = 0&: lngTmp02 = 0&
11970                 For lngX = 1& To lngRecs
11980                   If ![ledghid_grpnum] = lngTmp01 Or lngTmp01 = 0& Then
                          ' ** First one or same group as last record.
11990                     lngTmp01 = ![ledghid_grpnum]
12000                     lngTmp02 = (lngMaxGrp + 1&)
12010                   Else
                          ' ** New group.
12020                     lngMaxGrp = lngMaxGrp + 1&  ' ** Max is now last group's number.
12030                     lngTmp01 = ![ledghid_grpnum]
12040                     lngTmp02 = (lngMaxGrp + 1&)
12050                   End If
12060                   .Edit
12070                   ![ledghid_grpnum] = lngTmp02
12080                   ![ledghid_username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
12090                   ![ledghid_datemodified] = Now()
12100                   .Update
12110                   If lngX < lngRecs Then .MoveNext
12120                 Next
12130                 .Close
12140               End With
12150               Set rst = Nothing
12160               Set qdf = Nothing
12170               lngMaxGrp = lngTmp02
12180             End If  ' ** lngHoles.

12190           End If  ' ** lngMaxGrp.

                ' ** So far, these are all NORM hidden types.
12200           lngHidType = 1&

12210           Set rst = .OpenRecordset("tblLedgerHidden", dbOpenDynaset, dbConsistent)
12220           With rst
12230             For lngX = 0& To (lngItems - 1&)
                    ' ** Each item will generate 2 entries.

                    ' ** 000000000000040_0150_002578_002617_Purchase__Sold_____
                    ' ** 000000000000232_1780_155307_155309_Cost Adj._Cost Adj.
12240               strUniqueID = Right(String(15, "0") & arr_varItem(I_ACTNO, lngX), 15) & "_"
12250               strUniqueID = strUniqueID & Right(String(4, "0") & CStr(arr_varItem(I_ASTNO, lngX)), 4) & "_"
12260               strUniqueID = strUniqueID & Right(String(6, "0") & CStr(arr_varItem(I_JNO1, lngX)), 6) & "_"
12270               strUniqueID = strUniqueID & Right(String(6, "0") & CStr(arr_varItem(I_JNO2, lngX)), 6) & "_"
12280               strUniqueID = strUniqueID & Left(arr_varHidn(H_JTYPE, arr_varItem(I_HELEM1, lngX)) & String(9, "_"), 9) & "_"
12290               strUniqueID = strUniqueID & Left(arr_varHidn(H_JTYPE, arr_varItem(I_HELEM2, lngX)) & String(9, "_"), 9)

12300               .AddNew
                    ' ** ![ledghid_id] : AutoNumber.
12310               ![journalno] = arr_varItem(I_JNO1, lngX)
12320               ![accountno] = arr_varItem(I_ACTNO, lngX)
12330               ![assetno] = arr_varItem(I_ASTNO, lngX)
12340               ![transdate] = arr_varHidn(H_TDAT, arr_varItem(I_HELEM1, lngX))
12350               ![ledghid_cnt] = 2
12360               ![ledghid_grpnum] = (lngMaxGrp + 1&)
12370               ![ledghid_ord] = 1&
12380               ![ledghidtype_type] = lngHidType
12390               ![ledghid_uniqueid] = strUniqueID
12400               ![ledghid_username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
12410               ![ledghid_datemodified] = Now()
12420               .Update

12430               .AddNew
                    ' ** ![ledghid_id] : AutoNumber.
12440               ![journalno] = arr_varItem(I_JNO2, lngX)
12450               ![accountno] = arr_varItem(I_ACTNO, lngX)
12460               ![assetno] = arr_varItem(I_ASTNO, lngX)
12470               ![transdate] = arr_varHidn(H_TDAT, arr_varItem(I_HELEM2, lngX))
12480               ![ledghid_cnt] = 2
12490               ![ledghid_grpnum] = (lngMaxGrp + 1&)
12500               ![ledghid_ord] = 2&
12510               ![ledghidtype_type] = lngHidType
12520               ![ledghid_uniqueid] = strUniqueID
12530               ![ledghid_username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
12540               ![ledghid_datemodified] = Now()
12550               .Update

12560               lngMaxGrp = lngMaxGrp + 1&
12570             Next
12580             .Close
12590           End With
12600           Set rst = Nothing

12610         End If  ' ** lngItems.

12620       Else
              ' ** None missing.
12630         Beep
12640         DoCmd.Hourglass False
12650         MsgBox "No missing hidden groups were found for this account.", vbInformation + vbOKOnly, "Nothing To Do"
12660       End If

12670       .Close
12680     End With
12690     Set dbs = Nothing

12700     .frmAccountHideTrans2_Hidden_Sub.Form.Requery

12710     Beep
12720     DoCmd.Hourglass False

12730   End With

EXITP:
12740   Set rst = Nothing
12750   Set qdf = Nothing
12760   Set dbs = Nothing
12770   Exit Sub

ERRH:
12780   DoCmd.Hourglass False
12790   Select Case ERR.Number
        Case Else
12800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12810   End Select
12820   Resume EXITP

End Sub
