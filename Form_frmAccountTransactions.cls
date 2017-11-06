VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountTransactions"

'VGC 04/20/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   IncludeCurrency()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Include Archive:  H {chkIncludeArchive}
' **   Include Currency: U {chkIncludeCurrency}
' **   Exit:             X {cmdClose}

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngTotWidth_Diff As Long, lngClose_Left As Long, lngSub_Width As Long
Private lngDetailHLine01_Width As Long, lngDetailHLine02_Width As Long, lngDetailVLine03_Left As Long

Private blnHasForEx As Boolean, blnHasForExThis As Boolean
Private strAccountNo As String, strShortName As String
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim strFilter As String
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = Trim(.OpenArgs)
150           If CharCnt(strTmp01, "~") = 2 Then  ' ** Module Function: modStringFuncs.
                ' ** strCallingForm ~ accountno ~ shortname

160             blnIsOpen = True

170             gstrFormQuerySpec = THIS_NAME
180             blnHasForEx = False: blnHasForExThis = False

190             intPos01 = InStr(strTmp01, "~")
200             strCallingForm = Left(strTmp01, (intPos01 - 1))
210             strTmp01 = Mid(strTmp01, (intPos01 + 1))
220             intPos01 = InStr(strTmp01, "~")
230             strAccountNo = Left(strTmp01, (intPos01 - 1))
240             strShortName = Mid(strTmp01, (intPos01 + 1))

250             .accountno = strAccountNo
260             .shortname = strShortName

270             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

280             SetArchiveOption  ' ** Procedure: Below.

290             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

300             chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.

310             strFilter = "[accountno] = '" & strAccountNo & "'"
320             .frmAccountTransactions_Sub.Form.Filter = strFilter
330             .frmAccountTransactions_Sub.Form.FilterOn = True

340             .frmAccountTransactions_Sub.Form.SortNow "Form_Load"  ' ** Form Procedure: frmAccountTransactions_Sub.

350             blnIsOpen = False

360           Else
370             Cancel = -1
380           End If
390         Else
400           Cancel = -1
410         End If
420       Else
430         Cancel = -1
440       End If
450     End With

460     DoCmd.Hourglass False

470     If Cancel = -1 Then
480       Beep
490       DoCmd.Hourglass False
500       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
510     End If

EXITP:
520     Exit Sub

ERRH:
530     DoCmd.Hourglass False
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form

610     With Me

620       Set dbs = CurrentDb
630       With dbs
            ' ** Check for foreign currencies anywhere.
            ' ** qryAccountProfile_Transactions_06_03 (Union of qryAccountProfile_Transactions_06_01
            ' ** (Ledger, just curr_id <> 150), qryAccountProfile_Transactions_06_02
            ' ** (LedgerArchive, just curr_id <> 150)), grouped, with cnt.
640         Set qdf = .QueryDefs("qryAccountProfile_Transactions_06_04")
650         Set rst = qdf.OpenRecordset
660         With rst
670           If .BOF = True And .EOF = True Then
680             blnHasForEx = False
690           Else
700             .MoveFirst
710             If ![cnt] = 0 Then
720               blnHasForEx = False
730             Else
740               blnHasForEx = True
750             End If
760           End If
770           .Close
780         End With
790         Set rst = Nothing
800         Set qdf = Nothing
810         If blnHasForEx = False Then
              ' ** MasterAsset, grouped, just curr_id <> 150, with cnt.
820           Set qdf = .QueryDefs("qryAccountProfile_Transactions_08")
830           Set rst = qdf.OpenRecordset
840           With rst
850             If .BOF = True And .EOF = True Then
                  ' ** Same as Zero.
860             Else
870               .MoveFirst
880               If ![cnt] > 0 Then
890                 blnHasForEx = True
900               End If
910             End If
920             .Close
930           End With
940           Set rst = Nothing
950           Set qdf = Nothing
960         End If
970         If blnHasForEx = True Then
              ' ** Check for foreign currencies in this account.
              ' ** qryAccountProfile_Transactions_03 (Union of qryAccountProfile_Transactions_01_02
              ' ** (qryAccountProfile_Transactions_01_01 (Ledger, linked toi tblCurrency, with
              ' ** 'Liability (+)'/'Liability (-)' broken out, totdescx, by specified GlobalVarGet('gstrAccountNo')),
              ' ** linked to Account, MasterAsset, tblJournalType), qryAccountProfile_Transactions_02_02
              ' ** (qryAccountProfile_Transactions_02_01 (LedgerArchive, linked to tblCurrency, with
              ' ** 'Liability (+)'/'Liability (-)' broken out, totdescx, by specified GlobalVarGet('gstrAccountNo')),
              ' ** linked to Account, MasterAsset, tblJournalType)), grouped, with cnt.
980           Set qdf = .QueryDefs("qryAccountProfile_Transactions_07")
990           Set rst = qdf.OpenRecordset
1000          With rst
1010            If .BOF = True And .EOF = True Then
1020              blnHasForExThis = False
1030            Else
1040              .MoveFirst
1050              If ![cnt] = 0 Then
1060                blnHasForExThis = False
1070              Else
1080                blnHasForExThis = True
1090              End If
1100            End If
1110            .Close
1120          End With
1130          Set rst = Nothing
1140          Set qdf = Nothing
1150        End If
1160        .Close
1170      End With
1180      Set dbs = Nothing

          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1190      lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

          ' ** Variables are fed empty, then populated ByRef.
1200      GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

1210      FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

1220      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
1230      If lngMonitorCnt > 1& Then
1240        EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
1250        lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
1260        If lngMonitorNum > 1& Then
1270          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
1280        End If
1290      End If

1300      .form_width_line.Left = 0&  ' ** Just get it out of the way.

1310      Set frm = .frmAccountTransactions_Sub.Form

1320      lngForm_Width = .Width
1330      lngClose_Left = .cmdClose.Left
1340      lngSub_Width = .frmAccountTransactions_Sub.Width
1350      lngDetailHLine01_Width = .Detail_hline01.Width
1360      lngDetailHLine02_Width = .Detail_hline02.Width
1370      lngDetailVLine03_Left = .Detail_vline03.Left
1380      lngTotWidth_Diff = (((frm.cost_usd.Left + frm.cost_usd.Width) - (frm.Cost.Left + frm.Cost.Width)) - lngTpp)

1390      Select Case blnHasForEx
          Case True
1400        .chkIncludeCurrency.Enabled = True
1410        Select Case blnHasForExThis
            Case True
1420          With .frmAccountTransactions_Sub.Form
1430            .ICash.Visible = False
1440            .icash_str.Visible = True
1450            .PCash.Visible = False
1460            .pcash_str.Visible = True
1470            .Cost.Visible = False
1480            .cost_str.Visible = True
1490          End With
1500        Case False
1510          With .frmAccountTransactions_Sub.Form
1520            .ICash.Visible = True
1530            .icash_str.Visible = False
1540            .PCash.Visible = True
1550            .pcash_str.Visible = False
1560            .Cost.Visible = True
1570            .cost_str.Visible = False
1580          End With
1590        End Select
1600      Case False
1610        .chkIncludeCurrency = False
1620        .chkIncludeCurrency.Enabled = False
1630        With .frmAccountTransactions_Sub.Form
1640          .ICash.Visible = True
1650          .icash_str.Visible = False
1660          .PCash.Visible = True
1670          .pcash_str.Visible = False
1680          .Cost.Visible = True
1690          .cost_str.Visible = False
1700        End With
1710      End Select
1720      chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
1730      DoEvents

1740      .frmAccountTransactions_Sub.SetFocus
1750      .frmAccountTransactions_Sub.Form.FocusHolder.SetFocus

1760      DoCmd.Hourglass False

1770    End With

EXITP:
1780    Set frm = Nothing
1790    Set rst = Nothing
1800    Set qdf = Nothing
1810    Set dbs = Nothing
1820    Exit Sub

ERRH:
1830    DoCmd.Hourglass False
1840    Select Case ERR.Number
        Case Else
1850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1860    End Select
1870    Resume EXITP

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
1980        cmdClose_Click  ' ** Procedure: Below.
1990      End Select
2000    End If

        ' ** Alt keys.
2010    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2020      Select Case intRetVal
          Case vbKeyX
2030        intRetVal = 0
2040        cmdClose_Click  ' ** Procedure: Below.
2050      End Select
2060    End If

        ' ** Ctrl-Shift keys.
2070    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2080      Select Case intRetVal
          Case vbKeyF
2090        With Me
2100          intRetVal = 0
2110          .FocusHolder.SetFocus
2120        End With
2130      End Select
2140    End If

EXITP:
2150    KeyCode = intRetVal
2160    Exit Sub

ERRH:
2170    intRetVal = 0
2180    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2190    Case Else
2200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2210    End Select
2220    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim blnRetVal As Boolean

2310    DoCmd.Hourglass True
2320    DoEvents

        ' ** Make sure all reports are closed.
2330    If Reports.Count > 0 Then
2340      Do While Reports.Count > 0
2350        DoCmd.Close acReport, Reports(0).Name
2360      Loop
2370    End If

2380    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2390    blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.

2400    gstrReturningForm = THIS_NAME
2410    gblnSetFocus = True
2420    If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
2430      Forms(strCallingForm).TimerInterval = 250&
2440      DoCmd.SelectObject acForm, strCallingForm, False
2450    Else
2460      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2470    End If

EXITP:
2480    Exit Sub

ERRH:
2490    DoCmd.Hourglass False
2500    Select Case ERR.Number
        Case Else
2510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2520    End Select
2530    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2610    With Me

2620      .TimerInterval = 0&

2630      DoCmd.Close acForm, THIS_NAME

2640    End With

EXITP:
2650    Exit Sub

ERRH:
2660    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2670    Case Else
2680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2690    End Select
2700    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab
2870        With Me
2880          intRetVal = 0
2890          lngRecsCur = .frmAccountTransactions_Sub.Form.RecCnt  ' ** Form Function: frmAccountTransactions_Sub.
2900          If lngRecsCur > 0& Then
2910            .frmAccountTransactions_Sub.SetFocus
2920            .frmAccountTransactions_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccountTransactions_Sub.
2930            .frmAccountTransactions_Sub.Form.FocusHolder.SetFocus
2940          Else
2950            If .chkIncludeArchive.Enabled = True Then
2960              .chkIncludeArchive.SetFocus
2970            ElseIf .chkIncludeCurrency.Enabled = True Then
2980              .chkIncludeCurrency.SetFocus
2990            Else
                  ' ** Nowhere else to go!
3000            End If
3010          End If
3020        End With
3030      End Select
3040    End If

        ' ** Shift keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3060      Select Case intRetVal
          Case vbKeyTab
3070        With Me
3080          intRetVal = 0
3090          If .chkIncludeCurrency.Enabled = True Then
3100            .chkIncludeCurrency.SetFocus
3110          ElseIf .chkIncludeArchive.Enabled = True Then
3120            .chkIncludeArchive.SetFocus
3130          Else
3140            lngRecsCur = .frmAccountTransactions_Sub.Form.RecCnt  ' ** Form Function: frmAccountTransactions_Sub.
3150            If lngRecsCur > 0& Then
3160              .frmAccountTransactions_Sub.SetFocus
3170              .frmAccountTransactions_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountTransactions_Sub.
3180              .frmAccountTransactions_Sub.Form.FocusHolder.SetFocus
3190            Else
                  ' ** Nowhere else to go!
3200            End If
3210          End If
3220        End With
3230      End Select
3240    End If

EXITP:
3250    KeyCode = intRetVal
3260    Exit Sub

ERRH:
3270    intRetVal = 0
3280    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
3290    Case Else
3300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3310    End Select
3320    Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

        Dim lngJournalNo As Long

3410    With Me
3420      Select Case .chkIncludeArchive
          Case True
3430        .chkIncludeArchive_lbl.FontBold = True
3440        .archive_trans_lbl.Visible = True
3450        .archive_trans_lbl_box.Visible = True
3460        lngJournalNo = .frmAccountTransactions_Sub.Form.journalno
3470        If .frmAccountTransactions_Sub.Form.RecordSource <> "qryAccountProfile_Transactions_05_01" Then
3480          .frmAccountTransactions_Sub.Form.RecordSource = "qryAccountProfile_Transactions_05_01"
3490        End If
3500        .frmAccountTransactions_Sub.Form.Requery
3510      Case False
3520        .chkIncludeArchive_lbl.FontBold = False
3530        .archive_trans_lbl.Visible = False
3540        .archive_trans_lbl_box.Visible = False
3550        lngJournalNo = .frmAccountTransactions_Sub.Form.journalno
3560        If .frmAccountTransactions_Sub.Form.RecordSource <> "qryAccountProfile_Transactions_05_02" Then
3570          .frmAccountTransactions_Sub.Form.RecordSource = "qryAccountProfile_Transactions_05_02"
3580        End If
3590        .frmAccountTransactions_Sub.Form.Requery
3600      End Select
3610      .frmAccountTransactions_Sub.Form.OrderByOn = False
3620      .frmAccountTransactions_Sub.Form.OrderByOn = True
3630      .frmAccountTransactions_Sub.SetFocus
3640      .frmAccountTransactions_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccountTransactions_Sub.
3650    End With

EXITP:
3660    Exit Sub

ERRH:
3670    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
3680    Case Else
3690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3700    End Select
3710    Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3820    intShiftDown = (Shift And acShiftMask) > 0
3830    intAltDown = (Shift And acAltMask) > 0
3840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3870        With Me
3880          intRetVal = 0
3890          If .chkIncludeCurrency.Enabled = True Then
3900            .chkIncludeCurrency.SetFocus
3910          Else
3920            .cmdClose.SetFocus
3930          End If
3940        End With
3950      End Select
3960    End If

        ' ** Shift keys.
3970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3980      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3990        With Me
4000          intRetVal = 0
4010          lngRecsCur = .frmAccountTransactions_Sub.Form.RecCnt  ' ** Form Function: frmAccountTransactions_Sub.
4020          If lngRecsCur > 0& Then
4030            .frmAccountTransactions_Sub.SetFocus
4040            .frmAccountTransactions_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountTransactions_Sub.
4050            .frmAccountTransactions_Sub.Form.FocusHolder.SetFocus
4060          Else
4070            .cmdClose.SetFocus
4080          End If
4090        End With
4100      End Select
4110    End If

EXITP:
4120    KeyCode = intRetVal
4130    Exit Sub

ERRH:
4140    intRetVal = 0
4150    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
4160    Case Else
4170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4180    End Select
4190    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

4210    With Me
4220      Select Case .chkIncludeCurrency
          Case True
4230        .chkIncludeCurrency_lbl.FontBold = True
4240        .frmAccountTransactions_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountTransactions_Sub.
4250        IncludeCurrency True  ' ** Procedure: Below.
4260        With .frmAccountTransactions_Sub.Form
4270          .ICash.Visible = False
4280          .icash_str.Visible = True
4290          .PCash.Visible = False
4300          .pcash_str.Visible = True
4310          .Cost.Visible = False
4320          .cost_str.Visible = True
4330        End With
4340      Case False
4350        .chkIncludeCurrency_lbl.FontBold = False
4360        .frmAccountTransactions_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountTransactions_Sub.
4370        IncludeCurrency False  ' ** Procedure: Below.
4380        Select Case blnHasForExThis
            Case True
              ' ** Always show the correct symbol.
4390          With .frmAccountTransactions_Sub.Form
4400            If .icash_str.Visible = False Then
4410              .ICash.Visible = False
4420              .icash_str.Visible = True
4430              .PCash.Visible = False
4440              .pcash_str.Visible = True
4450              .Cost.Visible = False
4460              .cost_str.Visible = True
4470            End If
4480          End With
4490        Case False
4500          With .frmAccountTransactions_Sub.Form
4510            .ICash.Visible = True
4520            .icash_str.Visible = False
4530            .PCash.Visible = True
4540            .pcash_str.Visible = False
4550            .Cost.Visible = True
4560            .cost_str.Visible = False
4570          End With
4580        End Select
4590      End Select
4600      .frmAccountTransactions_Sub.SetFocus
4610    End With

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4720    intShiftDown = (Shift And acShiftMask) > 0
4730    intAltDown = (Shift And acAltMask) > 0
4740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4770        With Me
4780          intRetVal = 0
4790          .cmdClose.SetFocus
4800        End With
4810      End Select
4820    End If

        ' ** Shift keys.
4830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4850        With Me
4860          intRetVal = 0
4870          If .chkIncludeArchive.Enabled = True Then
4880            .chkIncludeArchive.SetFocus
4890          Else
4900            lngRecsCur = .frmAccountTransactions_Sub.Form.RecCnt  ' ** Form Function: frmAccountTransactions_Sub.
4910            If lngRecsCur > 0& Then
4920              .frmAccountTransactions_Sub.SetFocus
4930              .frmAccountTransactions_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountTransactions_Sub.
4940              .frmAccountTransactions_Sub.Form.FocusHolder.SetFocus
4950            Else
4960              .cmdClose.SetFocus
4970            End If
4980          End If
4990        End With
5000      End Select
5010    End If

EXITP:
5020    KeyCode = intRetVal
5030    Exit Sub

ERRH:
5040    intRetVal = 0
5050    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
5060    Case Else
5070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5080    End Select
5090    Resume EXITP

End Sub

Private Sub SetArchiveOption()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "SetArchiveOption"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnFound1 As Boolean, blnFound2 As Boolean

5110    With Me

5120      blnFound1 = False: blnFound2 = False

5130      Set dbs = CurrentDb
5140      With dbs

5150        Set rst = .OpenRecordset("LedgerArchive", dbOpenDynaset, dbReadOnly)
5160        With rst
5170          If .BOF = True And .EOF = True Then
                ' ** No archive whatsoever!
5180          Else
5190            .MoveLast
5200            If .RecordCount > 1 Then
5210              blnFound1 = True
5220            End If
5230          End If
5240          .Close
5250        End With
5260        Set rst = Nothing

5270        If blnFound1 = True Then
              ' ** LedgerArchive, with 'Liability (+)'/'Liability (-)' broken out,
              ' ** with totdescx, by specified GlobalVarGet('gstrAccountNo').
5280          Set qdf = .QueryDefs("qryAccountProfile_Transactions_02_01")
5290          Set rst = qdf.OpenRecordset
5300          With rst
5310            If .BOF = True And .EOF = True Then
                  ' ** This account has no archive.
5320            Else
5330              .MoveLast
5340              If .RecordCount > 0 Then
5350                blnFound2 = True
5360              End If
5370            End If
5380            .Close
5390          End With
5400          Set rst = Nothing
5410          Set qdf = Nothing
5420        End If  ' ** blnFound1.

5430        .Close
5440      End With
5450      Set dbs = Nothing

5460      Select Case blnFound1
          Case True
5470        Select Case blnFound2
            Case True
5480          .chkIncludeArchive.Enabled = True
5490          If .chkIncludeArchive_lbl.Caption <> "Include Arc&hive" Then
5500            .chkIncludeArchive_lbl.Caption = "Include Arc&hive"
5510          End If
5520          .chkIncludeArchive_lbl4_dim.Visible = False
5530          .chkIncludeArchive_lbl4_dim_hi.Visible = False
5540          .chkIncludeArchive_lbl2_dim.Visible = False
5550          .chkIncludeArchive_lbl2_dim_hi.Visible = False
5560        Case False
5570          .chkIncludeArchive.Enabled = False
5580          .chkIncludeArchive_lbl.Visible = False
5590          .chkIncludeArchive_lbl4_dim.Visible = True
5600          .chkIncludeArchive_lbl4_dim_hi.Visible = True
5610          .chkIncludeArchive_lbl2_dim.Visible = False
5620          .chkIncludeArchive_lbl2_dim_hi.Visible = False
5630        End Select
5640      Case False
5650        .chkIncludeArchive.Enabled = False
5660        .chkIncludeArchive_lbl.Visible = False
5670        .chkIncludeArchive_lbl2_dim.Visible = True
5680        .chkIncludeArchive_lbl2_dim_hi.Visible = True
5690        .chkIncludeArchive_lbl4_dim.Visible = False
5700        .chkIncludeArchive_lbl4_dim_hi.Visible = False
5710      End Select

5720    End With

EXITP:
5730    Set rst = Nothing
5740    Set qdf = Nothing
5750    Set dbs = Nothing
5760    Exit Sub

ERRH:
5770    Select Case ERR.Number
        Case Else
5780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5790    End Select
5800    Resume EXITP

End Sub

Private Sub IncludeCurrency(blnShow As Boolean)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long

5910    With Me

5920      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5930        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5940      End If

5950      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
5960      lngMonitorNum = 1&: lngTmp01 = 0&
5970      EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
5980      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

5990      Select Case blnShow
          Case True
6000        DoCmd.SelectObject acForm, THIS_NAME, False
6010        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
6020        DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, lngFrm_Height  'lngFrm_Top
6030        If lngMonitorNum > 1& Then
6040          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
6050        End If
6060        .Width = lngForm_Width
6070        .cmdClose.Left = lngClose_Left
6080        .frmAccountTransactions_Sub.Width = lngSub_Width
6090        .Detail_hline01.Width = lngDetailHLine01_Width
6100        .Detail_hline02.Width = lngDetailHLine02_Width
6110        .Detail_hline03.Width = lngDetailHLine02_Width
6120        .Detail_hline04.Width = lngDetailHLine01_Width
6130        .Detail_vline03.Left = lngDetailVLine03_Left
6140        .Detail_vline04.Left = (lngDetailVLine03_Left + lngTpp)
6150        .Nav_box01.Width = .frmAccountTransactions_Sub.Width
6160        .Header_vline01.Left = lngForm_Width
6170        .Header_vline02.Left = lngForm_Width
6180        .Footer_vline01.Left = lngForm_Width
6190        .Footer_vline02.Left = lngForm_Width
6200        .Header_hline01.Width = lngForm_Width
6210        .Header_hline02.Width = lngForm_Width
6220        .Footer_hline01.Width = lngForm_Width
6230        .Footer_hline02.Width = lngForm_Width
6240      Case False
6250        .cmdClose.Left = (lngClose_Left - lngTotWidth_Diff)
6260        .frmAccountTransactions_Sub.Width = (lngSub_Width - lngTotWidth_Diff)
6270        .Nav_box01.Width = .frmAccountTransactions_Sub.Width
6280        .Detail_hline01.Width = (lngDetailHLine01_Width - lngTotWidth_Diff)
6290        .Detail_hline02.Width = (lngDetailHLine02_Width - lngTotWidth_Diff)
6300        .Detail_hline03.Width = (lngDetailHLine02_Width - lngTotWidth_Diff)
6310        .Detail_hline04.Width = (lngDetailHLine01_Width - lngTotWidth_Diff)
6320        .Detail_vline03.Left = (lngDetailVLine03_Left - lngTotWidth_Diff)
6330        .Detail_vline04.Left = (.Detail_vline03.Left + lngTpp)
6340        .Header_vline01.Left = (lngForm_Width - lngTotWidth_Diff)
6350        .Header_vline02.Left = (lngForm_Width - lngTotWidth_Diff)
6360        .Footer_vline01.Left = (lngForm_Width - lngTotWidth_Diff)
6370        .Footer_vline02.Left = (lngForm_Width - lngTotWidth_Diff)
6380        .Header_hline01.Width = (lngForm_Width - lngTotWidth_Diff)
6390        .Header_hline02.Width = (lngForm_Width - lngTotWidth_Diff)
6400        .Footer_hline01.Width = (lngForm_Width - lngTotWidth_Diff)
6410        .Footer_hline02.Width = (lngForm_Width - lngTotWidth_Diff)
6420        .Width = (lngForm_Width - lngTotWidth_Diff)
6430        DoCmd.SelectObject acForm, THIS_NAME, False
6440        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
6450        DoCmd.MoveSize (lngFrm_Left + (lngTotWidth_Diff / 2)), lngTmp01, (lngFrm_Width - lngTotWidth_Diff), lngFrm_Height  'lngFrm_Top
6460        If lngMonitorNum > 1& Then
6470          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
6480        End If
6490      End Select
6500      DoEvents

6510    End With

EXITP:
6520    Exit Sub

ERRH:
6530    Select Case ERR.Number
        Case Else
6540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6550    End Select
6560    Resume EXITP

End Sub
