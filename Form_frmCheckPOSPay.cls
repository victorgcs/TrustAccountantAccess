VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckPOSPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckPOSPay"

'VGC 06/21/2017: CHANGES!

'SHOULD THIS FILTER BY USER?

' ** Shortcut Alt keys responsive from this form:
' **   Issue Date:       A {Issue_Date}
' **   Check Number:     B {ppd_checknum on frmCheckPOSPay_Sub3}
' **   Close:            C {cmdClose}
' **   Delete:           D {cmdDelete}
' **   Edit:             E {cmdEdit}
' **   Files:            F {opgView_optFiles}
' **   Generate POSPay:  G {cmdGeneratePOSPay}
' **   Checks:           H {opgView_optChecks}
' **   Edit POSPay:      I {cmdEditPosPay}
' **   Check Number:     K {CheckNum on frmCheckPOSPay_Sub1}
' **   Include Currency: L {chkIncludeCurrency}
' **   Check Memo:       M {JrnlMemo_Memo on frmCheckPOSPay_Sub1}
' **   Regenerate:       N {cmdRegenerate}
' **   Comments:         O {description on frmCheckPOSPay_Sub1}
' **   Print Report:     P {cmdPrintReport}
' **   Currency:         R {curr_id on frmCheckPOSPay_Sub1}
' **   Posting Date:     S {transdate on frmCheckPOSPay_Sub1}
' **   Delete POSPay:    T {cmdDeletePosPay}
' **   Update:           U {cmdUpdate}
' **   Void:             V {ppd_void on frmCheckPOSPay_Sub3}
' **   Exit:             X {cmdClose}
' **   Payee:            Y {RecurringItem on frmCheckPOSPay_Sub1}
'J  Q  W  Z

'I CAN'T FIGURE OUT WHAT TO GIVE THESE!
' **   Description:      P {pp_description on frmCheckPOSPay_Sub2}  ' ** Not visible at
' **   Payee:            P {ppd_payee on frmCheckPOSPay_Sub3}       ' ** the same time.

' ** Shortcut Ctrl keys responsive from this form:
' **   Next:             N {MoveRec on frmCheckPOSPay_Sub2}
' **   Previous:         P {MoveRec on frmCheckPOSPay_Sub2}
' **   Save:             S {cmdSave on frmCheckPOSPay_Sub1}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private lngRecsCur As Long
Private blnHasChecks As Boolean, lngCurChecks As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           DoCmd.Hourglass True  ' ** Make sure it's still running.
150           DoEvents

160           strCallingForm = Trim(.OpenArgs)

170           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

              ' ** Reload currency stuff.
180           gblnHasForEx = HasForEx_All  ' ** Module Function: modCurrencyFuncs.
190           DoEvents
200           If gblnHasForEx = True Then
210             HasForEx_Load  ' ** Module Procedure: modCurrencyFuncs.
220             DoEvents
230           End If

240           DoCmd.Hourglass True  ' ** Make sure it's still running.
250           DoEvents

260           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
270           DoEvents

280           Select Case gblnHasForEx
              Case True
290             .chkIncludeCurrency.Enabled = True
                ' ** Make sure this happens after the array is loaded.
300             chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
310           Case False
320             .chkIncludeCurrency.Enabled = False
330           End Select

340           DoCmd.Hourglass True  ' ** Make sure it's still running.
350           DoEvents

360           opgView_AfterUpdate  ' ** Procedure: Below.

370           DoCmd.Hourglass False

380         Else
390           Cancel = -1
400         End If
410       Else
420         Cancel = -1
430       End If
440     End With

450     If Cancel = -1 Then
460       Beep
470       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
480     End If

EXITP:
490     Exit Sub

ERRH:
500     DoCmd.Hourglass False
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

610     With Me

620       CLR_DISABLED_FG = CLR_DKGRY
630       CLR_DISABLED_BG = CLR_LTTEAL

640       blnHasChecks = False

650       Set dbs = CurrentDb

660       DoCmd.Hourglass True  ' ** Make sure it's still running.
670       DoEvents

          '         ***************
          'THESE DON'T NEED TO BE DONE EVERY TIME!

          ' ** Update qryCheckPOSPay_13_01 (RecurringItems,
          ' ** just leading spaces, with RecurringItem_new).
680       Set qdf = dbs.QueryDefs("qryCheckPOSPay_13_02")
690       qdf.Execute
700       Set qdf = Nothing
710       DoEvents

          ' ** Ohana specific.
          ' ** Update qryCheckPOSPay_13_03 (RecurringItems,
          ' ** just missing spaces, with RecurringItem_new).
          'Set qdf = dbs.QueryDefs("qryCheckPOSPay_13_04")
          'qdf.Execute
          'Set qdf = Nothing
          'DoEvents

          ' ** Update qryCheckPOSPay_15_01 (Ledger, just
          ' ** leading spaces, with RecurringItem_new).
720       Set qdf = dbs.QueryDefs("qryCheckPOSPay_15_02")
730       qdf.Execute
740       Set qdf = Nothing
750       DoEvents

          ' ** Update qryCheckPOSPay_15_03 (LedgerArchive, just
          ' ** leading spaces, with RecurringItem_new).
760       Set qdf = dbs.QueryDefs("qryCheckPOSPay_15_04")
770       qdf.Execute
780       Set qdf = Nothing
790       DoEvents
          '         ***************

800       Sub1_Load  ' ** Procedure: Below

810       DoCmd.Hourglass True  ' ** Make sure it's still running.
820       DoEvents

830       lngRecsCur = .frmCheckPOSPay_Sub2.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub2.
840       If lngRecsCur = 0& Then
850         .cmdRegenerate.Enabled = False
860         .cmdEditPosPay.Enabled = False
870         .cmdDeletePosPay.Enabled = False
880         .pp_date.BackStyle = acBackStyleTransparent
890         .pp_date.BorderColor = WIN_CLR_DISR
900         .pp_description.BackStyle = acBackStyleTransparent
910         .pp_description.BorderColor = WIN_CLR_DISR
920         .frmCheckPOSPay_Sub2.Form.Sub2_Disable False  ' ** Form Procedure: frmCheckPOSPay_Sub2.
930         .FocusHolder.SetFocus
940         .frmCheckPOSPay_Sub2.Enabled = False
950         .frmCheckPOSPay_Sub3.Form.Sub3_Disable False  ' ** Form Procedure: frmCheckPOSPay_Sub3.
960         .FocusHolder.SetFocus
970         .frmCheckPOSPay_Sub3.Enabled = False
980         .frmCheckPOSPay_Sub1.Visible = False
990         .frmCheckPOSPay_Sub1_lbl_dim.Visible = True
1000        .frmCheckPOSPay_Sub1_lbl_dim_hi.Visible = True
1010      End If

100       lngRecsCur = .frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
110       If lngRecsCur = 0& Then
120         .cmdEdit.Enabled = False
130         .cmdDelete.Enabled = False
140         .cmdUpdate.Enabled = False
150       End If

1020      DoCmd.Hourglass False

1030    End With

EXITP:
1040    Set rst = Nothing
1050    Set qdf = Nothing
1060    Set dbs = Nothing
1070    Exit Sub

ERRH:
1080    DoCmd.Hourglass False
1090    Select Case ERR.Number
        Case Else
1100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1110    End Select
1120    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Check Number:     B {ppd_checknum on frmCheckPOSPay_Sub3}
        ' **   Check Number:     K {CheckNum on frmCheckPOSPay_Sub1}
        ' **   Check Memo:       M {JrnlMemo_Memo on frmCheckPOSPay_Sub1}
        ' **   Comments:         O {description on frmCheckPOSPay_Sub1}
        ' **   Currency:         R {curr_id on frmCheckPOSPay_Sub1}
        ' **   Posting Date:     S {transdate on frmCheckPOSPay_Sub1}
        ' **   Exit:             X {cmdClose}
        ' **   Payee:            Y {RecurringItem on frmCheckPOSPay_Sub1}

        'I CAN'T FIGURE OUT WHAT TO GIVE THESE!
        ' **   Description:      P {pp_description on frmCheckPOSPay_Sub2}
        ' **   Payee:            P {ppd_payee on frmCheckPOSPay_Sub3}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next:             N {MoveRec on frmCheckPOSPay_Sub2}
        ' **   Previous:         P {MoveRec on frmCheckPOSPay_Sub2}
        ' **   Save:             S {cmdSave on frmCheckPOSPay_Sub1}

        ' ** Alt keys.
1150    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyB
1170        With Me
1180          intRetVal = 0
1190          If .frmCheckPOSPay_Sub3.Visible = True Then
1200            .frmCheckPOSPay_Sub3.SetFocus
1210            .frmCheckPOSPay_Sub3.Form.ppd_checknum.SetFocus
1220          Else
1230            Beep
1240          End If
1250        End With
1260      Case vbKeyK
1270        With Me
1280          intRetVal = 0
1290          .frmCheckPOSPay_Sub1.SetFocus
1300          .frmCheckPOSPay_Sub1.Form.CheckNum.SetFocus
1310        End With
1320      Case vbKeyM
1330        With Me
1340          intRetVal = 0
1350          .frmCheckPOSPay_Sub1.SetFocus
1360          .frmCheckPOSPay_Sub1.Form.JrnlMemo_Memo.SetFocus
1370        End With
1380      Case vbKeyO
1390        With Me
1400          intRetVal = 0
1410          .frmCheckPOSPay_Sub1.SetFocus
1420          .frmCheckPOSPay_Sub1.Form.description.SetFocus
1430        End With
'1440      Case vbKeyP
'1450        With Me
'1460          intRetVal = 0
'1470          If .frmCheckPOSPay_Sub2.Visible = True Then
'1480            .frmCheckPOSPay_Sub2.SetFocus
'1490            .frmCheckPOSPay_Sub2.Form.pp_description.SetFocus
'1500          ElseIf .frmCheckPOSPay_Sub3.Visible = True Then
'1510            .frmCheckPOSPay_Sub3.SetFocus
'1520            .frmCheckPOSPay_Sub3.Form.ppd_payee.SetFocus
'1530          End If
'1540        End With
1440      Case vbKeyR
1450        With Me
1460          intRetVal = 0
1470          .frmCheckPOSPay_Sub1.SetFocus
1480          .frmCheckPOSPay_Sub1.Form.curr_id.SetFocus
1490        End With
1500      Case vbKeyS
1510        With Me
1520          intRetVal = 0
1530          .frmCheckPOSPay_Sub1.SetFocus
1540          .frmCheckPOSPay_Sub1.Form.transdate.SetFocus
1550        End With
1560      Case vbKeyX
1570        intRetVal = 0
1580        cmdClose_Click  ' ** Procedure: Below.
1590      Case vbKeyY
1600        With Me
1610          intRetVal = 0
1620          .frmCheckPOSPay_Sub1.SetFocus
1630          .frmCheckPOSPay_Sub1.Form.RecurringItem.SetFocus
1640        End With
1650      End Select
1660    End If

        ' ** Ctrl keys.
1670    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1680      Select Case intRetVal
          Case vbKeyN
1690        With Me
1700          intRetVal = 0
1710          .frmCheckPOSPay_Sub2.Form.MoveRecX acCmdRecordsGoToNext  ' ** Form Procedure: frmCheckPOSPay_Sub2.
1720        End With
1730      Case vbKeyP
1740        With Me
1750          intRetVal = 0
1760          .frmCheckPOSPay_Sub2.Form.MoveRecX acCmdRecordsGoToPrevious  ' ** Form Procedure: frmCheckPOSPay_Sub2.
1770        End With
1780      Case vbKeyS
1790        With Me
1800          intRetVal = 0
              '.frmCheckPOSPay_Sub1.SetFocus
              '.frmCheckPOSPay_Sub1.Form.cmdSave_Click  ' ** Form Procedure: frmCheckPOSPay_Sub1.
1810        End With
1820      End Select
1830    End If

        ' ** Ctrl-Shift keys.
1840    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1850      Select Case intRetVal
          Case vbKeyF
1860        With Me
1870          intRetVal = 0
1880          .FocusHolder.SetFocus
1890        End With
1900      End Select
1910    End If

EXITP:
1920    KeyCode = intRetVal
1930    Exit Sub

ERRH:
1940    intRetVal = 0
1950    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
1960    Case Else
1970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1980    End Select
1990    Resume EXITP

End Sub

Private Sub Form_Timer()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2010    With Me
2020      .TimerInterval = 0&
2030      If gblnSetFocus = True Then
2040        gblnSetFocus = False
2050        strFormName = vbNullString
2060  On Error Resume Next
2070        Do Until strFormName = THIS_NAME
2080          DoCmd.SelectObject acForm, THIS_NAME, False
2090          strFormName = Screen.ActiveForm.Name
2100        Loop
2110  On Error GoTo ERRH
2120      End If
2130    End With

EXITP:
2140    Exit Sub

ERRH:
2150    DoCmd.Hourglass False
2160    Select Case ERR.Number
        Case Else
2170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2180    End Select
2190    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim lngRecsChanged As Long
        Dim msgResponse As VbMsgBoxResult

2210    With Me

2220      lngRecsChanged = HasChanged  ' ** Function: Below.

2230      If lngRecsChanged > 0& Then
2240        DoCmd.Hourglass False
2250        Beep
2260        msgResponse = MsgBox("You have made changes to the Journal, but have not" & vbCrLf & _
              "confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
              "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved")
2270      Else
2280        msgResponse = vbYes
2290      End If

2300      Select Case msgResponse
          Case vbYes

2310        DoCmd.Hourglass True
2320        DoEvents

2330        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2340        gstrReturningForm = THIS_NAME
2350        gblnSetFocus = True
2360        If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
2370          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2380        Else
2390          DoCmd.SelectObject acForm, strCallingForm, False
2400        End If

2410      Case Else
2420        Cancel = -1
2430      End Select

2440    End With

EXITP:
2450    Exit Sub

ERRH:
2460    DoCmd.Hourglass False
2470    Select Case ERR.Number
        Case Else
2480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2490    End Select
2500    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2610    With Me
2620      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2630        Select Case blnPrintReport_Focus
            Case True
2640          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2650          .cmdPrintReport_raised_img.Visible = False
2660        Case False
2670          .cmdPrintReport_raised_img.Visible = True
2680          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2690        End Select
2700        .cmdPrintReport_raised_focus_img.Visible = False
2710        .cmdPrintReport_raised_focus_dots_img.Visible = False
2720        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2730        .cmdPrintReport_raised_img_dis.Visible = False
2740      End If
2750    End With

EXITP:
2760    Exit Sub

ERRH:
2770    Select Case ERR.Number
        Case Else
2780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2790    End Select
2800    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2920    Exit Sub

ERRH:
2930    Select Case ERR.Number
        Case Else
2940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950    End Select
2960    Resume EXITP

End Sub

Public Sub cmdClose_Click()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3010    DoCmd.Hourglass True
3020    DoEvents

3030    DoCmd.Close acForm, THIS_NAME

EXITP:
3040    Exit Sub

ERRH:
3050    Select Case ERR.Number
        Case Else
3060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3070    End Select
3080    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
3170        With Me
3180          intRetVal = 0
3190          If .frmCheckPOSPay_Sub1.Enabled = True Then
3200            .frmCheckPOSPay_Sub1.SetFocus
3210          ElseIf .frmCheckPOSPay_Sub2.Visible = True And .frmCheckPOSPay_Sub2.Enabled = True Then
3220            .frmCheckPOSPay_Sub2.SetFocus
3230          ElseIf .frmCheckPOSPay_Sub3.Visible = True And .frmCheckPOSPay_Sub3.Enabled = True Then
3240            .frmCheckPOSPay_Sub3.SetFocus
3250          Else
3260            .FocusHolder.SetFocus
3270          End If
3280        End With
3290      End Select
3300    End If

        ' ** Shift keys.
3310    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3320      Select Case intRetVal
          Case vbKeyTab
3330        With Me
3340          intRetVal = 0
3350          If .cmdPrintReport.Enabled = True Then
3360            .cmdPrintReport.SetFocus
3370          ElseIf .cmdDeletePosPay.Enabled = True Then
3380            .cmdDeletePosPay.SetFocus
3390          ElseIf .cmdEditPosPay.Enabled = True Then
3400            .cmdEditPosPay.SetFocus
3410          ElseIf .cmdRegenerate.Enabled = True Then
3420            .cmdRegenerate.SetFocus
3430          Else
3440            .opgView.SetFocus
3450          End If
3460        End With
3470      End Select
3480    End If

EXITP:
3490    KeyCode = intRetVal
3500    Exit Sub

ERRH:
3510    intRetVal = 0
3520    Select Case ERR.Number
        Case Else
3530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3540    End Select
3550    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

3610    With Me
3620      Select Case chkIncludeCurrency
          Case True
3630        .chkIncludeCurrency_lbl.FontBold = True
3640        .frmCheckPOSPay_Sub1.Form.IncludeCurrency True  ' ** Form Procedure: frmCheckPOSPay_Sub1.
3650      Case False
3660        .chkIncludeCurrency_lbl.FontBold = False
3670        .frmCheckPOSPay_Sub1.Form.IncludeCurrency False  ' ** Form Procedure: frmCheckPOSPay_Sub1.
3680      End Select
3690    End With

EXITP:
3700    Exit Sub

ERRH:
3710    Select Case ERR.Number
        Case Else
3720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3730    End Select
3740    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

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
          Case vbKeyTab
3870        With Me
3880          intRetVal = 0
3890          If .cmdGeneratePOSPay.Enabled = True Then
3900            .cmdGeneratePOSPay.SetFocus
3910          ElseIf .cmdEdit.Enabled = True Then
3920            .cmdEdit.SetFocus
3930          ElseIf .cmdDelete.Enabled = True Then
3940            .cmdDelete.SetFocus
3950          ElseIf .cmdUpdate.Enabled = True Then
3960            .cmdUpdate.SetFocus
3970          Else
3980            .opgView.SetFocus
3990          End If
4000        End With
4010      End Select
4020    End If

        ' ** Shift keys.
4030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4040      Select Case intRetVal
          Case vbKeyTab
4050        With Me
4060          intRetVal = 0
4070          If .frmCheckPOSPay_Sub1.Enabled = True Then
4080            .frmCheckPOSPay_Sub1.SetFocus
4090          Else
4100            .cmdClose.SetFocus
4110          End If
4120        End With
4130      End Select
4140    End If

EXITP:
4150    KeyCode = intRetVal
4160    Exit Sub

ERRH:
4170    intRetVal = 0
4180    Select Case ERR.Number
        Case Else
4190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4200    End Select
4210    Resume EXITP

End Sub

Private Sub cmdGeneratePOSPay_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdGeneratePOSPay_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strPath As String, strFile As String, strPathFile As String, strDesc As String
        Dim strDocName As String, strSortNow As String
        Dim lngRecs As Long, lngRecsChanged As Long, lngPPID As Long, lngPPDID As Long
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, lngTmp03 As Long, lngTmp04 As Long, dblTmp05 As Double
        Dim lngX As Long

4310    With Me

4320      DoCmd.Hourglass True
4330      DoEvents

4340      lngRecsCur = .frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
4350      If lngRecsCur > 0& Then

4360        lngTmp03 = .frmCheckPOSPay_Sub1.Form.Recalc_NoRNAN  ' ** Form Function: frmCheckPOSPay_Sub1.
4370        If lngTmp03 = lngRecsCur Then
              ' ** None can be exported.
4380          DoCmd.Hourglass False
4390          Beep
4400          MsgBox "There are no valid checks. No file can be generated.", vbInformation + vbOKOnly, "Nothing To Do"
4410        Else
4420          If lngTmp03 > 0 Then
                ' ** Let them know.
4430            Beep
4440            DoCmd.Hourglass False
4450            msgResponse = MsgBox(CStr(lngTmp03) & " of your checks are missing bank information " & _
                  "and will not be included." & vbCrLf & vbCrLf & "Do you wish to continue?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "List Has Invalid Checks")
4460          Else
4470            msgResponse = vbYes
4480          End If

4490          If msgResponse = vbYes Then

4500            DoCmd.Hourglass True
4510            DoEvents

4520            msgResponse = vbCancel

4530            lngTmp04 = .frmCheckPOSPay_Sub1.Form.Recalc_CheckNum_NoNum  ' ** Form Function: frmCheckPOSPay_Sub1.
4540            If lngTmp04 = lngRecsCur Then
                  ' ** None can be exported.
4550              DoCmd.Hourglass False
4560              Beep
4570              MsgBox "There are no valid checks. No file can be generated.", vbInformation + vbOKOnly, "Nothing To Do"
4580            Else
4590              If lngTmp04 > 0 Then
                    ' ** Let them know.
4600                Beep
4610                DoCmd.Hourglass False
4620                msgResponse = MsgBox(CStr(lngTmp04) & " of your checks are missing a check number " & _
                      "and will not be included." & vbCrLf & vbCrLf & "Do you wish to continue?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "List Has Invalid Checks")
4630              Else
4640                msgResponse = vbYes
4650              End If
4660            End If

4670          End If  ' ** msgResponse.

4680          If msgResponse = vbYes Then

4690            DoCmd.Hourglass True
4700            DoEvents

4710            lngRecsChanged = HasChanged  ' ** Function: Below.
4720            If lngRecsChanged = 0& Then

4730              DoCmd.Hourglass False
4740              msgResponse = MsgBox("This will generate a POSPay file for the checks listed above." & vbCrLf & vbCrLf & _
                    "Proceed?", vbQuestion + vbOKCancel, "Generate A POSPay File")

4750              If msgResponse = vbOK Then

4760                DoCmd.Hourglass True
4770                DoEvents

4780                Set dbs = CurrentDb

                    ' ** tblCheckPOSPay_Journal, as POSPay file format, in Courier font.
4790                Set qdf = dbs.QueryDefs("qryCheckPOSPay_21_01")
4800                Set rst1 = qdf.OpenRecordset
4810                If rst1.BOF = True And rst1.EOF = True Then
                      ' ** Something bad happend!
4820                  lngCurChecks = 0&
4830                  rst1.Close
4840                  Set rst1 = Nothing
4850                  Set qdf = Nothing
4860                  DoCmd.Hourglass False
4870                  Beep
4880                  MsgBox "A problem as occurred and this process cannot continue.", vbExclamation + vbOKOnly, "Unknown Error"
4890                Else
4900                  rst1.MoveLast
4910                  lngCurChecks = rst1.RecordCount
4920                  rst1.Close
4930                  Set rst1 = Nothing
4940                  Set qdf = Nothing
4950                End If

4960                If lngCurChecks > 0 Then

4970                  If IsNull(.UserReportPath) = True Then
4980                    strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
4990                  Else
5000                    strPath = .UserReportPath
5010                  End If
5020                  gdatEndDate = .Issue_Date
5030                  strFile = "POSPay" & Format(gdatEndDate, "mmddyy") & ".txt"

5040                  DoCmd.Hourglass False
5050                  strPathFile = FileSaveDialog("txt", strFile, strPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

5060                  If strPathFile <> vbNullString Then

5070                    DoCmd.Hourglass True
5080                    DoEvents

5090                    strPath = Parse_Path(strPathFile)  ' ** Module Function: modFileUtilities.
5100                    strFile = Parse_File(strPathFile)  ' ** Module Function: modFileUtilities.
5110                    .UserReportPath = strPath

5120                    strDesc = CStr(lngCurChecks) & " Check" & IIf(lngCurChecks = 1, vbNullString, "s")
5130                    strDesc = strDesc & " issued " & Format(gdatEndDate, "mmmm d, yyyy")
5140                    strDesc = strDesc & ", drawn on "
                        ' ** qryCheckPOSPay_21_01 (tblCheckPOSPay_Journal, PrintCheck = True, as
                        ' ** POSPay file format, in Courier font), grouped by ABA_TRC, with cnt_chk.
5150                    varTmp00 = DCount("*", "qryCheckPOSPay_22_01")
5160                    If varTmp00 = 1 Then
                          ' ** qryCheckPOSPay_21_01 (tblCheckPOSPay_Journal, PrintCheck = True, as
                          ' ** POSPay file format, in Courier font), grouped by ABA_TRC, with cnt_chk.
5170                      varTmp00 = DLookup("[Bank_Name]", "qryCheckPOSPay_22_01")
5180                      strDesc = strDesc & varTmp00
5190                    Else
5200                      strDesc = strDesc & "various banks"
5210                    End If
5220                    strDesc = strDesc & ", from "
                        ' ** qryCheckPOSPay_21_01 (tblCheckPOSPay_Journal, PrintCheck = True,
                        ' ** as POSPay file format, in Courier font), linked to Account,
                        ' ** grouped, by ABA_TRC, Bank_AccountNumbert, with cnt_chk.
5230                    varTmp00 = DCount("*", "qryCheckPOSPay_22_02")
5240                    If varTmp00 = 1 Then
5250                      varTmp00 = DLookup("[Bank_AccountNumber]", "qryCheckPOSPay_22_02")
5260                      strDesc = strDesc & "account " & varTmp00 & "."
5270                    Else
5280                      strDesc = strDesc & "various accounts."
5290                    End If
5300                    DoEvents

5310                    Set rst1 = dbs.OpenRecordset("tblCheckPOSPay", dbOpenDynaset, dbConsistent)
5320                    With rst1
5330                      .AddNew
                          ' ** ![pp_id] : AutoNumber.
5340                      ![pp_date] = gdatEndDate
5350                      ![pp_description] = strDesc
5360                      ![pp_pathfile] = strPathFile
5370                      ![pp_checks] = lngCurChecks
5380                      ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
5390                      ![pp_user] = GetUserName  ' ** Module Function: modFileUtilities.
5400                      ![pp_datemodified] = Now()
5410                      .Update
5420                      .Bookmark = .LastModified
5430                      lngPPID = ![pp_id]
5440                      .Close
5450                    End With  ' ** rst1.
5460                    Set rst1 = Nothing
5470                    DoEvents

                        ' ** tblCheckPOSPay_Journal, as new tblCheckPOSPay_Detail
                        ' ** records, by specified [ppid], GlobalVarGet('gdatEndDate').
5480                    Set qdf = dbs.QueryDefs("qryCheckPOSPay_23_01")
5490                    With qdf.Parameters
5500                      ![ppid] = lngPPID
5510                    End With
5520                    Set rst1 = qdf.OpenRecordset
5530                    Set rst2 = dbs.OpenRecordset("tblCheckPOSPay_Detail", dbOpenDynaset, dbConsistent)
5540                    With rst1
5550                      .MoveLast
5560                      lngRecs = .RecordCount
5570                      .MoveFirst
5580                      For lngX = 1& To lngRecs
5590                        With rst2
5600                          .AddNew
5610                          ![pp_id] = rst1![pp_id]
                              ' ** ![ppd_id] : AutoNumber.
5620                          ![Journal_ID] = rst1![Journal_ID]
5630                          ![journalno] = rst1![journalno]
5640                          ![ppd_checknum] = rst1![ppd_checknum]
5650                          ![ppd_issuedate] = rst1![ppd_issuedate]
5660                          ![accountno] = rst1![accountno]
5670                          ![ppd_amount] = rst1![ppd_amount]
5680                          ![ppd_payee] = rst1![ppd_payee]
5690                          ![RecurringItem_ID] = rst1![RecurringItem_ID]
5700                          ![ppd_bank_name] = rst1![ppd_bank_name]
5710                          ![ppd_aba_trc] = rst1![ppd_aba_trc]
5720                          ![ppd_bank_account] = rst1![ppd_bank_account]
5730                          ![curr_id] = rst1![curr_id]
5740                          ![ppd_void] = False
5750                          ![Username] = rst1![Username]
5760                          ![ppd_user] = rst1![ppd_user]
5770                          ![ppd_datemodified] = Now()
5780                          .Update
5790                          .Bookmark = .LastModified
5800                          lngPPDID = ![ppd_id]
5810                        End With  ' ** rst2.
                            ' ** Using a base-255 number, so 2-char text can store ppd_id. Max 65025.
5820                        If lngPPDID <= 255& Then
5830                          strTmp01 = Chr(lngPPDID)
5840                          strTmp02 = vbNullString
5850                        Else
5860                          dblTmp05 = (lngPPDID / 255)
5870                          lngTmp03 = Int(dblTmp05)  ' ** Don't use CInt()!
5880                          strTmp01 = Chr(lngTmp03)
5890                          lngTmp03 = (lngPPDID Mod 255)
5900                          strTmp02 = Chr(lngTmp03)
5910                        End If
5920                        .Edit
5930                        ![assettype] = strTmp01 & strTmp02
5940                        .Update
5950                        If lngX < lngRecs Then .MoveNext
5960                      Next  ' ** lngX.
5970                      .Close
5980                    End With  ' ** rst1.
5990                    rst2.Close
6000                    Set rst1 = Nothing
6010                    Set rst2 = Nothing
6020                    DoEvents

                        ' ** Update qryCheckPOSPay_16_02 (Journal, linked to qryCheckPOSPay_01_01
                        ' ** (tblCheckPOSPay_Journal, with assettype, pp_id.), with assettype_new).
6030                    Set qdf = dbs.QueryDefs("qryCheckPOSPay_16_03")
6040                    qdf.Execute
6050                    Set qdf = Nothing
6060                    DoEvents

                        ' ** Update tblCheckPOSPay_Journal, via subquery to qryCheckPOSPay_01_02
                        ' ** (qryCheckPOSPay_01_01 (tblCheckPOSPay_Journal, with assettype, pp_id), linked to
                        ' ** tblCheckPOSPay_Detail, for Journal entries already generated), for ppj_pospay = True.
6070                    Set qdf = dbs.QueryDefs("qryCheckPOSPay_01_03")
6080                    qdf.Execute
6090                    Set qdf = Nothing
6100                    DoEvents

                        ' ** To decode assettype:
                        ' ** If Len([assettype]) = 1 Then
                        ' **   ppd_id = Asc([assettype])
                        ' ** Else
                        ' **   ppd_id = (Asc(Left([assettype],1)) * 255) + Asc(Right([assettype],1))
                        ' ** End If

                        ' ** Decided to do it manually so I can grab the ppd_id as it's being created.
                        ' ** Append qryCheckPOSPay_23_01 (tblCheckPOSPay_Journal, as
                        ' ** new tblCheckPOSPay_Detail records, by specified [ppid],
                        ' ** GlobalVarGet('gdatEndDate')) to tblCheckPOSPay_Detail.
                        'Set qdf = dbs.QueryDefs("qryCheckPOSPay_23_02")
                        'With qdf.Parameters
                        '  ![ppid] = lngPPID
                        'End With
                        'qdf.Execute dbFailOnError
                        'Set qdf = Nothing
                        'DoEvents

                        ' ** Generate the POSPay file using the export spec I created.
6110                    DoCmd.TransferText acExportFixed, "Umpqua_Bank_Export_Specification", "qryCheckPOSPay_21_01", strPathFile, False
6120                    DoEvents

6130                    If .frmCheckPOSPay_Sub2.Form.pp_id_lbl_dim_hi.Visible = True Then
6140                      .frmCheckPOSPay_Sub2.Enabled = True
6150                      .frmCheckPOSPay_Sub2.Form.Sub2_Disable True  ' ** Form Procedure: frmCheckPOSPay_Sub2.
6160                      .frmCheckPOSPay_Sub3.Enabled = True
6170                      .frmCheckPOSPay_Sub3.Form.Sub3_Disable True  ' ** Form Procedure: frmCheckPOSPay_Sub3.
6180                      .pp_date.BackStyle = acBackStyleNormal
6190                      .pp_date.BorderColor = CLR_LTBLU2
6200                      .pp_description.BackStyle = acBackStyleNormal
6210                      .pp_description.BorderColor = CLR_LTBLU2
6220                    End If

6230                    .frmCheckPOSPay_Sub2.Form.Requery
6240                    .frmCheckPOSPay_Sub2.Form.MoveRec 0, lngPPID  ' ** Form Procedure: frmCheckPOSPay_Sub2.
6250                    DoEvents

6260                    .frmCheckPOSPay_Sub1.Form.Requery

6270                    If .cmdDeletePosPay.Enabled = False Then
6280                      .cmdDeletePosPay.Enabled = True
6290                      .cmdRegenerate.Enabled = True
6300                    End If

6310                    DoCmd.Hourglass False
6320                    Beep
6330                    msgResponse = MsgBox("The POSPay file has been created." & vbCrLf & vbCrLf & _
                          "Do you wish to see it?", vbQuestion + vbYesNo + vbDefaultButton2, "Export Successful")

6340                    If msgResponse = vbYes Then
6350                      strDocName = "rptPOSPay_Export_01"
6360                      strSortNow = .frmCheckPOSPay_Sub1.Form.SortNow_Get  ' ** Form Function: frmCheckPOSPay_Sub1.
                          '##GTR_Ref: rptPOSPay_Export_01
6370                      DoCmd.OpenReport strDocName, acViewPreview, , , , THIS_NAME & "~" & strSortNow
6380                      DoCmd.Maximize
6390                      DoCmd.RunCommand acCmdFitToWindow
6400                    End If

6410                  End If  ' ** strPathFile.
6420                End If  ' ** lngCurChecks.
6430                dbs.Close
6440                Set dbs = Nothing

6450              End If  ' ** msgResponse.
6460            Else
6470              If .cmdUpdate.Enabled = False Then .cmdUpdate.Enabled = True
6480              DoCmd.Hourglass False
6490              Beep
6500              MsgBox "You have made changes to the Journal." & vbCrLf & _
                    "Please save your changes by clicking the Update button before proceeding.", _
                    vbInformation + vbOKOnly, "Changes Must Be Saved"
6510            End If  ' ** lngRecsChanged.

6520          End If  ' ** msgResponse.
6530        End If  ' ** lngTmp03.
6540      Else
6550        DoCmd.Hourglass False
6560        Beep
6570        MsgBox "There are no checks in the Journal.", vbInformation + vbOKOnly, "Nothing To Do"
6580      End If  ' ** lngRecsCur.

6590    End With

EXITP:
6600    Set rst1 = Nothing
6610    Set rst2 = Nothing
6620    Set qdf = Nothing
6630    Set dbs = Nothing
6640    Exit Sub

ERRH:
6650    Select Case ERR.Number
        Case Else
6660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6670    End Select
6680    Resume EXITP

End Sub

Private Sub cmdGeneratePOSPay_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdGeneratePOSPay_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6720    intShiftDown = (Shift And acShiftMask) > 0
6730    intAltDown = (Shift And acAltMask) > 0
6740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6760      Select Case intRetVal
          Case vbKeyTab
6770        With Me
6780          intRetVal = 0
6790          If .cmdEdit.Enabled = True Then
6800            .cmdEdit.SetFocus
6810          ElseIf .cmdDelete.Enabled = True Then
6820            .cmdDelete.SetFocus
6830          ElseIf .cmdUpdate.Enabled = True Then
6840            .cmdUpdate.SetFocus
6850          Else
6860            .opgView.SetFocus
6870          End If
6880        End With
6890      End Select
6900    End If

        ' ** Shift keys.
6910    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6920      Select Case intRetVal
          Case vbKeyTab
6930        With Me
6940          intRetVal = 0
6950          If .chkIncludeCurrency.Enabled = True Then
6960            .chkIncludeCurrency.SetFocus
6970          Else
6980            .frmCheckPOSPay_Sub1.SetFocus
6990          End If
7000        End With
7010      End Select
7020    End If

EXITP:
7030    KeyCode = intRetVal
7040    Exit Sub

ERRH:
7050    intRetVal = 0
7060    Select Case ERR.Number
        Case Else
7070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7080    End Select
7090    Resume EXITP

End Sub

Private Sub cmdEdit_Click()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_Click"

7110    With Me
7120      Select Case .frmCheckPOSPay_Sub1.Form.PrintCheck.Locked
          Case True
            ' ** So unlock it for editing.
7130        With .frmCheckPOSPay_Sub1.Form
7140          .PrintCheck_box3.Visible = True
7150          .PrintCheck.Locked = False
7160          .CheckNum.ForeColor = CLR_BLK
7170          .CheckNum.BackColor = CLR_WHT
7180          .CheckNum.Locked = False
7190          .transdate.ForeColor = CLR_BLK
7200          .transdate.BackColor = CLR_WHT
7210          .transdate.Locked = False
7220          .ICash.ForeColor = CLR_BLK
7230          .ICash.BackColor = CLR_WHT
7240          .ICash.Locked = False
7250          .PCash.ForeColor = CLR_BLK
7260          .PCash.BackColor = CLR_WHT
7270          .PCash.Locked = False
7280          .RecurringItem.ForeColor = CLR_BLK
7290          .RecurringItem.BackColor = CLR_WHT
7300          .RecurringItem.Locked = False
7310          .JrnlMemo_Memo.ForeColor = CLR_BLK
7320          .JrnlMemo_Memo.BackColor = CLR_WHT
7330          .JrnlMemo_Memo.Locked = False
7340          .description.ForeColor = CLR_BLK
7350          .description.BackColor = CLR_WHT
7360          .description.Locked = False
7370        End With
7380        .frmCheckPOSPay_Sub1.SetFocus
7390      Case False
            ' ** Lock it up.
7400        .frmCheckPOSPay_Sub1.SetFocus
7410        With .frmCheckPOSPay_Sub1.Form
7420          .FocusHolder.SetFocus
7430          .PrintCheck.Locked = True
7440          .PrintCheck_box3.Visible = False
7450          .CheckNum.Locked = True
7460          .CheckNum.ForeColor = CLR_DISABLED_FG
7470          .CheckNum.BackColor = CLR_DISABLED_BG
7480          .transdate.Locked = True
7490          .transdate.ForeColor = CLR_DISABLED_FG
7500          .transdate.BackColor = CLR_DISABLED_BG
7510          .ICash.Locked = True
7520          .ICash.ForeColor = CLR_DISABLED_FG
7530          .ICash.BackColor = CLR_DISABLED_BG
7540          .PCash.Locked = True
7550          .PCash.ForeColor = CLR_DISABLED_FG
7560          .PCash.BackColor = CLR_DISABLED_BG
7570          .RecurringItem.Locked = True
7580          .RecurringItem.ForeColor = CLR_DISABLED_FG
7590          .RecurringItem.BackColor = CLR_DISABLED_BG
7600          .JrnlMemo_Memo.Locked = True
7610          .JrnlMemo_Memo.ForeColor = CLR_DISABLED_FG
7620          .JrnlMemo_Memo.BackColor = CLR_DISABLED_BG
7630          .description.Locked = True
7640          .description.ForeColor = CLR_DISABLED_FG
7650          .description.BackColor = CLR_DISABLED_BG
7660        End With
7670      End Select
7680    End With

EXITP:
7690    Exit Sub

ERRH:
7700    Select Case ERR.Number
        Case Else
7710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7720    End Select
7730    Resume EXITP

End Sub

Private Sub cmdEdit_KeyDown(KeyCode As Integer, Shift As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEdit_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7820    intShiftDown = (Shift And acShiftMask) > 0
7830    intAltDown = (Shift And acAltMask) > 0
7840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7860      Select Case intRetVal
          Case vbKeyTab
7870        With Me
7880          intRetVal = 0
7890          If .cmdDelete.Enabled = True Then
7900            .cmdDelete.SetFocus
7910          ElseIf .cmdUpdate.Enabled = True Then
7920            .cmdUpdate.SetFocus
7930          Else
7940            .opgView.SetFocus
7950          End If
7960        End With
7970      End Select
7980    End If

        ' ** Shift keys.
7990    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8000      Select Case intRetVal
          Case vbKeyTab
8010        With Me
8020          intRetVal = 0
8030          If .cmdGeneratePOSPay.Enabled = True Then
8040            .cmdGeneratePOSPay.SetFocus
8050          ElseIf .chkIncludeCurrency.Enabled = True Then
8060            .chkIncludeCurrency.SetFocus
8070          Else
8080            .frmCheckPOSPay_Sub1.SetFocus
8090          End If
8100        End With
8110      End Select
8120    End If

EXITP:
8130    KeyCode = intRetVal
8140    Exit Sub

ERRH:
8150    intRetVal = 0
8160    Select Case ERR.Number
        Case Else
8170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8180    End Select
8190    Resume EXITP

End Sub

Private Sub cmdDelete_Click()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, frm As Access.Form
        Dim lngPPJID As Long, lngJrnlID As Long
        Dim lngRecsChanged As Long, lngCheckNum As Long
        Dim strPayee As String
        Dim msgResponse As VbMsgBoxResult

8210    With Me
8220      Set frm = .frmCheckPOSPay_Sub1.Form
8230      If IsNull(frm.ppj_id) = False Then
8240        If frm.ppj_pospay = False Then
8250          If IsNull(frm.CheckNum) = False Then
8260            lngCheckNum = frm.CheckNum
8270          End If
8280          strPayee = frm.RecurringItem
8290          Beep
8300          msgResponse = MsgBox("Are you sure you want to delete this Journal entry?" & vbCrLf & vbCrLf & _
                "Check #: " & IIf(lngCheckNum = 0&, vbNullString, CStr(lngCheckNum)) & vbCrLf & _
                "Payee: " & strPayee, vbQuestion + vbYesNo + vbDefaultButton2, "Delete Journal Entry")
8310          If msgResponse = vbYes Then
8320            DoCmd.Hourglass True
8330            DoEvents
8340            lngPPJID = frm.ppj_id
8350            lngJrnlID = frm.Journal_ID
8360            Set dbs = CurrentDb
8370            With dbs
                  ' ** Delete Journal, by specified [jid].
8380              Set qdf = .QueryDefs("qryCheckPOSPay_18_01")
8390              With qdf.Parameters
8400                ![jid] = lngJrnlID
8410              End With
8420              qdf.Execute
8430              Set qdf = Nothing
8440              DoEvents
                  ' ** Delete tblCheckPOSPay_Journal, by specified [ppjid].
8450              Set qdf = .QueryDefs("qryCheckPOSPay_18_02")
8460              With qdf.Parameters
8470                ![ppjid] = lngPPJID
8480              End With
8490              qdf.Execute
8500              Set qdf = Nothing
8510              DoEvents
8520              .Close
8530            End With  ' ** dbs.
8540            Set dbs = Nothing
8550            DoEvents
8560            frm.Requery
8570            lngRecsCur = frm.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
8580            .CurChecks_lbl.Caption = CStr(lngRecsCur)
8590            lngRecsChanged = HasChanged  ' ** Function: Below.
8600            If lngRecsChanged = 0& And .cmdUpdate.Enabled = True Then
8610              .FocusHolder.SetFocus
8620              .cmdUpdate.Enabled = False
8630            End If
8640            DoCmd.Hourglass False
8650            Beep
8660          End If
8670        Else
8680          Beep
8690        End If
8700      Else
8710        Beep
8720      End If
8730    End With

EXITP:
8740    Set frm = Nothing
8750    Set qdf = Nothing
8760    Set dbs = Nothing
8770    Exit Sub

ERRH:
8780    Select Case ERR.Number
        Case Else
8790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8800    End Select
8810    Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

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
          Case vbKeyTab
8970        With Me
8980          intRetVal = 0
8990          If .cmdUpdate.Enabled = True Then
9000            .cmdUpdate.SetFocus
9010          Else
9020            .opgView.SetFocus
9030          End If
9040        End With
9050      End Select
9060    End If

        ' ** Shift keys.
9070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9080      Select Case intRetVal
          Case vbKeyTab
9090        With Me
9100          intRetVal = 0
9110          .cmdEdit.SetFocus
9120        End With
9130      End Select
9140    End If

EXITP:
9150    KeyCode = intRetVal
9160    Exit Sub

ERRH:
9170    intRetVal = 0
9180    Select Case ERR.Number
        Case Else
9190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9200    End Select
9210    Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecsChanged As Long, lngRecs As Long
        Dim blnContinue As Boolean

9310    With Me

9320      DoCmd.Hourglass True
9330      DoEvents

9340      blnContinue = True

9350      lngRecsChanged = HasChanged  ' ** Function: Below.

9360      Set dbs = CurrentDb
9370      With dbs

9380        If lngRecsChanged > 0& Then

              ' ** Update qryCheckPOSPay_06 (Journal, linked to tblCheckPOSPay_Journal,
              ' ** just ppj_changed = True, with .._new fields), for PrintCheck.
9390          Set qdf = .QueryDefs("qryCheckPOSPay_07_01")
9400          qdf.Execute
9410          Set qdf = Nothing
9420          DoEvents

              ' ** Update qryCheckPOSPay_06 (Journal, linked to tblCheckPOSPay_Journal,
              ' ** just ppj_changed = True, with .._new fields), for CheckNum.
9430          Set qdf = .QueryDefs("qryCheckPOSPay_07_02")
9440          qdf.Execute
9450          Set qdf = Nothing
9460          DoEvents

              ' ** Update qryCheckPOSPay_06 (Journal, linked to tblCheckPOSPay_Journal,
              ' ** just ppj_changed = True, with .._new fields), for transdate.
9470          Set qdf = .QueryDefs("qryCheckPOSPay_07_03")
9480          qdf.Execute
9490          Set qdf = Nothing
9500          DoEvents

              ' ** Update qryCheckPOSPay_06 (Journal, linked to tblCheckPOSPay_Journal,
              ' ** just ppj_changed = True, with .._new fields), for icash.
9510          Set qdf = .QueryDefs("qryCheckPOSPay_07_04")
9520          qdf.Execute
9530          Set qdf = Nothing
9540          DoEvents

              ' ** Update qryCheckPOSPay_06 (Journal, linked to tblCheckPOSPay_Journal,
              ' ** just ppj_changed = True, with .._new fields), for pcash.
9550          Set qdf = .QueryDefs("qryCheckPOSPay_07_05")
9560          qdf.Execute
9570          Set qdf = Nothing
9580          DoEvents

              ' ** Update qryCheckPOSPay_06 (Journal, linked to tblCheckPOSPay_Journal,
              ' ** just ppj_changed = True, with .._new fields), for RecurringItem.
9590          Set qdf = .QueryDefs("qryCheckPOSPay_07_06")
9600          qdf.Execute
9610          Set qdf = Nothing
9620          DoEvents

              ' ** Update qryCheckPOSPay_06 (Journal, linked to tblCheckPOSPay_Journal,
              ' ** just ppj_changed = True, with .._new fields), for description.
9630          Set qdf = .QueryDefs("qryCheckPOSPay_07_07")
9640          qdf.Execute
9650          Set qdf = Nothing
9660          DoEvents

              ' ** Now check the Memo's.

              ' ** tblJournal_Memo, grouped, with cnt_memo.
9670          Set qdf = .QueryDefs("qryCheckPOSPay_08")
9680          Set rst = qdf.OpenRecordset
9690          If rst.BOF = True And rst.EOF = True Then
                ' ** No Memo's whatsoever.
9700            rst.Close
9710            Set rst = Nothing
9720            Set qdf = Nothing
9730            DoEvents
9740          Else
9750            rst.MoveFirst
9760            lngRecs = rst![cnt_memo]
9770            If lngRecs > 0& Then
9780              rst.Close
9790              Set rst = Nothing
9800              Set qdf = Nothing
9810              DoEvents

                  ' ** Delete tblJournal_Memo, for JrnlMemo_Memo = Null.
9820              Set qdf = .QueryDefs("qryCheckPOSPay_09")
9830              qdf.Execute
9840              Set qdf = Nothing
9850              DoEvents

                  ' ** Update qryCheckPOSPay_10 (tblCheckPOSPay_Journal,
                  ' ** linked to tblJournal_Memo, just discrepancies).
9860              Set qdf = .QueryDefs("qryCheckPOSPay_11")
9870              qdf.Execute
9880              Set qdf = Nothing
9890              DoEvents

9900            Else
9910              rst.Close
9920              Set rst = Nothing
9930              Set qdf = Nothing
9940              DoEvents
9950            End If

9960          End If

              ' ** Update tblCheckPOSPay_Journal, for ppj_changed = False.
9970          Set qdf = .QueryDefs("qryCheckPOSPay_12")
9980          qdf.Execute
9990          Set qdf = Nothing
10000         DoEvents

10010       Else
10020         blnContinue = False
10030         DoCmd.Hourglass False
10040         Beep
10050         MsgBox "No changes were found.", vbInformation + vbOKOnly, "Nothing To Do"
10060       End If

10070       .Close
10080     End With  ' ** dbs.
10090     Set dbs = Nothing
10100     DoEvents

10110     If blnContinue = True Then

10120       Sub1_Load  ' ** Procedure: Below
10130       DoEvents

10140       lngRecsChanged = HasChanged  ' ** Function: Below.

10150       .FocusHolder.SetFocus
10160       DoEvents

10170       .frmCheckPOSPay_Sub1.Form.Requery

10180       DoCmd.Hourglass False
10190       Beep
10200       MsgBox "Update successful.", vbInformation + vbOKOnly, "Update Finished"

10210       If lngRecsChanged = 0& Then
10220         .FocusHolder.SetFocus
10230         .cmdUpdate.Enabled = False
10240       End If

10250     End If

10260   End With

EXITP:
10270   Set rst = Nothing
10280   Set qdf = Nothing
10290   Set dbs = Nothing
10300   Exit Sub

ERRH:
10310   DoCmd.Hourglass False
10320   Select Case ERR.Number
        Case Else
10330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10340   End Select
10350   Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10420   intShiftDown = (Shift And acShiftMask) > 0
10430   intAltDown = (Shift And acAltMask) > 0
10440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10460     Select Case intRetVal
          Case vbKeyTab
10470       With Me
10480         intRetVal = 0
10490         .opgView.SetFocus
10500       End With
10510     End Select
10520   End If

        ' ** Shift keys.
10530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10540     Select Case intRetVal
          Case vbKeyTab
10550       With Me
10560         intRetVal = 0
10570         .cmdDelete.SetFocus
10580       End With
10590     End Select
10600   End If

EXITP:
10610   KeyCode = intRetVal
10620   Exit Sub

ERRH:
10630   intRetVal = 0
10640   Select Case ERR.Number
        Case Else
10650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10660   End Select
10670   Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10720   intShiftDown = (Shift And acShiftMask) > 0
10730   intAltDown = (Shift And acAltMask) > 0
10740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10770       With Me
10780         intRetVal = 0
10790         .opgView.SetFocus
10800       End With
10810     End Select
10820   End If

        ' ** Shift keys.
10830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10850       With Me
10860         intRetVal = 0
10870         If .frmCheckPOSPay_Sub3.Visible = True And .frmCheckPOSPay_Sub3.Enabled = True Then
10880           .frmCheckPOSPay_Sub3.SetFocus
10890         ElseIf .frmCheckPOSPay_Sub2.Visible = True And .frmCheckPOSPay_Sub2.Enabled = True Then
10900           .frmCheckPOSPay_Sub2.SetFocus
10910         ElseIf .frmCheckPOSPay_Sub1.Visible = True And .frmCheckPOSPay_Sub1.Enabled = True Then
10920           .frmCheckPOSPay_Sub1.SetFocus
10930         Else
10940           .cmdClose.SetFocus
10950         End If
10960       End With
10970     End Select
10980   End If

EXITP:
10990   KeyCode = intRetVal
11000   Exit Sub

ERRH:
11010   intRetVal = 0
11020   Select Case ERR.Number
        Case Else
11030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11040   End Select
11050   Resume EXITP

End Sub

Private Sub opgView_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_AfterUpdate"

11110   With Me
11120     Select Case .opgView
          Case .opgView_optFiles.OptionValue
11130       .opgView_optFiles_lbl.FontBold = True
11140       .opgView_optChecks_lbl.FontBold = False
11150       .frmCheckPOSPay_Sub2.Visible = True
11160       .frmCheckPOSPay_Sub2_box.Visible = True
11170       .frmCheckPOSPay_Sub3.Visible = False
11180       .frmCheckPOSPay_Sub3_box.Visible = False
11190       .cmdEditPosPay.Enabled = False
11200       lngRecsCur = .frmCheckPOSPay_Sub2.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub2.
11210       If lngRecsCur > 0& Then
11220         .cmdDeletePosPay.Enabled = True
11230         .cmdRegenerate.Enabled = True
11240 On Error Resume Next
11250         .frmCheckPOSPay_Sub2.SetFocus
11260 On Error GoTo ERRH
11270       Else
11280         .cmdDeletePosPay.Enabled = False
11290         .cmdRegenerate.Enabled = False
11300       End If
11310     Case .opgView_optChecks.OptionValue
11320       .opgView_optFiles_lbl.FontBold = False
11330       .opgView_optChecks_lbl.FontBold = True
11340       .frmCheckPOSPay_Sub3.Visible = True
11350       .frmCheckPOSPay_Sub3_box.Visible = True
11360       .FocusHolder.SetFocus
11370       .frmCheckPOSPay_Sub2.Visible = False
11380       .frmCheckPOSPay_Sub2_box.Visible = False
11390       .cmdDeletePosPay.Enabled = False
11400       lngRecsCur = .frmCheckPOSPay_Sub3.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub3.
11410       If lngRecsCur > 0& Then
11420         .cmdEditPosPay.Enabled = True
11430         .cmdRegenerate.Enabled = True
11440 On Error Resume Next
11450         .frmCheckPOSPay_Sub3.SetFocus
11460 On Error GoTo ERRH
11470       Else
11480         .cmdEditPosPay.Enabled = False
11490         .cmdRegenerate.Enabled = False
11500       End If
11510     End Select
11520   End With

EXITP:
11530   Exit Sub

ERRH:
11540   Select Case ERR.Number
        Case Else
11550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11560   End Select
11570   Resume EXITP

End Sub

Private Sub opgView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_MouseMove"

11610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11620   Exit Sub

ERRH:
11630   Select Case ERR.Number
        Case Else
11640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11650   End Select
11660   Resume EXITP

End Sub

Private Sub opgView_optFiles_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_optFiles_KeyDown"

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
11790         If .cmdRegenerate.Enabled = True Then
11800           .cmdRegenerate.SetFocus
11810         ElseIf .cmdEditPosPay.Enabled = True Then
11820           .cmdEditPosPay.SetFocus
11830         ElseIf .cmdDeletePosPay.Enabled = True Then
11840           .cmdDeletePosPay.SetFocus
11850         ElseIf .cmdPrintReport.Enabled = True Then
11860           .cmdPrintReport.SetFocus
11870         Else
11880           .cmdClose.SetFocus
11890         End If
11900       End With
11910     End Select
11920   End If

        ' ** Shift keys.
11930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11950       With Me
11960         intRetVal = 0
11970         If .frmCheckPOSPay_Sub2.Visible = True Then
11980           lngRecsCur = .frmCheckPOSPay_Sub2.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub2.
11990           If lngRecsCur > 0& Then
12000             .frmCheckPOSPay_Sub2.SetFocus  ' ** Let's not do the last record bit.
12010           Else
12020             If .cmdUpdate.Enabled = True Then
12030               .cmdUpdate.SetFocus
12040             ElseIf .cmdDelete.Enabled = True Then
12050               .cmdDelete.SetFocus
12060             ElseIf .cmdEdit.Enabled = True Then
12070               .cmdEdit.SetFocus
12080             ElseIf .cmdGeneratePOSPay.Enabled = True Then
12090               .cmdGeneratePOSPay.SetFocus
12100             Else
12110               lngRecsCur = .frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
12120               If lngRecsCur > 0& Then
12130                 .frmCheckPOSPay_Sub1.SetFocus
12140               Else
12150                 .cmdClose.SetFocus
12160               End If
12170             End If
12180           End If
12190         ElseIf .frmCheckPOSPay_Sub3.Visible = True Then
12200           lngRecsCur = .frmCheckPOSPay_Sub3.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub3.
12210           If lngRecsCur > 0& Then
12220             .frmCheckPOSPay_Sub3.SetFocus
12230           Else
12240             If .cmdUpdate.Enabled = True Then
12250               .cmdUpdate.SetFocus
12260             ElseIf .cmdDelete.Enabled = True Then
12270               .cmdDelete.SetFocus
12280             ElseIf .cmdEdit.Enabled = True Then
12290               .cmdEdit.SetFocus
12300             ElseIf .cmdGeneratePOSPay.Enabled = True Then
12310               .cmdGeneratePOSPay.SetFocus
12320             Else
12330               lngRecsCur = .frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
12340               If lngRecsCur > 0& Then
12350                 .frmCheckPOSPay_Sub1.SetFocus
12360               Else
12370                 .cmdClose.SetFocus
12380               End If
12390             End If
12400           End If
12410         End If
12420       End With
12430     End Select
12440   End If

EXITP:
12450   KeyCode = intRetVal
12460   Exit Sub

ERRH:
12470   intRetVal = 0
12480   Select Case ERR.Number
        Case Else
12490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12500   End Select
12510   Resume EXITP

End Sub

Private Sub opgView_optFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_optFiles_MouseMove"

12610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12620   Exit Sub

ERRH:
12630   Select Case ERR.Number
        Case Else
12640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12650   End Select
12660   Resume EXITP

End Sub

Private Sub opgView_optChecks_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_optChecks_KeyDown"

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
          Case vbKeyTab
12770       With Me
12780         intRetVal = 0
12790         If .cmdRegenerate.Enabled = True Then
12800           .cmdRegenerate.SetFocus
12810         ElseIf .cmdEditPosPay.Enabled = True Then
12820           .cmdEditPosPay.SetFocus
12830         ElseIf .cmdDeletePosPay.Enabled = True Then
12840           .cmdDeletePosPay.SetFocus
12850         ElseIf .cmdPrintReport.Enabled = True Then
12860           .cmdPrintReport.SetFocus
12870         Else
12880           .cmdClose.SetFocus
12890         End If
12900       End With
12910     End Select
12920   End If

        ' ** Shift keys.
12930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12940     Select Case intRetVal
          Case vbKeyTab
12950       With Me
12960         intRetVal = 0
12970         If .frmCheckPOSPay_Sub2.Visible = True Then
12980           lngRecsCur = .frmCheckPOSPay_Sub2.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub2.
12990           If lngRecsCur > 0& Then
13000             .frmCheckPOSPay_Sub2.SetFocus  ' ** Let's not do the last record bit.
13010           Else
13020             If .cmdUpdate.Enabled = True Then
13030               .cmdUpdate.SetFocus
13040             ElseIf .cmdDelete.Enabled = True Then
13050               .cmdDelete.SetFocus
13060             ElseIf .cmdEdit.Enabled = True Then
13070               .cmdEdit.SetFocus
13080             ElseIf .cmdGeneratePOSPay.Enabled = True Then
13090               .cmdGeneratePOSPay.SetFocus
13100             Else
13110               lngRecsCur = .frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
13120               If lngRecsCur > 0& Then
13130                 .frmCheckPOSPay_Sub1.SetFocus
13140               Else
13150                 .cmdClose.SetFocus
13160               End If
13170             End If
13180           End If
13190         ElseIf .frmCheckPOSPay_Sub3.Visible = True Then
13200           lngRecsCur = .frmCheckPOSPay_Sub3.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub3.
13210           If lngRecsCur > 0& Then
13220             .frmCheckPOSPay_Sub3.SetFocus
13230           Else
13240             If .cmdUpdate.Enabled = True Then
13250               .cmdUpdate.SetFocus
13260             ElseIf .cmdDelete.Enabled = True Then
13270               .cmdDelete.SetFocus
13280             ElseIf .cmdEdit.Enabled = True Then
13290               .cmdEdit.SetFocus
13300             ElseIf .cmdGeneratePOSPay.Enabled = True Then
13310               .cmdGeneratePOSPay.SetFocus
13320             Else
13330               lngRecsCur = .frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
13340               If lngRecsCur > 0& Then
13350                 .frmCheckPOSPay_Sub1.SetFocus
13360               Else
13370                 .cmdClose.SetFocus
13380               End If
13390             End If
13400           End If
13410         End If
13420       End With
13430     End Select
13440   End If

EXITP:
13450   KeyCode = intRetVal
13460   Exit Sub

ERRH:
13470   intRetVal = 0
13480   Select Case ERR.Number
        Case Else
13490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13500   End Select
13510   Resume EXITP

End Sub

Private Sub opgView_optChecks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgView_optChecks_MouseMove"

13610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13620   Exit Sub

ERRH:
13630   Select Case ERR.Number
        Case Else
13640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13650   End Select
13660   Resume EXITP

End Sub

Private Sub cmdRegenerate_Click()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRegenerate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngPPID As Long
        Dim strDesc As String, strDocName As String, strSort As String
        Dim strPath As String, strFile As String, strPathFile As String
        Dim blnContinue As Boolean
        Dim msgResponse As VbMsgBoxResult

13710   With Me

13720     blnContinue = False
13730     If .frmCheckPOSPay_Sub2.Visible = True And .frmCheckPOSPay_Sub2.Enabled = True Then
13740       blnContinue = True
13750     ElseIf .frmCheckPOSPay_Sub3.Visible = True And .frmCheckPOSPay_Sub3.Enabled = True Then
13760       blnContinue = True
13770     End If

13780     If blnContinue = True Then
13790       blnContinue = False

            ' ** The 2 suforms are coordinated and accessible whether or not they're visible.
13800       lngRecsCur = .frmCheckPOSPay_Sub2.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub2p.
13810       If lngRecsCur > 0& Then
13820         If IsNull(.frmCheckPOSPay_Sub2.Form.pp_id) = False Then
13830           lngPPID = .frmCheckPOSPay_Sub2.Form.pp_id
13840           strDesc = .frmCheckPOSPay_Sub2.Form.pp_description
13850           Beep
13860           msgResponse = MsgBox("Do you wish to regenerate this POSPay file?" & vbCrLf & vbCrLf & _
                  strDesc, vbQuestion + vbYesNo, "Regenerate POSPay File")
13870           If msgResponse = vbYes Then
13880             blnContinue = True
13890             DoCmd.Hourglass True
13900             DoEvents
13910           End If
13920         Else
13930           Beep
13940         End If  ' ** IsNull().
13950       Else
13960         Beep
13970       End If  ' ** lngRecsCur.

13980     Else
13990       Beep
14000     End If  ' ** blnContinue.

14010     If blnContinue = True Then

14020       Set dbs = CurrentDb
14030       gstrFormQuerySpec = THIS_NAME

            ' ** tblCheckPOSPay, linked to tblCheckPOSPay_Detail, as
            ' ** POSPay file format, in Courier font, by specified FormRef("ppid").
14040       Set qdf = dbs.QueryDefs("qryCheckPOSPay_19_01")
14050       Set rst = qdf.OpenRecordset
14060       If rst.BOF = True And rst.EOF = True Then
              ' ** Something bad happened!
14070         lngCurChecks = 0&
14080         rst.Close
14090         Set rst = Nothing
14100         Set qdf = Nothing
14110         DoCmd.Hourglass False
14120         Beep
14130         MsgBox "A problem as occurred and this process cannot continue.", vbExclamation + vbOKOnly, "Unknown Error"
14140       Else
14150         rst.MoveLast
14160         lngCurChecks = rst.RecordCount
14170         rst.Close
14180         Set rst = Nothing
14190         Set qdf = Nothing
14200       End If

14210       dbs.Close
14220       Set dbs = Nothing

14230       If lngCurChecks > 0& Then

14240         If IsNull(.UserReportPath) = True Then
14250           strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
14260         Else
14270           strPath = .UserReportPath
14280         End If
14290         gdatEndDate = .frmCheckPOSPay_Sub2.Form.pp_date
14300         strFile = "POSPay" & Format(gdatEndDate, "mmddyy") & ".txt"

14310         DoCmd.Hourglass False
14320         strPathFile = FileSaveDialog("txt", strFile, strPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

14330         If strPathFile <> vbNullString Then

14340           DoCmd.Hourglass True
14350           DoEvents

14360           strPath = Parse_Path(strPathFile)  ' ** Module Function: modFileUtilities.
14370           strFile = Parse_File(strPathFile)  ' ** Module Function: modFileUtilities.
14380           .UserReportPath = strPath
14390           gstrFormQuerySpec = THIS_NAME

                ' ** Generate the POSPay file using the export spec I created.
14400           DoCmd.TransferText acExportFixed, "Umpqua_Bank_Export_Specification", "qryCheckPOSPay_19_01", strPathFile, False
14410           DoEvents

14420           DoCmd.Hourglass False
14430           Beep
14440           msgResponse = MsgBox("The POSPay file has been created." & vbCrLf & vbCrLf & _
                  "Do you wish to see it?", vbQuestion + vbYesNo + vbDefaultButton2, "Export Successful")

14450           If msgResponse = vbYes Then
14460             gstrFormQuerySpec = THIS_NAME
14470             strDocName = "rptPOSPay_Export_02"
14480             strSort = .frmCheckPOSPay_Sub3.Form.SortNow_Get  ' ** Form Function: frmCheckPOSPay_Sub3.
                  '##GTR_Ref: rptPOSPay_Export_02
14490             DoCmd.OpenReport strDocName, acViewPreview, , , , THIS_NAME & "~" & strSort
14500             DoCmd.Maximize
14510             DoCmd.RunCommand acCmdFitToWindow
14520           End If

14530         End If  ' ** strPathFile.
14540       End If  ' ** lngCurChecks.
14550     End If  ' ** blnContinue.

14560   End With

EXITP:
14570   Set rst = Nothing
14580   Set qdf = Nothing
14590   Set dbs = Nothing
14600   Exit Sub

ERRH:
14610   Select Case ERR.Number
        Case Else
14620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14630   End Select
14640   Resume EXITP

End Sub

Private Sub cmdRegenerate_KeyDown(KeyCode As Integer, Shift As Integer)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRegenerate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14720   intShiftDown = (Shift And acShiftMask) > 0
14730   intAltDown = (Shift And acAltMask) > 0
14740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14760     Select Case intRetVal
          Case vbKeyTab
14770       With Me
14780         intRetVal = 0
14790         If .cmdEditPosPay.Enabled = True Then
14800           .cmdEditPosPay.SetFocus
14810         ElseIf .cmdDeletePosPay.Enabled = True Then
14820           .cmdDeletePosPay.SetFocus
14830         ElseIf .cmdPrintReport.Enabled = True Then
14840           .cmdPrintReport.SetFocus
14850         Else
14860           .cmdClose.SetFocus
14870         End If
14880       End With
14890     End Select
14900   End If

        ' ** Shift keys.
14910   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14920     Select Case intRetVal
          Case vbKeyTab
14930       With Me
14940         intRetVal = 0
14950         .opgView.SetFocus
14960       End With
14970     End Select
14980   End If

EXITP:
14990   KeyCode = intRetVal
15000   Exit Sub

ERRH:
15010   intRetVal = 0
15020   Select Case ERR.Number
        Case Else
15030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15040   End Select
15050   Resume EXITP

End Sub

Private Sub cmdEditPosPay_Click()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEditPosPay_Click"

15110   With Me
15120     Select Case .frmCheckPOSPay_Sub3.Form.ppd_void.Locked
          Case True
            ' ** So unlock it for editing.
15130       With .frmCheckPOSPay_Sub3.Form
15140         .ppd_void.Locked = False
15150         .ppd_void_box3.Visible = True
15160       End With
15170     Case False
            ' ** Lock it up.
15180       With .frmCheckPOSPay_Sub3.Form
15190         .ppd_void.Locked = True
15200         .ppd_void_box3.Visible = False
15210       End With
15220     End Select
15230   End With

EXITP:
15240   Exit Sub

ERRH:
15250   Select Case ERR.Number
        Case Else
15260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15270   End Select
15280   Resume EXITP

End Sub

Private Sub cmdEditPosPay_KeyDown(KeyCode As Integer, Shift As Integer)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdEditPosPay_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15320   intShiftDown = (Shift And acShiftMask) > 0
15330   intAltDown = (Shift And acAltMask) > 0
15340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15360     Select Case intRetVal
          Case vbKeyTab
15370       With Me
15380         intRetVal = 0
15390         If .cmdDeletePosPay.Enabled = True Then
15400           .cmdDeletePosPay.SetFocus
15410         ElseIf .cmdPrintReport.Enabled = True Then
15420           .cmdPrintReport.SetFocus
15430         Else
15440           .cmdClose.SetFocus
15450         End If
15460       End With
15470     End Select
15480   End If

        ' ** Shift keys.
15490   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15500     Select Case intRetVal
          Case vbKeyTab
15510       With Me
15520         intRetVal = 0
15530         If .cmdRegenerate.Enabled = True Then
15540           .cmdRegenerate.SetFocus
15550         Else
15560           .opgView.SetFocus
15570         End If
15580       End With
15590     End Select
15600   End If

EXITP:
15610   KeyCode = intRetVal
15620   Exit Sub

ERRH:
15630   intRetVal = 0
15640   Select Case ERR.Number
        Case Else
15650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15660   End Select
15670   Resume EXITP

End Sub

Private Sub cmdDeletePosPay_Click()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDeletePosPay_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngPPID As Long
        Dim intMode As Integer
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String, datTmp02 As Date

15710   With Me

15720     DoCmd.Hourglass True
15730     DoEvents

15740     lngRecsCur = .frmCheckPOSPay_Sub2.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub2.
15750     If lngRecsCur > 0& Then
15760       lngRecsCur = .frmCheckPOSPay_Sub3.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub3.
15770       If IsNull(.frmCheckPOSPay_Sub2.Form.pp_id) = False Then
15780         lngPPID = .frmCheckPOSPay_Sub2.Form.pp_id
15790         datTmp02 = .frmCheckPOSPay_Sub2.Form.pp_date
15800         strTmp01 = .frmCheckPOSPay_Sub2.Form.pp_description
15810         msgResponse = MsgBox("Are you sure you want to delete this POSPay record?" & vbCrLf & _
                "Date: " & Format(datTmp02, "mm/dd/yyyy") & vbCrLf & strTmp01, _
                vbQuestion + vbYesNo + vbDefaultButton2, "Delete POSPay Record")
15820         If msgResponse = vbYes Then
15830           DoCmd.Hourglass True
15840           DoEvents
15850           Set dbs = CurrentDb
15860           With dbs
                  ' ** Find out if this is the one currently showing above, or a different one.
                  ' ** intMode = 1 : Journal
                  ' ** intMode = 2 : Ledger

                  ' ** qryCheckPOSPay_16_05 (tblCheckPOSPay_Detail, by specified [ppid]), linked
                  ' ** to qryCheckPOSPay_16_06 (tblCheckPOSPay_Journal, with assettype, ppd_id).
15870             Set qdf = .QueryDefs("qryCheckPOSPay_16_07")
15880             With qdf.Parameters
15890               ![ppid] = lngPPID
15900             End With
15910             Set rst = qdf.OpenRecordset
15920             With rst
15930               If .BOF = True And .EOF = True Then
                      ' ** Not in Journal.
15940                 intMode = 0
15950               Else
15960                 .MoveLast
15970                 If lngRecsCur = .RecordCount Then
15980                   intMode = 1
15990                 Else
                        ' ** I don't see how it could be anything else!
16000                   intMode = -1
16010                 End If
16020               End If
16030               .Close
16040             End With  ' ** rst.
16050             Set rst = Nothing
16060             Set qdf = Nothing
16070             DoEvents

16080             If intMode = 0 Then
                    ' ** It must be an older one.
                    ' ** qryCheckPOSPay_16_05 (tblCheckPOSPay_Detail, by specified [ppid]), linked to Ledger.
16090               Set qdf = .QueryDefs("qryCheckPOSPay_16_08")
16100               With qdf.Parameters
16110                 ![ppid] = lngPPID
16120               End With
16130               Set rst = qdf.OpenRecordset
16140               With rst
16150                 If .BOF = True And .EOF = True Then
                        ' ** OK. They're virtual transactions.
16160                   intMode = -2
16170                 Else
16180                   .MoveLast
16190                   If lngRecsCur = .RecordCount Then
16200                     intMode = 2
16210                   Else
                          ' ** Out of options.
16220                     intMode = -3
16230                   End If
16240                 End If
16250                 .Close
16260               End With  ' ** rst.
16270               Set rst = Nothing
16280               Set qdf = Nothing
16290               DoEvents
16300             End If  ' ** intMode.

16310             Select Case intMode
                  Case 1
                    ' ** Update Journal, for assettype = Null.
16320               Set qdf = .QueryDefs("qryCheckPOSPay_16_11")
16330               qdf.Execute
16340               Set qdf = Nothing
16350               DoEvents
                    ' ** Update tblCheckPOSPay_Journal, for assettype = Null, ppj_pospay = False.
16360               Set qdf = .QueryDefs("qryCheckPOSPay_16_12")
16370               qdf.Execute
16380               Set qdf = Nothing
16390               DoEvents
16400             Case 2
                    ' ** There's no indication in the Ledger related to POSPay.
16410             End Select

                  ' ** Delete tblCheckPOSPay, by specified [ppid].
16420             Set qdf = .QueryDefs("qryCheckPOSPay_16_01")
16430             With qdf.Parameters
16440               ![ppid] = lngPPID
16450             End With
16460             qdf.Execute
16470             Set qdf = Nothing

16480             .Close
16490           End With
16500           Set dbs = Nothing

16510           .frmCheckPOSPay_Sub1.Form.Requery
16520           .frmCheckPOSPay_Sub2.Form.Requery
16530           DoEvents

16540           lngRecsCur = .frmCheckPOSPay_Sub2.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub2.
16550           If lngRecsCur = 0& Then
16560             .FocusHolder.SetFocus
16570             DoEvents
16580             .cmdDeletePosPay.Enabled = False
16590             .cmdRegenerate.Enabled = False
16600             .pp_date = Null
16610             .pp_description = Null
16620             .frmCheckPOSPay_Sub2.Form.Sub2_Disable False  ' ** Form Function: frmCheckPOSPay_Sub2.
16630             .FocusHolder.SetFocus
16640             .frmCheckPOSPay_Sub3.Form.Sub3_Disable False  ' ** Form Function: frmCheckPOSPay_Sub3.
16650             .FocusHolder.SetFocus
16660           End If

16670           Beep
16680           DoCmd.Hourglass False

16690         End If
16700       Else
16710         DoCmd.Hourglass False
16720         Beep
16730         MsgBox "No record has been selected.", vbInformation + vbOKOnly, "Nothing To Do"
16740       End If
16750     Else
16760       DoCmd.Hourglass False
16770       Beep
16780       MsgBox "There are no records.", vbInformation + vbOKOnly, "Nothing To Do"
16790     End If

16800   End With

EXITP:
16810   Set rst = Nothing
16820   Set qdf = Nothing
16830   Set dbs = Nothing
16840   Exit Sub

ERRH:
16850   DoCmd.Hourglass False
16860   Select Case ERR.Number
        Case Else
16870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16880   End Select
16890   Resume EXITP

End Sub

Private Sub cmdDeletePosPay_KeyDown(KeyCode As Integer, Shift As Integer)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDeletePosPay_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16920   intShiftDown = (Shift And acShiftMask) > 0
16930   intAltDown = (Shift And acAltMask) > 0
16940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16960     Select Case intRetVal
          Case vbKeyTab
16970       With Me
16980         intRetVal = 0
16990         If .cmdPrintReport.Enabled = True Then
17000           .cmdPrintReport.SetFocus
17010         Else
17020           .cmdClose.SetFocus
17030         End If
17040       End With
17050     End Select
17060   End If

        ' ** Shift keys.
17070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17080     Select Case intRetVal
          Case vbKeyTab
17090       With Me
17100         intRetVal = 0
17110         If .cmdEditPosPay.Enabled = True Then
17120           .cmdEditPosPay.SetFocus
17130         ElseIf .cmdRegenerate.Enabled = True Then
17140           .cmdRegenerate.SetFocus
17150         Else
17160           .opgView.SetFocus
17170         End If
17180       End With
17190     End Select
17200   End If

EXITP:
17210   KeyCode = intRetVal
17220   Exit Sub

ERRH:
17230   intRetVal = 0
17240   Select Case ERR.Number
        Case Else
17250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17260   End Select
17270   Resume EXITP

End Sub

Private Sub cmdPrintReport_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_lbl2_MouseMove"

17310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17320   Exit Sub

ERRH:
17330   Select Case ERR.Number
        Case Else
17340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17350   End Select
17360   Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

17410   With Me
17420     blnPrintReport_Focus = True
17430     .cmdPrintReport_raised_semifocus_dots_img.Visible = True
17440     .cmdPrintReport_raised_img.Visible = False
17450     .cmdPrintReport_raised_focus_img.Visible = False
17460     .cmdPrintReport_raised_focus_dots_img.Visible = False
17470     .cmdPrintReport_sunken_focus_dots_img.Visible = False
17480     .cmdPrintReport_raised_img_dis.Visible = False
17490   End With

EXITP:
17500   Exit Sub

ERRH:
17510   Select Case ERR.Number
        Case Else
17520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17530   End Select
17540   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

17610   With Me
17620     blnPrintReport_MouseDown = True
17630     .cmdPrintReport_sunken_focus_dots_img.Visible = True
17640     .cmdPrintReport_raised_img.Visible = False
17650     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
17660     .cmdPrintReport_raised_focus_img.Visible = False
17670     .cmdPrintReport_raised_focus_dots_img.Visible = False
17680     .cmdPrintReport_raised_img_dis.Visible = False
17690   End With

EXITP:
17700   Exit Sub

ERRH:
17710   Select Case ERR.Number
        Case Else
17720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17730   End Select
17740   Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strSort As String

17810   With Me
17820     lngRecsCur = .frmCheckPOSPay_Sub2.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub2.
17830     If lngRecsCur > 0& Then
17840       If IsNull(.pp_id) = False Then
17850         If .pp_id > 0& Then
                ' ** The report's query uses FormRef() to get the current pp_id of Sub 2.
17860           gstrFormQuerySpec = THIS_NAME
17870           strDocName = "rptPOSPay_Export_02"
17880           strSort = .frmCheckPOSPay_Sub3.Form.SortNow_Get  ' ** Form Function: frmCheckPOSPay_Sub3.
17890           If gblnDev_Debug Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
17900             DoCmd.OpenReport strDocName, acViewPreview, , , , THIS_NAME & "~" & strSort
17910             DoCmd.Maximize
17920             DoCmd.RunCommand acCmdFitToWindow
17930             .Visible = False
17940           Else
17950             DoCmd.OpenReport strDocName, acViewNormal, , , , THIS_NAME & "~" & strSort
17960           End If
17970         Else
17980           Beep
17990         End If
18000       Else
18010         Beep
18020       End If
18030     Else
18040       Beep
18050       MsgBox "There are no previously generated files.", vbInformation + vbOKOnly, "Nothing To Do"
18060     End If
18070   End With

EXITP:
18080   Exit Sub

ERRH:
18090   Select Case ERR.Number
        Case Else
18100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18110   End Select
18120   Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18220   intShiftDown = (Shift And acShiftMask) > 0
18230   intAltDown = (Shift And acAltMask) > 0
18240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18260     Select Case intRetVal
          Case vbKeyTab
18270       With Me
18280         intRetVal = 0
18290         .cmdClose.SetFocus
18300       End With
18310     End Select
18320   End If

        ' ** Shift keys.
18330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18340     Select Case intRetVal
          Case vbKeyTab
18350       With Me
18360         intRetVal = 0
18370         If .cmdDeletePosPay.Enabled = True Then
18380           .cmdDeletePosPay.SetFocus
18390         ElseIf .cmdEditPosPay.Enabled = True Then
18400           .cmdEditPosPay.SetFocus
18410         ElseIf .cmdRegenerate.Enabled = True Then
18420           .cmdRegenerate.SetFocus
18430         Else
18440           .opgView.SetFocus
18450         End If
18460       End With
18470     End Select
18480   End If

EXITP:
18490   KeyCode = intRetVal
18500   Exit Sub

ERRH:
18510   intRetVal = 0
18520   Select Case ERR.Number
        Case Else
18530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18540   End Select
18550   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

18610   With Me
18620     If blnPrintReport_MouseDown = False Then
18630       Select Case blnPrintReport_Focus
            Case True
18640         .cmdPrintReport_raised_focus_dots_img.Visible = True
18650         .cmdPrintReport_raised_focus_img.Visible = False
18660       Case False
18670         .cmdPrintReport_raised_focus_img.Visible = True
18680         .cmdPrintReport_raised_focus_dots_img.Visible = False
18690       End Select
18700       .cmdPrintReport_raised_img.Visible = False
18710       .cmdPrintReport_raised_semifocus_dots_img.Visible = False
18720       .cmdPrintReport_sunken_focus_dots_img.Visible = False
18730       .cmdPrintReport_raised_img_dis.Visible = False
18740     End If
18750   End With

EXITP:
18760   Exit Sub

ERRH:
18770   Select Case ERR.Number
        Case Else
18780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18790   End Select
18800   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

18910   With Me
18920     .cmdPrintReport_raised_focus_dots_img.Visible = True
18930     .cmdPrintReport_raised_img.Visible = False
18940     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
18950     .cmdPrintReport_raised_focus_img.Visible = False
18960     .cmdPrintReport_sunken_focus_dots_img.Visible = False
18970     .cmdPrintReport_raised_img_dis.Visible = False
18980     blnPrintReport_MouseDown = False
18990   End With

EXITP:
19000   Exit Sub

ERRH:
19010   Select Case ERR.Number
        Case Else
19020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19030   End Select
19040   Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

19110   With Me
19120     .cmdPrintReport_raised_img.Visible = True
19130     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
19140     .cmdPrintReport_raised_focus_img.Visible = False
19150     .cmdPrintReport_raised_focus_dots_img.Visible = False
19160     .cmdPrintReport_sunken_focus_dots_img.Visible = False
19170     .cmdPrintReport_raised_img_dis.Visible = False
19180     blnPrintReport_Focus = False
19190   End With

EXITP:
19200   Exit Sub

ERRH:
19210   Select Case ERR.Number
        Case Else
19220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19230   End Select
19240   Resume EXITP

End Sub

Public Function HasChanged() As Long

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "HasChanged"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRetVal As Long

19310   lngRetVal = 0&

19320   Set dbs = CurrentDb
19330   With dbs

          ' ** tblCheckPOSPay_Journal, just ppj_changed = True.
19340     Set qdf = .QueryDefs("qryCheckPOSPay_05")
19350     Set rst = qdf.OpenRecordset
19360     With rst
19370       If .BOF = True And .EOF = True Then
19380         lngRetVal = 0&
19390       Else
19400         .MoveLast
19410         lngRetVal = .RecordCount
19420       End If

19430       .Close
19440     End With
19450     DoEvents

19460     .Close
19470   End With  ' ** dbs.

EXITP:
19480   Set rst = Nothing
19490   Set qdf = Nothing
19500   Set dbs = Nothing
19510   HasChanged = lngRetVal
19520   Exit Function

ERRH:
19530   DoCmd.Hourglass False
19540   lngRetVal = 0&
19550   Select Case ERR.Number
        Case Else
19560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19570   End Select
19580   Resume EXITP

End Function

Private Sub Sub1_Load()

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "Sub1_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngVoids As Long, lngPPID As Long
        Dim lngRecs As Long
        Dim lngAccts As Long
        Dim strMsg As String
        Dim lngX As Long

19610   With Me

19620     DoCmd.Hourglass True  ' ** Make sure it's still running.
19630     DoEvents

19640     Set dbs = CurrentDb

          ' ** Journal, linked to Account, just PrintCheck = True.
19650     Set qdf = dbs.QueryDefs("qryCheckPOSPay_01")
19660     Set rst = qdf.OpenRecordset
19670     If rst.BOF = True And rst.EOF = True Then
            ' ** There are no checks in the Journal.
19680       lngCurChecks = 0&
19690       rst.Close
19700       Set rst = Nothing
19710       Set qdf = Nothing
19720       .NoChecks_lbl.Visible = True
19730       DoEvents
19740     Else
19750       rst.MoveLast
19760       lngCurChecks = rst.RecordCount
19770       blnHasChecks = True
19780       rst.Close
19790       Set rst = Nothing
19800       Set qdf = Nothing
19810       .NoChecks_lbl.Visible = False
19820       DoEvents
19830     End If

19840     .CurChecks_lbl.Caption = CStr(lngCurChecks)
19850     DoEvents

          ' ** Empty tblCheckPOSPay_Journal.
19860     Set qdf = dbs.QueryDefs("qryCheckPOSPay_02")
19870     qdf.Execute
19880     Set qdf = Nothing
19890     DoEvents

19900     DoCmd.Hourglass True  ' ** Make sure it's still running.
19910     DoEvents

19920     If blnHasChecks = True Then

            ' ** Append qryCheckPOSPay_01 (Journal, linked to Account,
            ' ** just PrintCheck = True) to tblCheckPOSPay_Journal.
19930       Set qdf = dbs.QueryDefs("qryCheckPOSPay_03")
19940       qdf.Execute
19950       Set qdf = Nothing
19960       DoEvents

            ' ** Update tblCheckPOSPay_Journal, via subquery to qryCheckPOSPay_01_02
            ' ** (qryCheckPOSPay_01_01 (tblCheckPOSPay_Journal, with assettype, pp_id), linked to
            ' ** tblCheckPOSPay_Detail, for Journal entries already generated), for ppj_pospay = True.
19970       Set qdf = dbs.QueryDefs("qryCheckPOSPay_01_03")
19980       qdf.Execute
19990       Set qdf = Nothing
20000       DoEvents

20010       lngVoids = .frmCheckPOSPay_Sub3.Form.VoidCount  ' ** Form Function: frmCheckPOSPay_Sub3.
20020       .VoidCount = lngVoids
20030       If lngVoids > 0& Then
20040         lngPPID = .pp_id
              ' ** qryCheckPOSPay_16_05 (tblCheckPOSPay_Detail,
              ' ** by specified [ppid]), linked to qryCheckPOSPay_16_06
              ' ** (tblCheckPOSPay_Journal, with assettype, ppd_id).
20050         Set qdf = dbs.QueryDefs("qryCheckPOSPay_16_07")
20060         With qdf.Parameters
20070           ![ppid] = lngPPID
20080         End With
20090         Set rst = qdf.OpenRecordset
20100         If rst.BOF = True And rst.EOF = True Then
                ' ** Not the droids we're looking for.
20110           rst.Close
20120           Set rst = Nothing
20130           DoEvents
20140         Else
                ' ** Sub3 is showing Sub1, so update Sub1 with Sub3's Void's.
20150           rst.Close
20160           Set rst = Nothing
20170           DoEvents
20180           SetVoids  ' ** Procedure: Below.
20190           .opgView = .opgView_optChecks.OptionValue
20200           opgView_AfterUpdate  ' ** Procedure: Below.
20210         End If
20220       End If  ' ** lngVoids.

20230       DoCmd.Hourglass True  ' ** Make sure it's still running.
20240       DoEvents

            ' ** tblCheckPOSPay_Journal, just Bank_RoutingNumber = Null, Bank_Account_Number = Null.
20250       Set qdf = dbs.QueryDefs("qryCheckPOSPay_26_01")
20260       Set rst = qdf.OpenRecordset
20270       If rst.BOF = True And rst.EOF = True Then
              ' ** That's good.
20280         rst.Close
20290         Set rst = Nothing
20300         Set qdf = Nothing
20310       Else
20320         rst.MoveLast
20330         lngRecs = rst.RecordCount
20340         rst.Close
20350         Set rst = Nothing
20360         Set qdf = Nothing
              ' ** qryCheckPOSPay_26_01 (tblCheckPOSPay_Journal, just Bank_RoutingNumber = Null,
              ' ** Bank_Account_Number = Null), grouped by accountno, with cnt_chk.
20370         Set qdf = dbs.QueryDefs("qryCheckPOSPay_26_02")
20380         Set rst = qdf.OpenRecordset
20390         If rst.BOF = True And rst.EOF = True Then
                ' ** Shouldn't happen.
20400           rst.Close
20410           Set rst = Nothing
20420           Set qdf = Nothing
20430         Else
20440           rst.MoveLast
20450           lngAccts = rst.RecordCount
20460           rst.MoveFirst
20470           strMsg = CStr(lngRecs) & " of your checks " & IIf(lngRecs = 1&, "does", "do") & " " & _
                  "not have a bank routing number or bank" & vbCrLf & _
                  "account number, and cannot be included in the POSPay file." & vbCrLf & vbCrLf
20480           For lngX = 1& To lngAccts
20490             strMsg = strMsg & rst![accountno] & "  " & rst![shortname] & vbCrLf & "    Checks:  " & CStr(rst![cnt_chk]) & vbCrLf
20500             If lngX < lngAccts Then rst.MoveNext
20510           Next
20520           strMsg = strMsg & vbCrLf & "Add the information to the Account Profile to include them in this file."
20530           DoCmd.Hourglass False
20540           Beep
20550           MsgBox strMsg, vbInformation + vbOKOnly, "Missing Bank Numbers"
20560         End If
20570       End If

20580       DoCmd.Hourglass True  ' ** Make sure it's still running.
20590       DoEvents

            ' ** tblCheckPOSPay_Journal, just CheckNum = Null.
20600       Set qdf = dbs.QueryDefs("qryCheckPOSPay_27_01")
20610       Set rst = qdf.OpenRecordset
20620       If rst.BOF = True And rst.EOF = True Then
              ' ** No missing check numbers.
20630         rst.Close
20640         Set rst = Nothing
20650         Set qdf = Nothing
20660       Else
              ' ** Missing check numbers found.
20670         rst.MoveLast
20680         lngRecs = rst.RecordCount
20690         rst.Close
20700         Set rst = Nothing
20710         Set qdf = Nothing
20720         strMsg = CStr(lngRecs) & " of your checks " & IIf(lngRecs = 1&, "does", "do") & " not " & _
                "have a check number," & vbCrLf & "and cannot be included in the POSPay file." & vbCrLf & vbCrLf
20730         strMsg = strMsg & "Print the checks first, then return here to generate the file."
20740         DoCmd.Hourglass False
20750         Beep
20760         MsgBox strMsg, vbInformation + vbOKOnly, "Missing Check Numbers"
20770       End If

20780       DoCmd.Hourglass True  ' ** Make sure it's still running.
20790       DoEvents

20800       .frmCheckPOSPay_Sub1.Form.Requery

20810       .Issue_Date = .frmCheckPOSPay_Sub1.Form.transdate
20820       .FocusHolder.SetFocus
20830       .cmdUpdate.Enabled = False

20840       If .frmCheckPOSPay_Sub1.Form.ppj_pospay = False Then
20850         .cmdGeneratePOSPay.Enabled = True
20860       End If

20870     Else
20880       .frmCheckPOSPay_Sub1.Form.Requery
20890       .FocusHolder.SetFocus
20900       .cmdGeneratePOSPay.Enabled = False
20910       .cmdEdit.Enabled = False
20920       .cmdDelete.Enabled = False
20930       .cmdUpdate.Enabled = False
20940       .Issue_Date.Enabled = False
20950       .Issue_Date.BackStyle = acBackStyleTransparent
20960       .Issue_Date.BorderColor = WIN_CLR_DISR
20970       .frmCheckPOSPay_Sub1.Form.Sub1_Disable False  ' ** Form Procedure: frmCheckPOSPay_Sub1.
20980       DoEvents
20990       .FocusHolder.SetFocus
21000       .frmCheckPOSPay_Sub1.Enabled = False
21010     End If  ' ** blnHasChecks.

21020   End With

EXITP:
21030   Set rst = Nothing
21040   Set qdf = Nothing
21050   Set dbs = Nothing
21060   Exit Sub

ERRH:
21070   Select Case ERR.Number
        Case Else
21080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21090   End Select
21100   Resume EXITP

End Sub

Private Sub SetVoids()

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "SetVoids"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

21210   With Me
          ' ** Sub3 has Void's, and Sub1 is Sub3.
21220     gstrFormQuerySpec = THIS_NAME
21230     Set dbs = CurrentDb
          ' ** Update qryCheckPOSPay_16_14 (tblCheckPOSPay_Journal, with DLookups() to
          ' ** qryCheckPOSPay_16_13 (tblCheckPOSPay_Detail, linked to qryCheckPOSPay_04
          ' ** (tblCheckPOSPay_Journal, linked to tblJournal_Memo), just
          ' ** ppd_void = True, by specified FormRef("ppid"))).
21240     Set qdf = dbs.QueryDefs("qryCheckPOSPay_16_15")
21250     qdf.Execute
21260     Set qdf = Nothing
21270     .frmCheckPOSPay_Sub1.Form.Refresh
21280     DoEvents
21290     .frmCheckPOSPay_Sub1.Form.ShowVoid True  ' ** Form Procedure: frmCheckPOSPay_Sub1.
21300     .frmCheckPOSPay_Sub1.Form.Refresh
21310     DoEvents

21320   End With

EXITP:
21330   Set qdf = Nothing
21340   Set dbs = Nothing
21350   Exit Sub

ERRH:
21360   Select Case ERR.Number
        Case Else
21370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21380   End Select
21390   Resume EXITP

End Sub
