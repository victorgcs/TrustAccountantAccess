VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckPOSPay_Sub3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckPOSPay_Sub3"

'VGC 03/25/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Issue Date:       A {Issue_Date on frmCheckPOSPay}
' **   Check Number:     B {ppd_checknum}
' **   Close:            C {cmdClose on frmCheckPOSPay}
' **   Delete:           D {cmdDelete on frmCheckPOSPay}
' **   Edit:             E {cmdEdit on frmCheckPOSPay}
' **   Files:            F {opgView_optFiles on frmCheckPOSPay}
' **   Generate POSPay:  G {cmdGeneratePOSPay on frmCheckPOSPay}
' **   Checks:           H {opgView_optChecks on frmCheckPOSPay}
' **   Edit POSPay:      I {cmdEditPosPay on frmCheckPOSPay}
' **   Check Number:     K {CheckNum on frmCheckPOSPay_Sub1}
' **   Include Currency: L {chkIncludeCurrency on frmCheckPOSPay}
' **   Check Memo:       M {JrnlMemo_Memo on frmCheckPOSPay_Sub1}
' **   Regenerate:       N {cmdRegenerate on frmCheckPOSPay}
' **   Comments:         O {description on frmCheckPOSPay_Sub1}
' **   Print Report:     P {cmdPrintReport on frmCheckPOSPay}
' **   Currency:         R {curr_id on frmCheckPOSPay_Sub1}
' **   Posting Date:     S {transdate on frmCheckPOSPay_Sub1}
' **   Delete POSPay:    T {cmdDeletePosPay on frmCheckPOSPay}
' **   Update:           U {cmdUpdate on frmCheckPOSPay}
' **   Void:             V {ppd_void}
' **   Exit:             X {cmdClose on frmCheckPOSPay}
' **   Payee:            Y {RecurringItem on frmCheckPOSPay_Sub1}

'I CAN'T FIGURE OUT WHAT TO GIVE THESE!
' **   Description:      P {pp_description on frmCheckPOSPay_Sub2}  ' ** Not visible when here.
' **   Payee:            P {ppd_payee}

' ** Shortcut Ctrl keys responsive from this form:
' **   Next:             N {MoveRec on frmCheckPOSPay_Sub2}
' **   Previous:         P {MoveRec on frmCheckPOSPay_Sub2}
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[Journal_ID]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private varOldValue As Variant
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
150       If strCallingForm = "frmCheckPOSPay" Then

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

410     With Me

420       CLR_DISABLED_FG = CLR_DKGRY
430       CLR_DISABLED_BG = CLR_LTTEAL

440       .ppd_id.ForeColor = CLR_DISABLED_FG
450       .ppd_id.BackColor = CLR_DISABLED_BG
460       .journalno.ForeColor = CLR_DISABLED_FG
470       .journalno.BackColor = CLR_DISABLED_BG
480       .Journal_ID.ForeColor = CLR_DISABLED_FG
490       .Journal_ID.BackColor = CLR_DISABLED_BG
500       .ppd_checknum.ForeColor = CLR_DISABLED_FG
510       .ppd_checknum.BackColor = CLR_DISABLED_BG
520       .accountno.ForeColor = CLR_DISABLED_FG
530       .accountno.BackColor = CLR_DISABLED_BG
540       .shortname.ForeColor = CLR_DISABLED_FG
550       .shortname.BackColor = CLR_DISABLED_BG
560       .ppd_issuedate.ForeColor = CLR_DISABLED_FG
570       .ppd_issuedate.BackColor = CLR_DISABLED_BG
580       .ppd_amount.ForeColor = CLR_DISABLED_FG
590       .ppd_amount.BackColor = CLR_DISABLED_BG
600       .curr_id.ForeColor = CLR_DISABLED_FG
610       .curr_id.BackColor = CLR_DISABLED_BG
620       .ppd_payee.ForeColor = CLR_DISABLED_FG
630       .ppd_payee.BackColor = CLR_DISABLED_BG
640       .ppd_bank_name.ForeColor = CLR_DISABLED_FG
650       .ppd_bank_name.BackColor = CLR_DISABLED_BG
660       .ppd_aba_trc.ForeColor = CLR_DISABLED_FG
670       .ppd_aba_trc.BackColor = CLR_DISABLED_BG
680       .ppd_bank_account.ForeColor = CLR_DISABLED_FG
690       .ppd_bank_account.BackColor = CLR_DISABLED_BG
700       .Username.ForeColor = CLR_DISABLED_FG
710       .Username.BackColor = CLR_DISABLED_BG

720       SortNow THIS_PROC  ' ** Procedure: Below.

730     End With

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

Private Sub Form_Current()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

810     With Me
820       Select Case IsNull(.journalno)
          Case True
830         If .journalno.Visible = True Or .journalno_lbl.Visible = True Then
840           .journalno.Visible = False
850           .journalno_lbl.Visible = False
860           .journalno_lbl2.Visible = False
870           .journalno_lbl_line.Visible = False
880           .Journal_ID.Visible = True
890           .Journal_ID_lbl.Visible = True
900           .Journal_ID_lbl2.Visible = True
910           .Journal_ID_lbl_line.Visible = True
920         End If
930       Case False
940         If .Journal_ID.Visible = True Or .Journal_ID_lbl.Visible = True Then
950           .Journal_ID.Visible = False
960           .Journal_ID_lbl.Visible = False
970           .Journal_ID_lbl2.Visible = False
980           .Journal_ID_lbl_line.Visible = False
990           .journalno.Visible = True
1000          .journalno_lbl.Visible = True
1010          .journalno_lbl2.Visible = True
1020          .journalno_lbl_line.Visible = True
1030        End If
1040      End Select
1050    End With

EXITP:
1060    Exit Sub

ERRH:
1070    DoCmd.Hourglass False
1080    Select Case ERR.Number
        Case Else
1090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1100    End Select
1110    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Check Number:     B {ppd_checknum}
        ' **   Check Number:     K {CheckNum on frmCheckPOSPay_Sub1}
        ' **   Check Memo:       M {JrnlMemo_Memo on frmCheckPOSPay_Sub1}
        ' **   Comments:         O {description on frmCheckPOSPay_Sub1}
        ' **   Currency:         R {curr_id on frmCheckPOSPay_Sub1}
        ' **   Posting Date:     S {transdate on frmCheckPOSPay_Sub1}
        ' **   Void:             V {ppd_void)
        ' **   Exit:             X {cmdClose on frmCheckPOSPay}
        ' **   Payee:            Y {RecurringItem on frmCheckPOSPay_Sub1}

        'I CAN'T FIGURE OUT WHAT TO GIVE THIS!
        ' **   Payee:            P {ppd_payee}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next:             N {MoveRec on frmCheckPOSPay_Sub2}
        ' **   Previous:         P {MoveRec on frmCheckPOSPay_Sub2}
        ' **   Save:             S {cmdSave on frmCheckPOSPay_Sub1}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyUp
1270        With Me
1280          intRetVal = 0
1290          MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1300        End With
1310      Case vbKeyDown
1320        With Me
1330          intRetVal = 0
1340          MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1350        End With
1360      End Select
1370    End If

        ' ** Alt keys.
1380    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1390      Select Case intRetVal
          Case vbKeyB
1400        With Me
1410          intRetVal = 0
1420          .ppd_checknum.SetFocus
1430        End With
1440      Case vbKeyK
1450        With Me
1460          intRetVal = 0
1470          DoCmd.SelectObject acForm, .Parent.Name, False
1480          .Parent.frmCheckPOSPay_Sub1.SetFocus
1490          .Parent.frmCheckPOSPay_Sub1.Form.CheckNum.SetFocus
1500        End With
1510      Case vbKeyM
1520        With Me
1530          intRetVal = 0
1540          DoCmd.SelectObject acForm, .Parent.Name, False
1550          .Parent.frmCheckPOSPay_Sub1.SetFocus
1560          .Parent.frmCheckPOSPay_Sub1.Form.JrnlMemo_Memo.SetFocus
1570        End With
1580      Case vbKeyO
1590        With Me
1600          intRetVal = 0
1610          DoCmd.SelectObject acForm, .Parent.Name, False
1620          .Parent.frmCheckPOSPay_Sub1.SetFocus
1630          .Parent.frmCheckPOSPay_Sub1.Form.description.SetFocus
1640        End With
'1650      Case vbKeyP
'1660        With Me
'1670          intRetVal = 0
'1680          .ppd_payee.SetFocus
'1690        End With
1650      Case vbKeyR
1660        With Me
1670          intRetVal = 0
1680          DoCmd.SelectObject acForm, .Parent.Name, False
1690          .Parent.frmCheckPOSPay_Sub1.SetFocus
1700          .Parent.frmCheckPOSPay_Sub1.Form.curr_id.SetFocus
1710        End With
1720      Case vbKeyS
1730        With Me
1740          intRetVal = 0
1750          DoCmd.SelectObject acForm, .Parent.Name, False
1760          .Parent.frmCheckPOSPay_Sub1.SetFocus
1770          .Parent.frmCheckPOSPay_Sub1.Form.transdate.SetFocus
1780        End With
1790      Case vbKeyV
1800        With Me
1810          intRetVal = 0
1820          .FocusHolder2.SetFocus
1830          DoEvents
1840          .ppd_void.SetFocus
1850          If .ppd_void.Locked = False Then
1860            .ppd_void = (Not .ppd_void)
1870            ppd_void_AfterUpdate  ' ** Procedure: Below.
1880          End If
1890        End With
1900      Case vbKeyX
1910        With Me
1920          intRetVal = 0
1930          .Parent.cmdClose_Click  ' ** Form Procedure: frmCheckPOSPay.
1940        End With
1950      Case vbKeyY
1960        With Me
1970          intRetVal = 0
1980          DoCmd.SelectObject acForm, .Parent.Name, False
1990          .Parent.frmCheckPOSPay_Sub1.SetFocus
2000          .Parent.frmCheckPOSPay_Sub1.Form.RecurringItem.SetFocus
2010        End With
2020      End Select
2030    End If

        ' ** Ctrl keys.
2040    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2050      Select Case intRetVal
          Case vbKeyN
2060        With Me
2070          intRetVal = 0
2080          .Parent.frmCheckPOSPay_Sub2.Form.MoveRecX acCmdRecordsGoToNext  ' ** Form Procedure: frmCheckPOSPay_Sub2.
2090        End With
2100      Case vbKeyP
2110        With Me
2120          intRetVal = 0
2130          .Parent.frmCheckPOSPay_Sub2.Form.MoveRecX acCmdRecordsGoToPrevious  ' ** Form Procedure: frmCheckPOSPay_Sub2.
2140        End With
2150      Case vbKeyS
2160        intRetVal = 0
2170        cmdSave_Click  ' ** Procedure: Below.
2180      Case vbKeyTab, vbKeyReturn
2190        With Me
2200          intRetVal = 0
2210          DoCmd.SelectObject acForm, .Parent.Name, False
2220          If .Parent.cmdRegenerate.Enabled = True Then
2230            .Parent.cmdRegenerate.SetFocus
2240          ElseIf .Parent.cmdDeletePosPay.Enabled = True Then
2250            .Parent.cmdDeletePosPay.SetFocus
2260          Else
2270            .Parent.cmdClose.SetFocus
2280          End If
2290        End With
2300      Case vbKeyPageUp
2310        intRetVal = 0
2320        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2330      Case vbKeyPageDown
2340        intRetVal = 0
2350        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2360      End Select
2370    End If

        ' ** Ctrl-Shift keys.
2380    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2390      Select Case intRetVal
          Case vbKeyF
2400        With Me
2410          intRetVal = 0
2420          DoCmd.SelectObject acForm, .Parent.Name, False
2430          .Parent.FocusHolder.SetFocus
2440        End With
2450      Case vbKeyTab, vbKeyReturn
2460        With Me
2470          intRetVal = 0
2480          DoCmd.SelectObject acForm, .Parent.Name, False
2490          If .Parent.cmdUpdate.Enabled = True Then
2500            .Parent.cmdUpdate.SetFocus
2510          ElseIf .Parent.cmdDelete.Enabled = True Then
2520            .Parent.cmdDelete.SetFocus
2530          ElseIf .Parent.cmdEdit.Enabled = True Then
2540            .Parent.cmdEdit.SetFocus
2550          ElseIf .Parent.cmdGeneratePOSPay.Enabled = True Then
2560            .Parent.cmdGeneratePOSPay.SetFocus
2570          Else
2580            lngRecsCur = .Parent.frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
2590            If lngRecsCur > 0& Then
2600              .Parent.frmCheckPOSPay_Sub1.SetFocus
2610            Else
2620              .Parent.cmdClose.SetFocus
2630            End If
2640          End If
2650        End With
2660      End Select
2670    End If

EXITP:
2680    KeyCode = intRetVal
2690    Exit Sub

ERRH:
2700    intRetVal = 0
2710    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2720    Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
2730    Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2810    DoCmd.RunCommand acCmdSaveRecord
2820    DoEvents

EXITP:
2830    Exit Sub

ERRH:
2840    Select Case ERR.Number
        Case Else
2850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2860    End Select
2870    Resume EXITP

End Sub

Private Sub ppd_id_lbl_DblClick(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_id_lbl_DblClick"

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

Private Sub ppd_id_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_id_KeyDown"

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
          Case vbKeyTab
3070        With Me
3080          intRetVal = 0
3090          If .journalno.Visible = True And .journalno.Enabled = True Then
3100            .journalno.SetFocus
3110          Else
3120            .Journal_ID.SetFocus
3130          End If
3140        End With
3150      End Select
3160    End If

        ' ** Shift keys.
3170    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3180      Select Case intRetVal
          Case vbKeyTab
3190        With Me
3200          intRetVal = 0
3210          If .CurrentRecord > 1 Then
3220            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3230            .FocusHolder2.SetFocus
3240            DoEvents
3250            .Username.SetFocus
3260          Else
3270            DoCmd.SelectObject acForm, .Parent.Name, False
3280            If .Parent.cmdUpdate.Enabled = True Then
3290              .Parent.cmdUpdate.SetFocus
3300            ElseIf .Parent.cmdDelete.Enabled = True Then
3310              .Parent.cmdDelete.SetFocus
3320            ElseIf .Parent.cmdEdit.Enabled = True Then
3330              .Parent.cmdEdit.SetFocus
3340            ElseIf .Parent.cmdGeneratePOSPay.Enabled = True Then
3350              .Parent.cmdGeneratePOSPay.SetFocus
3360            Else
3370              lngRecsCur = .Parent.frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
3380              If lngRecsCur > 0& Then
3390                .Parent.frmCheckPOSPay_Sub1.SetFocus
3400              Else
3410                .Parent.cmdClose.SetFocus
3420              End If
3430            End If
3440          End If
3450        End With
3460      End Select
3470    End If

EXITP:
3480    KeyCode = intRetVal
3490    Exit Sub

ERRH:
3500    intRetVal = 0
3510    Select Case ERR.Number
        Case Else
3520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3530    End Select
3540    Resume EXITP

End Sub

Private Sub journalno_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "journalno_lbl_DblClick"

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

Private Sub Journal_ID_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Journal_ID_lbl_DblClick"

3710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub ppd_checknum_lbl_DblClick(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_checknum_lbl_DblClick"

3810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub ppd_void_lbl_DblClick(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_void_lbl_DblClick"

3910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub ppd_void_Enter()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_void_Enter"

4010    With Me
4020      varOldValue = .ppd_void
4030    End With

EXITP:
4040    Exit Sub

ERRH:
4050    Select Case ERR.Number
        Case Else
4060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4070    End Select
4080    Resume EXITP

End Sub

Private Sub ppd_void_AfterUpdate()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_void_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngPPID As Long, lngPPDID As Long, lngPPJID As Long, lngJrnlID As Long, lngJNo As Long
        Dim lngRecs As Long, lngRecsChanged As Long, lngCurChecks As Long
        Dim strPath As String, strFile As String, strPathFile As String
        Dim strDocName As String, strSort As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnFound1 As Boolean, blnFound2 As Boolean, blnContinue As Boolean
        Dim lngX As Long

4110    With Me

4120      cmdSave_Click  ' ** Procedure: Above.
4130      DoEvents

4140      blnContinue = False

4150      If .ppd_void = True And varOldValue = False Then

4160        Beep
4170        msgResponse = MsgBox("Are you sure you want to void this check?", vbQuestion + vbYesNo, "Void Check")
4180        If msgResponse = vbYes Then

4190          DoCmd.Hourglass True
4200          DoEvents

4210          lngPPDID = .ppd_id
4220          lngPPID = .pp_id
4230          lngPPJID = 0&: lngJrnlID = 0&: lngJNo = 0&
4240          blnFound1 = False: blnFound2 = False

4250          lngRecsCur = .Parent.frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
4260          If lngRecsCur > 0& Then
4270            Set rst = .Parent.frmCheckPOSPay_Sub1.Form.RecordsetClone
4280            With rst
4290              .MoveFirst
4300              .FindFirst "[ppd_id] = " & CStr(lngPPDID)
4310              If .NoMatch = False Then
4320                blnFound1 = True
4330                lngPPJID = ![ppj_id]
4340                lngJrnlID = ![Journal_ID]
4350              End If
4360              .Close
4370            End With
4380            Set rst = Nothing
4390          End If  ' ** lngRecsCur.

4400          Select Case blnFound1
              Case True
4410            .Parent.frmCheckPOSPay_Sub1.Form.MoveRec 0, lngPPJID  ' ** Form Procedure: frmCheckPOSPay_Sub1.
4420            DoEvents
4430            DoCmd.Hourglass False
4440            Beep
4450            msgResponse = MsgBox("This transaction has not been posted." & vbCrLf & _
                  "You may simply delete it from the Journal." & vbCrLf & vbCrLf & _
                  "Do you wish to do that now?", vbQuestion + vbYesNo, "Delete Journal Entry")
4460            DoCmd.Hourglass True
4470            DoEvents
4480            Select Case msgResponse
                Case vbYes
                  ' ** Delete from Journal, tblCheckPOSPay_Journal,
                  ' ** tblCheckPOSPay_Detail, and adjust count in tblCheckPOSPay.
4490              Set dbs = CurrentDb
                  ' ** Delete Journal, by specified [jid].
4500              Set qdf = dbs.QueryDefs("qryCheckPOSPay_18_01")
4510              With qdf.Parameters
4520                ![jid] = lngJrnlID
4530              End With
4540              qdf.Execute
4550              Set qdf = Nothing
4560              DoEvents
                  ' ** Delete tblCheckPOSPay_Journal, by specified [ppjid].
4570              Set qdf = dbs.QueryDefs("qryCheckPOSPay_18_02")
4580              With qdf.Parameters
4590                ![ppjid] = lngPPJID
4600              End With
4610              qdf.Execute
4620              Set qdf = Nothing
4630              DoEvents
                  ' ** Delete tblCheckPOSPay_Detail, by specified [ppdid].
4640              Set qdf = dbs.QueryDefs("qryCheckPOSPay_18_03")
4650              With qdf.Parameters
4660                ![ppdid] = lngPPDID
4670              End With
4680              qdf.Execute
4690              Set qdf = Nothing
4700              DoEvents
                  ' ** Update qryCheckPOSPay_18_04 (tblCheckPOSPay, with pp_checks_new, by specified [ppid]).
4710              Set qdf = dbs.QueryDefs("qryCheckPOSPay_18_05")
4720              With qdf.Parameters
4730                ![ppid] = lngPPID
4740              End With
4750              qdf.Execute
4760              Set qdf = Nothing
4770              DoEvents
4780              .Requery
4790              .Parent.frmCheckPOSPay_Sub2.Form.Refresh
4800              .Parent.frmCheckPOSPay_Sub1.Form.Requery
4810              DoEvents
4820              lngRecsCur = .Parent.frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Funtion: frmCheckPOSPay_Sub1.
4830              .Parent.CurChecks_lbl.Caption = lngRecsCur
4840              Beep
4850              DoCmd.Hourglass False
4860            Case Else
4870              .Parent.frmCheckPOSPay_Sub1.Form.ppj_void = True
                  '.Parent.frmCheckPOSPay_Sub1.Form.ppj_changed = True
4880              .Parent.frmCheckPOSPay_Sub1.Form.cmdSave_Click  ' ** Form Procedure: frmCheckPOSPay_Sub1.
                  ' ** Show Void in the Journal subform.
4890              .Parent.frmCheckPOSPay_Sub1.Form.ShowVoid True  ' ** Form Procedure: frmCheckPOSPay_Sub1.
4900              DoEvents
4910              .Parent.frmCheckPOSPay_Sub1.Form.Refresh
4920              DoCmd.Hourglass False
4930            End Select
4940          Case False
                ' ** Let's see if we can get it from the Ledger.
4950            If IsNull(.journalno) = False Then
4960              lngJNo = .journalno
4970              Set dbs = CurrentDb
4980              Set rst = dbs.OpenRecordset("Ledger", dbOpenDynaset, dbConsistent)
4990              rst.MoveFirst
5000              rst.FindFirst "[journalno] = " & CStr(lngJNo)
5010              If rst.NoMatch = False Then
                    ' ** Record found.
5020                blnFound2 = True
5030                rst.Close
5040                Set rst = Nothing
5050              Else
5060                rst.Close
5070                Set rst = Nothing
5080                Set rst = dbs.OpenRecordset("LedgerArchive", dbOpenDynaset, dbConsistent)
5090                If rst.BOF = True And rst.EOF = True Then
                      ' ** No LedgerArchive!
5100                  rst.Close
5110                  Set rst = Nothing
5120                Else
5130                  rst.MoveFirst
5140                  rst.FindFirst "[journalno] = " & CStr(lngJNo)
5150                  If rst.NoMatch = False Then
                        ' ** Record found in archive.
5160                    blnFound2 = True
5170                    rst.Close
5180                    Set rst = Nothing
5190                  Else
                        ' ** It's not in the Journal and it's not in the Ledger!
5200                    rst.Close
5210                    Set rst = Nothing
5220                  End If
5230                End If
5240              End If
5250              dbs.Close
5260              Set dbs = Nothing
5270            Else
                  ' ** It's not in the Journal, and it doesn't seem to have been posted!
5280            End If
5290          End Select  ' ** blnFound1.

5300          If blnFound2 = True Then
                ' ** Let's not do anything else but create the new file.
                ' ** They can note in the record's comments and create the reversing enty themselves. (Unless we add that functionality.)
5310            DoCmd.Hourglass False
5320            Beep
5330            msgResponse = MsgBox("Do you wish to create a revised POSPay file now?", _
                  vbQuestion + vbYesNo, "Generate New POSPay File")
5340            If msgResponse = vbYes Then
5350              blnContinue = True
5360            Else
                  ' ** We're done here.
5370            End If  ' ** msgResponse.
5380          End If  ' ** blnFound2.

5390          If blnFound1 = False And blnFound2 = False Then
5400            DoCmd.Hourglass False
5410            Beep
5420            msgResponse = MsgBox("Trust Accountant was unable to find the original transaction." & vbCrLf & _
                  "Do you wish to generate a revised POSPay file anyway?", vbQuestion + vbYesNo, "Record Not Found")
5430            If msgResponse = vbYes Then
5440              blnContinue = True
5450            End If  ' ** msgResponse.
5460          End If  ' ** blnFound1, blnFound2.

5470        Else
5480          .ppd_void = False
5490          cmdSave_Click  ' ** Procedure: Above.
5500        End If  ' ** msgResponse.

5510      ElseIf .ppd_void = False Then 'And varOldValue = True Then  ' ** They may un-void without leaving the record.
            ' ** Void canceled.

5520        DoCmd.Hourglass True
5530        DoEvents

5540        lngPPDID = .ppd_id
5550        lngPPID = .pp_id
5560        lngPPJID = 0&: lngJrnlID = 0&: lngJNo = 0&

5570        cmdSave_Click  ' ** Procedure: Above.
5580        DoEvents

5590        lngRecsCur = .Parent.frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
5600        blnFound1 = False
5610        If lngRecsCur > 0& Then
5620          Set rst = .Parent.frmCheckPOSPay_Sub1.Form.RecordsetClone
5630          With rst
5640            .MoveFirst
5650            If ![ppd_id] = lngPPDID Then
5660              blnFound1 = True
5670              .Edit
5680              ![ppj_void] = False
                  '![ppj_changed] = True
5690              ![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
5700              ![ppj_datemodified] = Now()
5710              .Update
5720              DoEvents
5730            Else
5740              .FindFirst "[ppd_id] = " & CStr(lngPPDID)
5750              If .NoMatch = False Then
5760                blnFound1 = True
5770                .Edit
5780                ![ppj_void] = False
                    '![ppj_changed] = True
5790                ![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
5800                ![ppj_datemodified] = Now()
5810                .Update
5820                DoEvents
5830              Else
                    ' ** If it's in the Ledger, nothing else to do. Just offer updated POSPay file.
5840                DoCmd.Hourglass False
5850                Beep
5860                msgResponse = MsgBox("Would you like to regenerate the POSPay file now?", _
                      vbQuestion + vbYesNo, "Regenerate POSPay File")
5870                If msgResponse = vbYes Then
5880                  blnContinue = True
5890                  DoCmd.Hourglass True
5900                  DoEvents
5910                End If
5920              End If
5930            End If
5940            .Close
5950          End With  ' ** rst.
5960          Set rst = Nothing
5970          DoEvents

5980          If blnFound1 = True Then
                ' ** If there are no more Void's, turn off ppj_void.
5990            .Parent.frmCheckPOSPay_Sub1.Form.Refresh
6000            blnFound2 = False
6010            Set rst = .Parent.frmCheckPOSPay_Sub1.Form.RecordsetClone
6020            With rst
6030              .MoveLast
6040              lngRecs = .RecordCount
6050              .MoveFirst
6060              For lngX = 1& To lngRecs
6070                If ![ppj_void] = True Then
6080                  blnFound2 = True
6090                  Exit For
6100                End If
6110                If lngX < lngRecs Then .MoveNext
6120              Next  ' ** lngX.
6130              .Close
6140            End With  ' ** rst.
6150            Set rst = Nothing
6160            DoEvents
6170            If blnFound2 = False Then
                  ' ** No more Void's, so turn it off.
6180              .Parent.frmCheckPOSPay_Sub1.Form.ShowVoid False  ' ** Form Procedure: frmCheckPOSPay_Sub1.
6190            End If
6200            DoEvents
6210            DoCmd.Hourglass False
6220          End If  ' ** blnFound1.

6230        Else
              ' ** If it's in the Ledger, nothing else to do. Just offer updated POSPay file.
6240          DoCmd.Hourglass False
6250          Beep
6260          msgResponse = MsgBox("Would you like to regenerate the POSPay file now?", _
                vbQuestion + vbYesNo, "Regenerate POSPay File")
6270          If msgResponse = vbYes Then
6280            blnContinue = True
6290            DoCmd.Hourglass True
6300            DoEvents
6310          End If
6320        End If  ' ** lngRecsCur.

6330      End If  ' ** ppj_void, varOldValue.

6340      lngRecsChanged = .Parent.HasChanged  ' ** Form Function: frmCheckPOSPay.
6350      If lngRecsChanged > 0& And .Parent.cmdUpdate.Enabled = False Then
6360        .Parent.cmdUpdate.Enabled = True
6370      End If
6380      DoEvents

6390      If blnContinue = True Then
            ' ** Regenerate the POSPay file.

6400        DoCmd.Hourglass True
6410        DoEvents

6420        Set dbs = CurrentDb
6430        gstrFormQuerySpec = .Parent.Name

            ' ** tblCheckPOSPay, linked to tblCheckPOSPay_Detail, as
            ' ** POSPay file format, in Courier font, by specified FormRef("ppid").
6440        Set qdf = dbs.QueryDefs("qryCheckPOSPay_19_01")
6450        Set rst = qdf.OpenRecordset
6460        If rst.BOF = True And rst.EOF = True Then
              ' ** Something bad happened!
6470          lngCurChecks = 0&
6480          rst.Close
6490          Set rst = Nothing
6500          Set qdf = Nothing
6510          DoCmd.Hourglass False
6520          Beep
6530          MsgBox "A problem as occurred and this process cannot continue.", vbExclamation + vbOKOnly, "Unknown Error"
6540        Else
6550          rst.MoveLast
6560          lngCurChecks = rst.RecordCount
6570          rst.Close
6580          Set rst = Nothing
6590          Set qdf = Nothing
6600        End If

6610        dbs.Close
6620        Set dbs = Nothing

6630        If lngCurChecks > 0& Then

6640          If IsNull(.Parent.UserReportPath) = True Then
6650            strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
6660          Else
6670            strPath = .Parent.UserReportPath
6680          End If
6690          gdatEndDate = .ppd_issuedate
6700          strFile = "POSPay" & Format(gdatEndDate, "mmddyy") & ".txt"

6710          DoCmd.Hourglass False
6720          strPathFile = FileSaveDialog("txt", strFile, strPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

6730          If strPathFile <> vbNullString Then

6740            DoCmd.Hourglass True
6750            DoEvents

6760            strPath = Parse_Path(strPathFile)  ' ** Module Function: modFileUtilities.
6770            strFile = Parse_File(strPathFile)  ' ** Module Function: modFileUtilities.
6780            .Parent.UserReportPath = strPath
6790            gstrFormQuerySpec = .Parent.Name

                ' ** Generate the POSPay file using the export spec I created.
6800            DoCmd.TransferText acExportFixed, "Umpqua_Bank_Export_Specification", "qryCheckPOSPay_19_01", strPathFile, False
6810            DoEvents

6820            DoCmd.Hourglass False
6830            Beep
6840            msgResponse = MsgBox("The POSPay file has been created." & vbCrLf & vbCrLf & _
                  "Do you wish to see it?", vbQuestion + vbYesNo + vbDefaultButton2, "Export Successful")

6850            If msgResponse = vbYes Then
6860              gstrFormQuerySpec = .Parent.Name
6870              strDocName = "rptPOSPay_Export_02"
6880              strSort = SortNow_Get  ' ** Function: Below.
6890              DoCmd.OpenReport strDocName, acViewPreview, , , , .Parent.Name & "~" & strSort
6900              DoCmd.Maximize
6910              DoCmd.RunCommand acCmdFitToWindow
6920            End If

6930          End If  ' ** strPathFile.
6940        End If  ' ** lngCurChecks.
6950      End If  ' ** blnContinue.

6960    End With

EXITP:
6970    Set rst = Nothing
6980    Set qdf = Nothing
6990    Set dbs = Nothing
7000    Exit Sub

ERRH:
7010    Select Case ERR.Number
        Case Else
7020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7030    End Select
7040    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

7110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7120    Exit Sub

ERRH:
7130    Select Case ERR.Number
        Case Else
7140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7150    End Select
7160    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

7210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7220    Exit Sub

ERRH:
7230    Select Case ERR.Number
        Case Else
7240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7250    End Select
7260    Resume EXITP

End Sub

Private Sub ppd_issuedate_lbl_DblClick(Cancel As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_issuedate_lbl_DblClick"

7310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7320    Exit Sub

ERRH:
7330    Select Case ERR.Number
        Case Else
7340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7350    End Select
7360    Resume EXITP

End Sub

Private Sub ppd_amount_lbl_DblClick(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_amount_lbl_DblClick"

7410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7420    Exit Sub

ERRH:
7430    Select Case ERR.Number
        Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

7510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7520    Exit Sub

ERRH:
7530    Select Case ERR.Number
        Case Else
7540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7550    End Select
7560    Resume EXITP

End Sub

Private Sub ppd_payee_lbl_DblClick(Cancel As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_payee_lbl_DblClick"

7610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7620    Exit Sub

ERRH:
7630    Select Case ERR.Number
        Case Else
7640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7650    End Select
7660    Resume EXITP

End Sub

Private Sub ppd_bank_name_lbl_DblClick(Cancel As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_bank_name_lbl_DblClick"

7710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7720    Exit Sub

ERRH:
7730    Select Case ERR.Number
        Case Else
7740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7750    End Select
7760    Resume EXITP

End Sub

Private Sub ppd_aba_trc_lbl_DblClick(Cancel As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_aba_trc_lbl_DblClick"

7810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7820    Exit Sub

ERRH:
7830    Select Case ERR.Number
        Case Else
7840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7850    End Select
7860    Resume EXITP

End Sub

Private Sub ppd_bank_account_lbl_DblClick(Cancel As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "ppd_bank_account_lbl_DblClick"

7910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7920    Exit Sub

ERRH:
7930    Select Case ERR.Number
        Case Else
7940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7950    End Select
7960    Resume EXITP

End Sub

Private Sub Username_lbl_DblClick(Cancel As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "Username_lbl_DblClick"

8010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8020    Exit Sub

ERRH:
8030    Select Case ERR.Number
        Case Else
8040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8050    End Select
8060    Resume EXITP

End Sub

Private Sub Username_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "Username_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8120    intShiftDown = (Shift And acShiftMask) > 0
8130    intAltDown = (Shift And acAltMask) > 0
8140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8160      Select Case intRetVal
          Case vbKeyTab
8170        With Me
8180          intRetVal = 0
8190          lngRecsCur = RecCnt  ' ** Function: Below.
8200          If .CurrentRecord < lngRecsCur Then
8210            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
8220            .ppd_id.SetFocus
8230          Else
8240            DoCmd.SelectObject acForm, .Parent.Name, False
8250            If .Parent.cmdRegenerate.Enabled = True Then
8260              .Parent.cmdRegenerate.SetFocus
8270            ElseIf .Parent.cmdDeletePosPay.Enabled = True Then
8280              .Parent.cmdDeletePosPay.SetFocus
8290            Else
8300              .Parent.cmdClose.SetFocus
8310            End If
8320          End If
8330        End With
8340      End Select
8350    End If

        ' ** Shift keys.
8360    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8370      Select Case intRetVal
          Case vbKeyTab
8380        With Me
8390          intRetVal = 0
8400          .ppd_bank_account.SetFocus
8410        End With
8420      End Select
8430    End If

EXITP:
8440    KeyCode = intRetVal
8450    Exit Sub

ERRH:
8460    intRetVal = 0
8470    Select Case ERR.Number
        Case Else
8480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8490    End Select
8500    Resume EXITP

End Sub

Public Function VoidCount() As Long

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "VoidCount"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long
        Dim lngRetVal As Long

8610    With Me
8620      lngRetVal = 0&
8630      lngRecsCur = RecCnt  ' ** Function: Below.
8640      If lngRecsCur > 0& Then
8650        Set rst = .RecordsetClone
8660        With rst
8670          .MoveLast
8680          lngRecs = .RecordCount
8690          .MoveFirst
8700          For lngX = 1& To lngRecs
8710            If ![ppd_void] = True Then
8720              lngRetVal = lngRetVal + 1&
8730            End If
8740            If lngX < lngRecs Then .MoveNext
8750          Next  ' ** lngX.
8760          .Close
8770        End With  ' ** rst.
8780        Set rst = Nothing
8790      End If
8800    End With

EXITP:
8810    Set rst = Nothing
8820    VoidCount = lngRetVal
8830    Exit Function

ERRH:
8840    lngRetVal = 0&
8850    Select Case ERR.Number
        Case Else
8860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8870    End Select
8880    Resume EXITP

End Function

Public Sub Sub3_Disable(blnAble As Boolean)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "Sub3_Disable"

8910    With Me
8920      Select Case blnAble
          Case True
8930        .ppd_id_lbl.ForeColor = CLR_DKGRY2
8940        .ppd_id_lbl_dim_hi.Visible = False
8950        .ppd_id_lbl_line.BorderColor = CLR_DKGRY
8960        .ppd_id_lbl_line_dim_hi.Visible = False
8970        .journalno_lbl.ForeColor = CLR_DKGRY2
8980        .journalno_lbl_dim_hi.Visible = False
8990        .journalno_lbl2.ForeColor = CLR_DKGRY2
9000        .journalno_lbl2_dim_hi.Visible = False
9010        .journalno_lbl_line.BorderColor = CLR_DKGRY
9020        .journalno_lbl_line_dim_hi.Visible = False
9030        .Journal_ID_lbl.ForeColor = CLR_DKGRY2
9040        .Journal_ID_lbl_dim_hi.Visible = False
9050        .Journal_ID_lbl2.ForeColor = CLR_DKGRY2
9060        .Journal_ID_lbl2_dim_hi.Visible = False
9070        .Journal_ID_lbl_line.BorderColor = CLR_DKGRY
9080        .Journal_ID_lbl_line_dim_hi.Visible = False
9090        .ppd_checknum_lbl.ForeColor = CLR_DKGRY2
9100        .ppd_checknum_lbl_dim_hi.Visible = False
9110        .ppd_checknum_lbl2.ForeColor = CLR_DKGRY2
9120        .ppd_checknum_lbl2_dim_hi.Visible = False
9130        .ppd_checknum_lbl_line.BorderColor = CLR_DKGRY
9140        .ppd_checknum_lbl_line_dim_hi.Visible = False
9150        .ppd_void_lbl.ForeColor = CLR_DKGRY2
9160        .ppd_void_lbl_dim_hi.Visible = False
9170        .ppd_void_lbl_line.BorderColor = CLR_DKGRY
9180        .ppd_void_lbl_line_dim_hi.Visible = False
9190        .accountno_lbl.ForeColor = CLR_DKGRY2
9200        .accountno_lbl_dim_hi.Visible = False
9210        .accountno_lbl2.ForeColor = CLR_DKGRY2
9220        .accountno_lbl2_dim_hi.Visible = False
9230        .accountno_lbl_line.BorderColor = CLR_DKGRY
9240        .accountno_lbl_line_dim_hi.Visible = False
9250        .shortname_lbl.ForeColor = CLR_DKGRY2
9260        .shortname_lbl_dim_hi.Visible = False
9270        .shortname_lbl_line.BorderColor = CLR_DKGRY
9280        .shortname_lbl_Line_dim_hi.Visible = False
9290        .ppd_issuedate_lbl.ForeColor = CLR_DKGRY2
9300        .ppd_issuedate_lbl_dim_hi.Visible = False
9310        .ppd_issuedate_lbl2.ForeColor = CLR_DKGRY2
9320        .ppd_issuedate_lbl2_dim_hi.Visible = False
9330        .ppd_issuedate_lbl_line.BorderColor = CLR_DKGRY
9340        .ppd_issuedate_lbl_line_dim_hi.Visible = False
9350        .ppd_amount_lbl.ForeColor = CLR_DKGRY2
9360        .ppd_amount_lbl_dim_hi.Visible = False
9370        .ppd_amount_lbl_line.BorderColor = CLR_DKGRY
9380        .ppd_amount_lbl_line_dim_hi.Visible = False
9390        .ppd_payee_lbl.ForeColor = CLR_DKGRY2
9400        .ppd_payee_lbl_dim_hi.Visible = False
9410        .ppd_payee_lbl_line.BorderColor = CLR_DKGRY
9420        .ppd_payee_lbl_line_dim_hi.Visible = False
9430        If .curr_id.Visible = True Then
9440          .curr_id_lbl.ForeColor = CLR_DKGRY2
9450          .curr_id_lbl_dim_hi.Visible = False
9460          .curr_id_lbl_line.BorderColor = CLR_DKGRY
9470          .curr_id_lbl_line_dim_hi.Visible = False
9480        End If
9490        .ppd_bank_name_lbl.ForeColor = CLR_DKGRY2
9500        .ppd_bank_name_lbl_dim_hi.Visible = False
9510        .ppd_bank_name_lbl_line.BorderColor = CLR_DKGRY
9520        .ppd_bank_name_lbl_line_dim_hi.Visible = False
9530        .ppd_aba_trc_lbl.ForeColor = CLR_DKGRY2
9540        .ppd_aba_trc_lbl_dim_hi.Visible = False
9550        .ppd_aba_trc_lbl2.ForeColor = CLR_DKGRY2
9560        .ppd_aba_trc_lbl2_dim_hi.Visible = False
9570        .ppd_aba_trc_lbl_line.BorderColor = CLR_DKGRY
9580        .ppd_aba_trc_lbl_line_dim_hi.Visible = False
9590        .ppd_bank_account_lbl.ForeColor = CLR_DKGRY2
9600        .ppd_bank_account_lbl_dim_hi.Visible = False
9610        .ppd_bank_account_lbl2.ForeColor = CLR_DKGRY2
9620        .ppd_bank_account_lbl2_dim_hi.Visible = False
9630        .ppd_bank_account_lbl_line.BorderColor = CLR_DKGRY
9640        .ppd_bank_account_lbl_line_dim_hi.Visible = False
9650        .Username_lbl.ForeColor = CLR_DKGRY2
9660        .Username_lbl_dim_hi.Visible = False
9670        .Username_lbl_line.BorderColor = CLR_DKGRY
9680        .Username_lbl_line_dim_hi.Visible = False
9690      Case False
9700        .ppd_id_lbl.ForeColor = WIN_CLR_DISF
9710        .ppd_id_lbl_dim_hi.Visible = True
9720        .ppd_id_lbl_line.BorderColor = WIN_CLR_DISR
9730        .ppd_id_lbl_line_dim_hi.Visible = True
9740        .journalno_lbl.ForeColor = WIN_CLR_DISF
9750        .journalno_lbl_dim_hi.Visible = True
9760        .journalno_lbl2.ForeColor = WIN_CLR_DISF
9770        .journalno_lbl2_dim_hi.Visible = True
9780        .journalno_lbl_line.BorderColor = WIN_CLR_DISR
9790        .journalno_lbl_line_dim_hi.Visible = True
9800        .Journal_ID_lbl.ForeColor = WIN_CLR_DISF
9810        .Journal_ID_lbl_dim_hi.Visible = True
9820        .Journal_ID_lbl2.ForeColor = WIN_CLR_DISF
9830        .Journal_ID_lbl2_dim_hi.Visible = True
9840        .Journal_ID_lbl_line.BorderColor = WIN_CLR_DISR
9850        .Journal_ID_lbl_line_dim_hi.Visible = True
9860        .ppd_checknum_lbl.ForeColor = WIN_CLR_DISF
9870        .ppd_checknum_lbl_dim_hi.Visible = True
9880        .ppd_checknum_lbl2.ForeColor = WIN_CLR_DISF
9890        .ppd_checknum_lbl2_dim_hi.Visible = True
9900        .ppd_checknum_lbl_line.BorderColor = WIN_CLR_DISR
9910        .ppd_checknum_lbl_line_dim_hi.Visible = True
9920        .ppd_void_lbl.ForeColor = WIN_CLR_DISF
9930        .ppd_void_lbl_dim_hi.Visible = True
9940        .ppd_void_lbl_line.BorderColor = WIN_CLR_DISR
9950        .ppd_void_lbl_line_dim_hi.Visible = True
9960        .accountno_lbl.ForeColor = WIN_CLR_DISF
9970        .accountno_lbl_dim_hi.Visible = True
9980        .accountno_lbl2.ForeColor = WIN_CLR_DISF
9990        .accountno_lbl2_dim_hi.Visible = True
10000       .accountno_lbl_line.BorderColor = WIN_CLR_DISR
10010       .accountno_lbl_line_dim_hi.Visible = True
10020       .shortname_lbl.ForeColor = WIN_CLR_DISF
10030       .shortname_lbl_dim_hi.Visible = True
10040       .shortname_lbl_line.BorderColor = WIN_CLR_DISR
10050       .shortname_lbl_Line_dim_hi.Visible = True
10060       .ppd_issuedate_lbl.ForeColor = WIN_CLR_DISF
10070       .ppd_issuedate_lbl_dim_hi.Visible = True
10080       .ppd_issuedate_lbl2.ForeColor = WIN_CLR_DISF
10090       .ppd_issuedate_lbl2_dim_hi.Visible = True
10100       .ppd_issuedate_lbl_line.BorderColor = WIN_CLR_DISR
10110       .ppd_issuedate_lbl_line_dim_hi.Visible = True
10120       .ppd_amount_lbl.ForeColor = WIN_CLR_DISF
10130       .ppd_amount_lbl_dim_hi.Visible = True
10140       .ppd_amount_lbl_line.BorderColor = WIN_CLR_DISR
10150       .ppd_amount_lbl_line_dim_hi.Visible = True
10160       .ppd_payee_lbl.ForeColor = WIN_CLR_DISF
10170       .ppd_payee_lbl_dim_hi.Visible = True
10180       .ppd_payee_lbl_line.BorderColor = WIN_CLR_DISR
10190       .ppd_payee_lbl_line_dim_hi.Visible = True
10200       If .curr_id.Visible = True Then
10210         .curr_id_lbl.ForeColor = WIN_CLR_DISF
10220         .curr_id_lbl_dim_hi.Visible = True
10230         .curr_id_lbl_line.BorderColor = WIN_CLR_DISR
10240         .curr_id_lbl_line_dim_hi.Visible = True
10250       End If
10260       .ppd_bank_name_lbl.ForeColor = WIN_CLR_DISF
10270       .ppd_bank_name_lbl_dim_hi.Visible = True
10280       .ppd_bank_name_lbl_line.BorderColor = WIN_CLR_DISR
10290       .ppd_bank_name_lbl_line_dim_hi.Visible = True
10300       .ppd_aba_trc_lbl.ForeColor = WIN_CLR_DISF
10310       .ppd_aba_trc_lbl_dim_hi.Visible = True
10320       .ppd_aba_trc_lbl2.ForeColor = WIN_CLR_DISF
10330       .ppd_aba_trc_lbl2_dim_hi.Visible = True
10340       .ppd_aba_trc_lbl_line.BorderColor = WIN_CLR_DISR
10350       .ppd_aba_trc_lbl_line_dim_hi.Visible = True
10360       .ppd_bank_account_lbl.ForeColor = WIN_CLR_DISF
10370       .ppd_bank_account_lbl_dim_hi.Visible = True
10380       .ppd_bank_account_lbl2.ForeColor = WIN_CLR_DISF
10390       .ppd_bank_account_lbl2_dim_hi.Visible = True
10400       .ppd_bank_account_lbl_line.BorderColor = WIN_CLR_DISR
10410       .ppd_bank_account_lbl_line_dim_hi.Visible = True
10420       .Username_lbl.ForeColor = WIN_CLR_DISF
10430       .Username_lbl_dim_hi.Visible = True
10440       .Username_lbl_line.BorderColor = WIN_CLR_DISR
10450       .Username_lbl_line_dim_hi.Visible = True
10460     End Select
10470   End With

EXITP:
10480   Exit Sub

ERRH:
10490   Select Case ERR.Number
        Case Else
10500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10510   End Select
10520   Resume EXITP

End Sub

Public Function RecCnt() As Long

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

10610   lngRetVal = 0&

10620   With Me
10630     Set rst = .RecordsetClone
10640     With rst
10650       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
10660       Else
10670         .MoveLast
10680         lngRetVal = .RecordCount
10690       End If
10700       .Close
10710     End With
10720   End With

EXITP:
10730   Set rst = Nothing
10740   RecCnt = lngRetVal
10750   Exit Function

ERRH:
10760   lngRetVal = 0&
10770   Select Case ERR.Number
        Case Else
10780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10790   End Select
10800   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

10910   Select Case intWhere
        Case acCmdRecordsGoToFirst
10920     DoCmd.RunCommand acCmdRecordsGoToFirst
10930   Case acCmdRecordsGoToLast
10940     DoCmd.RunCommand acCmdRecordsGoToLast
10950   Case acCmdRecordsGoToPrevious
10960     DoCmd.RunCommand acCmdRecordsGoToPrevious
10970   Case acCmdRecordsGoToNext
10980     DoCmd.RunCommand acCmdRecordsGoToNext
10990   Case Else
11000     If IsMissing(varID) = False Then
11010       With Me
11020         Set rst = .RecordsetClone
11030         With rst
11040           .FindFirst "[ppd_id] = " & CStr(varID)
11050           If .NoMatch = False Then
11060             Me.Bookmark = .Bookmark
11070           End If
11080           .Close
11090         End With
11100       End With
11110     End If
11120   End Select

EXITP:
11130   Set rst = Nothing
11140   Exit Sub

ERRH:
11150   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
11160   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
11170   Case Else
11180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11190   End Select
11200   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String

        Const strStdAsc As String = ", [accountno], [ppd_payee]"

11310   With Me
11320     If .ppd_id_lbl_dim_hi.Visible = False Then
11330       If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
11340         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
11350       End If
11360       .Controls(strSortLbl).Visible = False
11370       .Controls(strSortLine).Visible = False
11380       .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
11390       If strProc = "Form_Load" Then
11400         strCalled = "Journal_ID"
11410         strSortNow = strSortOrig
11420         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
11430         lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - (lngSortLbl_Width)) + (8& * lngTpp))
11440         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
11450         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
11460         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
11470         .Controls(strSortLbl).Top = lngSortLbl_Top
11480         .Controls(strSortLbl).Left = lngSortLbl_Left
11490         .Controls(strSortLbl).Caption = strArwUp
11500         .Controls(strSortLine).Top = lngSortLine_Top
11510         .Controls(strSortLine).Left = lngSortLine_Left
11520         .Controls(strSortLine).Width = lngSortLine_Width
11530         .Controls(strSortLbl).ForeColor = CLR_DKBLU
11540       Else
11550         strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
11560         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
11570         lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - (lngSortLbl_Width))
11580         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
11590         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
11600         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
11610         .Controls(strSortLbl).Top = lngSortLbl_Top
11620         .Controls(strSortLbl).Left = lngSortLbl_Left
11630         .Controls(strSortLine).Top = lngSortLine_Top
11640         .Controls(strSortLine).Left = lngSortLine_Left
11650         .Controls(strSortLine).Width = lngSortLine_Width
11660         If strCalled = "Journal_ID" Then
11670           lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
11680           .Controls(strSortLbl).Left = lngSortLbl_Left
11690           strSortAsc = strSortOrig
11700           If strSortNow = strSortAsc Then
11710             strSortNow = "[Journal_ID] DESC"
11720             .Controls(strSortLbl).Caption = strArwDn
11730             .Controls(strSortLbl).ForeColor = CLR_DKRED
11740           Else
11750             strSortNow = strSortAsc
11760             .Controls(strSortLbl).Caption = strArwUp
11770             .Controls(strSortLbl).ForeColor = CLR_DKBLU
11780           End If
11790         Else
11800           .Controls(strSortLbl).Caption = strArwUp
11810           .Controls(strSortLbl).ForeColor = CLR_DKBLU
11820           Select Case strCalled
                Case "ppd_id"
11830             strSortAsc = "[ppd_id]"
11840           Case "journalno"
11850             strSortAsc = "[journalno]"
11860           Case "ppd_checknum"
11870             strSortAsc = "[ppd_checknum]" & strStdAsc
11880           Case "ppd_void"
11890             strSortAsc = "[ppd_void]" & strStdAsc
11900             lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
11910             lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
11920             .Controls(strSortLine).Left = lngSortLine_Left
11930             .Controls(strSortLine).Width = lngSortLine_Width
11940             lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
11950             .Controls(strSortLbl).Left = lngSortLbl_Left
11960           Case "accountno"
11970             strSortAsc = "[accountno], [ppd_issuedate], [ppd_payee]"
11980           Case "shortname"
11990             strSortAsc = "[shortname], [ppd_issuedate], [ppd_payee]"
12000           Case "ppd_issuedate"
12010             strSortAsc = "[ppd_issuedate]" & strStdAsc
12020           Case "ppd_amount"
12030             strSortAsc = "[ppd_amount]" & strStdAsc
12040           Case "curr_id"
12050             strSortAsc = "[curr_code]" & strStdAsc
12060             lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
12070             lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
12080             .Controls(strSortLine).Left = lngSortLine_Left
12090             .Controls(strSortLine).Width = lngSortLine_Width
12100           Case "ppd_payee"
12110             strSortAsc = "[ppd_payee]" & strStdAsc
12120           Case "ppd_bank_name"
12130             strSortAsc = "[ppd_bank_name]" & strStdAsc
12140           Case "ppd_aba_trc"
12150             strSortAsc = "[ppd_aba_trc]" & strStdAsc
12160             lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
12170             .Controls(strSortLbl).Left = lngSortLbl_Left
12180           Case "ppd_bank_account"
12190             strSortAsc = "[ppd_bank_account]" & strStdAsc
12200             lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
12210             .Controls(strSortLbl).Left = lngSortLbl_Left
12220           Case "Username"
12230             strSortAsc = "[Username]" & strStdAsc
12240             lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
12250             .Controls(strSortLbl).Left = lngSortLbl_Left
12260           End Select
12270           If strSortNow = strSortAsc Then
12280             intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
12290             Select Case intCnt
                  Case 1
12300               strTmp01 = strSortAsc & " DESC"
12310             Case 2
12320               intPos01 = InStr(strSortAsc, ",")
12330               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
12340               strTmp02 = Mid(strSortAsc, intPos01) & " DESC"
12350               strTmp01 = strTmp01 & strTmp02
12360             Case 3
12370               intPos01 = InStr(strSortAsc, ",")
12380               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
12390               strTmp03 = Mid(strSortAsc, intPos01)
12400               intPos01 = InStr(2, strTmp03, ",")
12410               strTmp02 = Left(strTmp03, (intPos01 - 1)) & " DESC"
12420               strTmp03 = Mid(strTmp03, intPos01) & " DESC"
12430               strTmp01 = strTmp01 & strTmp02 & strTmp03
12440             Case 4
12450               intPos01 = InStr(strSortAsc, ",")
12460               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
12470               strTmp03 = Mid(strSortAsc, intPos01)
12480               intPos01 = InStr(2, strTmp03, ",")
12490               strTmp02 = Left(strTmp03, (intPos01 - 1)) & " DESC"
12500               strTmp01 = strTmp01 & strTmp02
12510               strTmp03 = Mid(strTmp03, intPos01)
12520               intPos01 = InStr(2, strTmp03, ",")
12530               strTmp02 = Left(strTmp03, (intPos01 - 1)) & " DESC"
12540               strTmp03 = Mid(strTmp03, intPos01) & " DESC"
12550               strTmp01 = strTmp01 & strTmp02 & strTmp03
12560             End Select
12570             strSortNow = strTmp01
12580             .Controls(strSortLbl).Caption = strArwDn
12590             .Controls(strSortLbl).ForeColor = CLR_DKRED
12600           Else
12610             strSortNow = strSortAsc
12620           End If
12630         End If
12640       End If
12650       .Controls(strSortLbl).Visible = True
12660       .Controls(strSortLine).Visible = True
12670       .OrderBy = strSortNow
12680       .OrderByOn = True
12690     End If
12700   End With

        ' ** Potential Sorts:
        ' **   [Journal_ID]
        ' **   [ppd_id]
        ' **   [journalno]
        ' **   [ppd_checknum], [accountno], [ppd_payee]
        'Check_Number
        ' **   [ppd_void], [accountno], [ppd_payee]
        'Issue_Type
        ' **   [accountno], [ppd_issuedate], [ppd_payee]
        ' **   [shortname], [ppd_issuedate], [ppd_payee]
        ' **   [ppd_issuedate], [accountno], [ppd_payee]
        'Issue_Date
        ' **   [ppd_amount], [accountno], [ppd_payee]
        'Amount
        ' **   [curr_code], [accountno], [ppd_payee]
        ' **   [ppd_payee], [accountno], [ppd_payee]
        'Payee_Name
        ' **   [ppd_bank_name], [accountno], [ppd_payee]
        ' **   [ppd_aba_trc], [accountno], [ppd_payee]
        'ABA_TRC
        ' **   [ppd_bank_account], [accountno], [ppd_payee]
        'Account
        ' **   [Username], [accountno], [ppd_payee]

        'Debit_Credit
        'Issue_Action

EXITP:
12710   Exit Sub

ERRH:
12720   Select Case ERR.Number
        Case Else
12730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12740   End Select
12750   Resume EXITP

End Sub

Public Function SortNow_Get() As String

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim blnIsDesc As Boolean
        Dim strTmp01 As String
        Dim intPos01 As Integer
        Dim strRetVal As String

        Const strStdAsc As String = ", [Account], [Payee_Name]"
        Const strStdAscD As String = ", [Account] DESC, [Payee_Name] DESC"

12810   blnIsDesc = False

12820   If strSortNow = vbNullString Then
12830     strRetVal = strSortOrig
12840   Else
12850     strRetVal = strSortNow
12860   End If

12870   If InStr(strRetVal, " DESC") > 0 Then
12880     blnIsDesc = True
12890   End If

12900   intPos01 = InStr(strRetVal, "]")
12910   strTmp01 = Left(strRetVal, intPos01)
12920   Select Case strTmp01
        Case "[ppd_checknum]"
          ' **   [ppd_checknum], [accountno], [ppd_payee]
12930     Select Case blnIsDesc
          Case True
12940       strRetVal = "[Check_Number] DESC" & strStdAscD
12950     Case False
12960       strRetVal = "[Check_Number]" & strStdAsc
12970     End Select
12980   Case "[ppd_void]"
          ' **   [ppd_void], [accountno], [ppd_payee]
12990     Select Case blnIsDesc
          Case True
13000       strRetVal = "[Issue_Type] DESC" & strStdAscD
13010     Case False
13020       strRetVal = "[Issue_Type]" & strStdAsc
13030     End Select
13040   Case "[ppd_issuedate]"
          ' **   [ppd_issuedate], [accountno], [ppd_payee]
13050     Select Case blnIsDesc
          Case True
13060       strRetVal = "[Issue_Date] DESC" & strStdAscD
13070     Case False
13080       strRetVal = "[Issue_Date]" & strStdAsc
13090     End Select
13100   Case "[ppd_amount]"
          ' **   [ppd_amount], [accountno], [ppd_payee]
13110     Select Case blnIsDesc
          Case True
13120       strRetVal = "[Amount] DESC" & strStdAscD
13130     Case False
13140       strRetVal = "[Amount]" & strStdAsc
13150     End Select
13160   Case "[ppd_payee]"
          ' **   [ppd_payee], [accountno], [ppd_payee]
13170     Select Case blnIsDesc
          Case True
13180       strRetVal = "[Payee_Name] DESC" & strStdAscD
13190     Case False
13200       strRetVal = "[Payee_Name]" & strStdAsc
13210     End Select
13220   Case "[ppd_aba_trc]"
          ' **   [ppd_aba_trc], [accountno], [ppd_payee]
13230     Select Case blnIsDesc
          Case True
13240       strRetVal = "[ABA_TRC] DESC" & strStdAscD
13250     Case False
13260       strRetVal = "[ABA_TRC]" & strStdAsc
13270     End Select
13280   Case "[ppd_bank_account]"
          ' **   [ppd_bank_account], [accountno], [ppd_payee]
13290     Select Case blnIsDesc
          Case True
13300       strRetVal = "[Account] DESC" & strStdAscD
13310     Case False
13320       strRetVal = "[Account]" & strStdAsc
13330     End Select
13340   Case Else
13350     Select Case strTmp01
          Case "[Journal_ID]", "[ppd_id]", "[journalno]"
            ' **   [Journal_ID]
            ' **   [ppd_id]
            ' **   [journalno]
13360       Select Case blnIsDesc
            Case True
13370         strRetVal = "[Issue_Date] DESC" & strStdAscD
13380       Case False
13390         strRetVal = "[Issue_Date]" & strStdAsc
13400       End Select
13410     Case "[accountno]", "[shortname]"
            ' **   [accountno], [ppd_issuedate], [ppd_payee]
            ' **   [shortname], [ppd_issuedate], [ppd_payee]
13420       Select Case blnIsDesc
            Case True
13430         strRetVal = "[Account] DESC" & strStdAscD
13440       Case False
13450         strRetVal = "[Account]" & strStdAsc
13460       End Select
13470     Case "[ppd_bank_name]"
            ' **   [ppd_bank_name], [accountno], [ppd_payee]
13480       Select Case blnIsDesc
            Case True
13490         strRetVal = "[ABA_TRC] DESC" & strStdAscD
13500       Case False
13510         strRetVal = "[ABA_TRC]" & strStdAsc
13520       End Select
13530     Case "[curr_code]"
            ' **   [curr_code], [accountno], [ppd_payee]
13540       Select Case blnIsDesc
            Case True
13550         strRetVal = "[Account] DESC, [Payee_Name] DESC"
13560       Case False
13570         strRetVal = "[Account], [Payee_Name]"
13580       End Select
13590     End Select
13600   End Select

        ' ** Fields not in subform:
        ' **   [Debit_Credit]
        ' **   [Issue_Action]

        ' ** Fields not on report:
        ' **   [Journal_ID]
        ' **   [ppd_id]
        ' **   [journalno]
        ' **   [accountno], [ppd_issuedate], [ppd_payee]
        ' **   [shortname], [ppd_issuedate], [ppd_payee]
        ' **   [curr_code], [accountno], [ppd_payee]
        ' **   [ppd_bank_name], [accountno], [ppd_payee]

EXITP:
13610   SortNow_Get = strRetVal
13620   Exit Function

ERRH:
13630   strRetVal = RET_ERR
13640   Select Case ERR.Number
        Case Else
13650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13660   End Select
13670   Resume EXITP

End Function
