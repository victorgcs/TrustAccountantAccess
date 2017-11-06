VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckPOSPay_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckPOSPay_Sub1"

'VGC 03/25/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Issue Date:       A {Issue_Date on frmCheckPOSPay}
' **   Check Number:     B {ppd_checknum on frmCheckPOSPay_Sub3}
' **   Close:            C {cmdClose on frmCheckPOSPay}
' **   Delete:           D {cmdDelete on frmCheckPOSPay}
' **   Edit:             E {cmdEdit on frmCheckPOSPay}
' **   Files:            F {opgView_optFiles on frmCheckPOSPay}
' **   Generate POSPay:  G {cmdGeneratePOSPay on frmCheckPOSPay}
' **   Checks:           H {opgView_optChecks on frmCheckPOSPay}
' **   Edit POSPay:      I {cmdEditPosPay on frmCheckPOSPay}
' **   Check Number:     K {CheckNum}
' **   Include Currency: L {chkIncludeCurrency on frmCheckPOSPay}
' **   Check Memo:       M {JrnlMemo_Memo}
' **   Regenerate:       N {cmdRegenerate on frmCheckPOSPay}
' **   Comments:         O {description}
' **   Print Report:     P {cmdPrintReport on frmCheckPOSPay}
' **   Currency:         R {curr_id}
' **   Posting Date:     S {transdate}
' **   Delete POSPay:    T {cmdDeletePosPay on frmCheckPOSPay}
' **   Update:           U {cmdUpdate on frmCheckPOSPay}
' **   Void:             V {ppd_void on frmCheckPOSPay_Sub3}
' **   Exit:             X {cmdClose on frmCheckPOSPay}
' **   Payee:            Y {RecurringItem}

'I CAN'T FIGURE OUT WHAT TO GIVE THESE!
' **   Description:      P {pp_description on frmCheckPOSPay_Sub2}  ' ** Not visible at
' **   Payee:            P {ppd_payee on frmCheckPOSPay_Sub3}       ' ** the same time.

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

' ** Array: arr_varCtl1().
Private lngCtl1s As Long, arr_varCtl1() As Variant
Private lngCtl2s As Long, arr_varCtl2() As Variant
Private Const C_ELEMS As Integer = 7  ' ** Arrays first-element UBound().
Private Const C_CNAM As Integer = 0
Private Const C_LFT  As Integer = 1
Private Const C_LBL  As Integer = 2
Private Const C_LBL2 As Integer = 3
Private Const C_LIN  As Integer = 4
Private Const C_DMLB As Integer = 5
Private Const C_DML2 As Integer = 6
Private Const C_DMLN As Integer = 7

Private lngCurrID_Left As Long, lngCurrID_Width As Long, lngCurrIDLbl_Left As Long
Private lngForm_Width As Long, lngVoid_Width As Long

Private blnChanged As Boolean, varOldValue As Variant
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

170         .form_width_line.Left = 0&  ' ** Just get it out of the way.

180         lngForm_Width = .Width  ' ** This is without curr_id showing.
190         lngCurrID_Left = .curr_id.Left
200         lngCurrID_Width = ((.curr_id.Left + .curr_id.Width) - (.amount.Left + .amount.Width))  ' ** Includes gutter.
210         lngCurrIDLbl_Left = .curr_id_lbl.Left
220         lngVoid_Width = ((.ppj_void_lbl_line.Left + (.ppj_void_lbl_line.Width - lngTpp)) - (.CheckNum.Left + .CheckNum.Width))
230         lngSortLbl_Width = .Sort_lbl.Width

240         CtlArray1_Load  ' ** Procedure: Below.
250         CtlArray2_Load  ' ** Procedure: Below.

260         blnChanged = False

270       Else
280         Cancel = -1
290       End If
300     End With

310     If Cancel = -1 Then
320       Beep
330       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
340     End If

EXITP:
350     Exit Sub

ERRH:
360     Select Case ERR.Number
        Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
380     End Select
390     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me

420       CLR_DISABLED_FG = CLR_DKGRY
430       CLR_DISABLED_BG = CLR_LTTEAL

440       .Journal_ID.ForeColor = CLR_DISABLED_FG
450       .Journal_ID.BackColor = CLR_DISABLED_BG
460       .journaltype.ForeColor = CLR_DISABLED_FG
470       .journaltype.BackColor = CLR_DISABLED_BG
480       .CheckNum.ForeColor = CLR_DISABLED_FG
490       .CheckNum.BackColor = CLR_DISABLED_BG
500       .accountno.ForeColor = CLR_DISABLED_FG
510       .accountno.BackColor = CLR_DISABLED_BG
520       .shortname.ForeColor = CLR_DISABLED_FG
530       .shortname.BackColor = CLR_DISABLED_BG
540       .transdate.ForeColor = CLR_DISABLED_FG
550       .transdate.BackColor = CLR_DISABLED_BG
560       .ICash.ForeColor = CLR_DISABLED_FG
570       .ICash.BackColor = CLR_DISABLED_BG
580       .PCash.ForeColor = CLR_DISABLED_FG
590       .PCash.BackColor = CLR_DISABLED_BG
600       .amount.ForeColor = CLR_DISABLED_FG
610       .amount.BackColor = CLR_DISABLED_BG
620       .RecurringItem.ForeColor = CLR_DISABLED_FG
630       .RecurringItem.BackColor = CLR_DISABLED_BG
640       .JrnlMemo_Memo.ForeColor = CLR_DISABLED_FG
650       .JrnlMemo_Memo.BackColor = CLR_DISABLED_BG
660       .description.ForeColor = CLR_DISABLED_FG
670       .description.BackColor = CLR_DISABLED_BG
680       .journal_USER.ForeColor = CLR_DISABLED_FG
690       .journal_USER.BackColor = CLR_DISABLED_BG
700       .curr_id.ForeColor = CLR_DISABLED_FG
710       .curr_id.BackColor = CLR_DISABLED_BG

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

        Dim strAccountNo As String
        Dim varTmp00 As Variant

810     With Me

820       blnChanged = .ppj_changed

830       If gblnHasForEx = True Then
840         strAccountNo = .accountno
            ' ** How will this change things?
850         gblnHasForExThis = HasForEx_Acct(strAccountNo)  ' ** Module Function: modCurrencyFuncs.
860       End If

870       Select Case .ppj_pospay
          Case True
880         .Parent.FocusHolder.SetFocus
890         DoEvents
900         .Parent.cmdGeneratePOSPay.Enabled = False
910         .Parent.cmdEdit.Enabled = False
920         .Parent.cmdDelete.Enabled = False
930         .Parent.cmdUpdate.Enabled = False  ' ** Which should already be disabled.
940         .Parent.Issue_Date.Enabled = False
950         .Parent.Issue_Date.Locked = True
960         .Parent.Issue_Date.ForeColor = CLR_DISABLED_FG
970         .Parent.Issue_Date.BackColor = CLR_DISABLED_BG
980       Case False
990         .Parent.cmdGeneratePOSPay.Enabled = True
1000        .Parent.cmdEdit.Enabled = True
1010        .Parent.cmdDelete.Enabled = True
1020        .Parent.Issue_Date.Enabled = True
1030        .Parent.Issue_Date.Locked = False
1040        .Parent.Issue_Date.ForeColor = CLR_BLK
1050        .Parent.Issue_Date.BackColor = CLR_WHT
1060      End Select

1070      varTmp00 = .Bank_NoRNAN_Tot
1080      Select Case IsNull(varTmp00)
          Case True
1090        .Parent.cmdUpdate.Enabled = False
1100      Case False
1110        If varTmp00 = 0 Then
1120          .Parent.cmdUpdate.Enabled = False
1130        Else
1140          .Parent.cmdUpdate.Enabled = True
1150        End If
1160      End Select

1170    End With

EXITP:
1180    Exit Sub

ERRH:
1190    Select Case ERR.Number
        Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1320    intShiftDown = (Shift And acShiftMask) > 0
1330    intAltDown = (Shift And acAltMask) > 0
1340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Check Number:     B {ppd_checknum on frmCheckPOSPay_Sub3}
        ' **   Check Number:     K {CheckNum}
        ' **   Check Memo:       M {JrnlMemo_Memo}
        ' **   Comments:         O {description}
        ' **   Currency:         R {curr_id}
        ' **   Posting Date:     S {transdate}
        ' **   Exit:             X {cmdClose on frmCheckPOSPay}
        ' **   Payee:            Y {RecurringItem}

        'I CAN'T FIGURE OUT WHAT TO GIVE THESE!
        ' **   Description:      P {pp_description on frmCheckPOSPay_Sub2}
        ' **   Payee:            P {ppd_payee on frmCheckPOSPay_Sub3}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next:             N {MoveRec on frmCheckPOSPay_Sub2}
        ' **   Previous:         P {MoveRec on frmCheckPOSPay_Sub2}
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyUp
1370        With Me
1380          intRetVal = 0
1390          MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1400        End With
1410      Case vbKeyDown
1420        With Me
1430          intRetVal = 0
1440          MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1450        End With
1460      End Select
1470    End If

        ' ** Alt keys.
1480    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1490      Select Case intRetVal
          Case vbKeyB
1500        With Me
1510          intRetVal = 0
1520          If .Parent.frmCheckPOSPay_Sub3.Visible = True Then
1530            DoCmd.SelectObject acForm, .Parent.Name, False
1540            .Parent.frmCheckPOSPay_Sub3.SetFocus
1550            .Parent.frmCheckPOSPay_Sub3.Form.ppd_checknum.SetFocus
1560          Else
1570            Beep
1580          End If
1590        End With
1600      Case vbKeyK
1610        With Me
1620          intRetVal = 0
1630          .CheckNum.SetFocus
1640        End With
1650      Case vbKeyM
1660        With Me
1670          intRetVal = 0
1680          .JrnlMemo_Memo.SetFocus
1690        End With
1700      Case vbKeyO
1710        With Me
1720          intRetVal = 0
1730          .description.SetFocus
1740        End With
'1750      Case vbKeyP
'1760        With Me
'1770          intRetVal = 0
'1780          DoCmd.SelectObject acForm, .Parent.Name, False
'1790          If .Parent.frmCheckPOSPay_Sub2.Visible = True Then
'1800            .Parent.frmCheckPOSPay_Sub2.SetFocus
'1810            .Parent.frmCheckPOSPay_Sub2.Form.pp_description.SetFocus
'1820          ElseIf .Parent.frmCheckPOSPay_Sub3.Visible = True Then
'1830            .Parent.frmCheckPOSPay_Sub3.SetFocus
'1840            .Parent.frmCheckPOSPay_Sub3.Form.ppd_payee.SetFocus
'1850          End If
'1860        End With
1750      Case vbKeyR
1760        With Me
1770          intRetVal = 0
1780          .curr_id.SetFocus
1790        End With
1800      Case vbKeyS
1810        With Me
1820          intRetVal = 0
1830          .transdate.SetFocus
1840        End With
1850      Case vbKeyX
1860        With Me
1870          intRetVal = 0
1880          .Parent.cmdClose_Click  ' ** Form Procedure: frmCheckPOSPay.
1890        End With
1900      Case vbKeyY
1910        With Me
1920          intRetVal = 0
1930          .RecurringItem.SetFocus
1940        End With
1950      End Select
1960    End If

        ' ** Ctrl keys.
1970    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1980      Select Case intRetVal
          Case vbKeyN
1990        With Me
2000          intRetVal = 0
2010          .Parent.frmCheckPOSPay_Sub2.Form.MoveRecX acCmdRecordsGoToNext  ' ** Form Procedure: frmCheckPOSPay_Sub2.
2020        End With
2030      Case vbKeyP
2040        With Me
2050          intRetVal = 0
2060          .Parent.frmCheckPOSPay_Sub2.Form.MoveRecX acCmdRecordsGoToPrevious  ' ** Form Procedure: frmCheckPOSPay_Sub2.
2070        End With
2080      Case vbKeyS
2090        intRetVal = 0
2100        cmdSave_Click  ' ** Procedure: Below.
2110      Case vbKeyTab, vbKeyReturn
2120        With Me
2130          intRetVal = 0
2140          DoCmd.SelectObject acForm, .Parent.Name, False
2150          If .Parent.cmdEdit.Enabled = True Then
2160            .Parent.cmdEdit.SetFocus
2170          ElseIf .Parent.cmdDelete.Enabled = True Then
2180            .Parent.cmdDelete.SetFocus
2190          ElseIf .Parent.cmdUpdate.Enabled = True Then
2200            .Parent.cmdUpdate.SetFocus
2210          Else
2220            .Parent.opgView.SetFocus
2230          End If
2240        End With
2250      Case vbKeyPageUp
2260        intRetVal = 0
2270        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2280      Case vbKeyPageDown
2290        intRetVal = 0
2300        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2310      End Select
2320    End If

        ' ** Ctrl-Shift keys.
2330    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyF
2350        With Me
2360          intRetVal = 0
2370          DoCmd.SelectObject acForm, .Parent.Name, False
2380          .Parent.FocusHolder.SetFocus
2390        End With
2400      Case vbKeyTab, vbKeyReturn
2410        With Me
2420          intRetVal = 0
2430          DoCmd.SelectObject acForm, .Parent.Name, False
2440          .Parent.cmdClose.SetFocus
2450        End With
2460      End Select
2470    End If

EXITP:
2480    KeyCode = intRetVal
2490    Exit Sub

ERRH:
2500    intRetVal = 0
2510    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2520    Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
2530    Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

2610    With Me
2620      .journal_USER = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2630      .ppj_datemodified = Now()
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

Public Sub cmdSave_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2710    With Me
2720  On Error Resume Next
2730      DoCmd.RunCommand acCmdSaveRecord
2740      If ERR.Number <> 0 Then
2750        Select Case ERR.Number
            Case 3314
              ' ** The field 'tblJournal_Memo.journaltype' cannot contain a Null value
              ' ** because the Required property for this field is set to True.
              ' ** THAT'S THE PURPOSE OF THE DEFAULT VALUE! WHY DON'T THEY WORK?
2760  On Error GoTo ERRH
2770          .JrnlMemo_journaltype = .journaltype
2780          .JrnlMemo_accountno = .accountno
2790          .JrnlMemo_Journal_ID = .Journal_ID
2800          .JrnlMemo_transdate = .transdate
2810          DoCmd.RunCommand acCmdSaveRecord
2820        Case Else
2830          MsgBox "ERROR: " & CStr(ERR.Number) & vbCrLf & ERR.description, vbExclamation + vbOKOnly, "Error: " & CStr(ERR.Number)
2840  On Error GoTo ERRH
2850        End Select
2860      Else
2870  On Error GoTo ERRH
2880      End If
2890      DoEvents
2900    End With

EXITP:
2910    Exit Sub

ERRH:
2920    Select Case ERR.Number
        Case Else
2930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2940    End Select
2950    Resume EXITP

End Sub

Private Sub ppj_pospay_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "ppj_pospay_lbl_DblClick"

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

Private Sub Journal_ID_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Journal_ID_lbl_DblClick"

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

Private Sub Journal_ID_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Journal_ID_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3220    intShiftDown = (Shift And acShiftMask) > 0
3230    intAltDown = (Shift And acAltMask) > 0
3240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3270        With Me
3280          intRetVal = 0
3290          .journaltype.SetFocus
3300        End With
3310      End Select
3320    End If

        ' ** Shift keys.
3330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3350        With Me
3360          intRetVal = 0
3370          If .CurrentRecord > 1 Then
3380            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3390            .FocusHolder2.SetFocus
3400            DoEvents
3410            .journal_USER.SetFocus
3420          Else
3430            DoCmd.SelectObject acForm, .Parent.Name, False
3440            .Parent.cmdClose.SetFocus
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

Private Sub journaltype_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "journaltype_lbl_DblClick"

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

Private Sub PrintCheck_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "PrintCheck_lbl_DblClick"

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

Private Sub PrintCheck_Enter()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "PrintCheck_Enter"

3810    With Me
3820      varOldValue = Empty
3830      varOldValue = .PrintCheck
3840    End With

EXITP:
3850    Exit Sub

ERRH:
3860    Select Case ERR.Number
        Case Else
3870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3880    End Select
3890    Resume EXITP

End Sub

Private Sub PrintCheck_AfterUpdate()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "PrintCheck_AfterUpdate"

3910    With Me
3920      If .PrintCheck <> varOldValue Then
3930        .ppj_changed = True
3940        cmdSave_Click  ' ** Procedure: Above.
3950        blnChanged = True
3960        UpdateChk  ' ** Procedure: Below.
3970        .Refresh
3980      End If
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

Private Sub CheckNum_lbl_DblClick(Cancel As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_lbl_DblClick"

4110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub CheckNum_Enter()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_Enter"

4210    With Me
4220      varOldValue = Empty
4230      varOldValue = .CheckNum  ' ** May be Null.
4240    End With

EXITP:
4250    Exit Sub

ERRH:
4260    Select Case ERR.Number
        Case Else
4270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4280    End Select
4290    Resume EXITP

End Sub

Private Sub CheckNum_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_AfterUpdate"

4310    With Me
4320      Select Case IsNull(.CheckNum)
          Case True
4330        Select Case IsNull(varOldValue)
            Case True
              ' ** Both Null, so leave it alone.
4340        Case False
4350          .ppj_changed = True
4360          blnChanged = True
4370          cmdSave_Click  ' ** Procedure: Above.
4380        End Select
4390      Case False
4400        Select Case IsNull(varOldValue)
            Case True
              ' ** New value.
4410          .ppj_changed = True
4420          blnChanged = True
4430          cmdSave_Click  ' ** Procedure: Above.
4440        Case False
4450          If .CheckNum <> varOldValue Then
                ' ** Different value.
4460            .ppj_changed = True
4470            blnChanged = True
4480            cmdSave_Click  ' ** Procedure: Above.
4490          End If
4500        End Select
4510      End Select
4520      UpdateChk  ' ** Procedure: Below.
4530    End With

EXITP:
4540    Exit Sub

ERRH:
4550    Select Case ERR.Number
        Case Else
4560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4570    End Select
4580    Resume EXITP

End Sub

Private Sub ppj_void_lbl_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "ppj_void_lbl_DblClick"

4610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

4710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

4810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4820    Exit Sub

ERRH:
4830    Select Case ERR.Number
        Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

4910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub transdate_Enter()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_Enter"

5010    With Me
5020      varOldValue = Empty
5030      varOldValue = .transdate
5040    End With

EXITP:
5050    Exit Sub

ERRH:
5060    Select Case ERR.Number
        Case Else
5070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5080    End Select
5090    Resume EXITP

End Sub

Private Sub transdate_AfterUpdate()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_AfterUpdate"

5110    With Me
5120      Select Case IsNull(.transdate)
          Case True
5130        .transdate = Date
5140        If .transdate <> varOldValue Then
5150          .ppj_changed = True
5160          blnChanged = True
5170          cmdSave_Click  ' ** Procedure: Above.
5180        End If
5190      Case False
5200        If .transdate <> varOldValue Then
5210          .ppj_changed = True
5220          blnChanged = True
5230          cmdSave_Click  ' ** Procedure: Above.
5240        End If
5250      End Select
5260      UpdateChk  ' ** Procedure: Below.
5270    End With

EXITP:
5280    Exit Sub

ERRH:
5290    Select Case ERR.Number
        Case Else
5300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5310    End Select
5320    Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_lbl_DblClick"

5410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Sub

Private Sub icash_Enter()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_Enter"

5510    With Me
5520      varOldValue = Empty
5530      varOldValue = .ICash
5540    End With

EXITP:
5550    Exit Sub

ERRH:
5560    Select Case ERR.Number
        Case Else
5570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5580    End Select
5590    Resume EXITP

End Sub

Private Sub icash_AfterUpdate()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_AfterUpdate"

5610    With Me
5620      Select Case IsNull(.ICash)
          Case True
5630        .ICash = 0
5640        If ICash <> varOldValue Then
5650          .ppj_changed = True
5660          blnChanged = True
5670          cmdSave_Click  ' ** Procedure: Above.
5680        End If
5690      Case False
5700        If ICash <> varOldValue Then
5710          .ppj_changed = True
5720          blnChanged = True
5730          cmdSave_Click  ' ** Procedure: Above.
5740        End If
5750      End Select
5760      UpdateChk  ' ** Procedure: Below.
5770    End With

EXITP:
5780    Exit Sub

ERRH:
5790    Select Case ERR.Number
        Case Else
5800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5810    End Select
5820    Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_lbl_DblClick"

5910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5920    Exit Sub

ERRH:
5930    Select Case ERR.Number
        Case Else
5940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5950    End Select
5960    Resume EXITP

End Sub

Private Sub pcash_Enter()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_Enter"

6010    With Me
6020      varOldValue = Empty
6030      varOldValue = .PCash
6040    End With

EXITP:
6050    Exit Sub

ERRH:
6060    Select Case ERR.Number
        Case Else
6070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6080    End Select
6090    Resume EXITP

End Sub

Private Sub pcash_AfterUpdate()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_AfterUpdate"

6110    With Me
6120      Select Case IsNull(.PCash)
          Case True
6130        .PCash = 0
6140        If PCash <> varOldValue Then
6150          .ppj_changed = True
6160          blnChanged = True
6170          cmdSave_Click  ' ** Procedure: Above.
6180        End If
6190      Case False
6200        If PCash <> varOldValue Then
6210          .ppj_changed = True
6220          blnChanged = True
6230          cmdSave_Click  ' ** Procedure: Above.
6240        End If
6250      End Select
6260      UpdateChk  ' ** Procedure: Below.
6270    End With

EXITP:
6280    Exit Sub

ERRH:
6290    Select Case ERR.Number
        Case Else
6300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6310    End Select
6320    Resume EXITP

End Sub

Private Sub amount_lbl_DblClick(Cancel As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "amount_lbl_DblClick"

6410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

6510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6520    Exit Sub

ERRH:
6530    Select Case ERR.Number
        Case Else
6540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6550    End Select
6560    Resume EXITP

End Sub

Private Sub RecurringItem_lbl_DblClick(Cancel As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_lbl_DblClick"

6610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6620    Exit Sub

ERRH:
6630    Select Case ERR.Number
        Case Else
6640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6650    End Select
6660    Resume EXITP

End Sub

Private Sub RecurringItem_Enter()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_Enter"

6710    With Me
6720      varOldValue = Empty
6730      varOldValue = .RecurringItem
6740    End With

EXITP:
6750    Exit Sub

ERRH:
6760    Select Case ERR.Number
        Case Else
6770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6780    End Select
6790    Resume EXITP

End Sub

Private Sub RecurringItem_AfterUpdate()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItem_AfterUpdate"

6810    With Me
6820      Select Case IsNull(.RecurringItem)
          Case True
6830        .RecurringItem = varOldValue
6840        cmdSave_Click  ' ** Procedure: Above.
6850      Case False
6860        If .RecurringItem <> varOldValue Then
6870          .ppj_changed = True
6880          blnChanged = True
6890          cmdSave_Click  ' ** Procedure: Above.
6900        End If
6910      End Select
6920      UpdateChk  ' ** Procedure: Below.
6930    End With

EXITP:
6940    Exit Sub

ERRH:
6950    Select Case ERR.Number
        Case Else
6960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6970    End Select
6980    Resume EXITP

End Sub

Private Sub JrnlMemo_Memo_lbl_DblClick(Cancel As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlMemo_Memo_lbl_DblClick"

7010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Private Sub JrnlMemo_Memo_Enter()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlMemo_Memo_Enter"

7110    With Me
7120      varOldValue = Empty
7130      varOldValue = .JrnlMemo_Memo
7140    End With

EXITP:
7150    Exit Sub

ERRH:
7160    Select Case ERR.Number
        Case Else
7170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7180    End Select
7190    Resume EXITP

End Sub

Private Sub JrnlMemo_Memo_AfterUpdate()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlMemo_Memo_AfterUpdate"

7210    With Me
7220      Select Case IsNull(.JrnlMemo_Memo)
          Case True
7230        Select Case IsNull(varOldValue)
            Case True
              ' ** Both Null, so leave it alone.
7240        Case False
7250          .ppj_changed = True
7260          blnChanged = True
7270          cmdSave_Click  ' ** Procedure: Above.
7280        End Select
7290      Case False
7300        Select Case IsNull(varOldValue)
            Case True
7310          .ppj_changed = True
7320          blnChanged = True
7330          cmdSave_Click  ' ** Procedure: Above.
7340        Case False
7350          If .JrnlMemo_Memo <> varOldValue Then
7360            .ppj_changed = True
7370            blnChanged = True
7380            cmdSave_Click  ' ** Procedure: Above.
7390          End If
7400        End Select
7410      End Select
7420      UpdateChk  ' ** Procedure: Below.
7430    End With

EXITP:
7440    Exit Sub

ERRH:
7450    Select Case ERR.Number
        Case Else
7460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7470    End Select
7480    Resume EXITP

End Sub

Private Sub description_lbl_DblClick(Cancel As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "description_lbl_DblClick"

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

Private Sub description_Enter()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "description_Enter"

7610    With Me
7620      varOldValue = Empty
7630      varOldValue = .description
7640    End With

EXITP:
7650    Exit Sub

ERRH:
7660    Select Case ERR.Number
        Case Else
7670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7680    End Select
7690    Resume EXITP

End Sub

Private Sub description_AfterUpdate()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "description_AfterUpdate"

7710    With Me
7720      Select Case IsNull(.description)
          Case True
7730        Select Case IsNull(varOldValue)
            Case True
              ' ** Both Null, so leave it alone.
7740        Case False
7750          .ppj_changed = True
7760          blnChanged = True
7770          cmdSave_Click  ' ** Procedure: Above.
7780        End Select
7790      Case False
7800        Select Case IsNull(varOldValue)
            Case True
7810          .ppj_changed = True
7820          blnChanged = True
7830          cmdSave_Click  ' ** Procedure: Above.
7840        Case False
7850          If .description <> varOldValue Then
7860            .ppj_changed = True
7870            blnChanged = True
7880            cmdSave_Click  ' ** Procedure: Above.
7890          End If
7900        End Select
7910      End Select
7920      UpdateChk  ' ** Procedure: Below.
7930    End With

EXITP:
7940    Exit Sub

ERRH:
7950    Select Case ERR.Number
        Case Else
7960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7970    End Select
7980    Resume EXITP

End Sub

Private Sub Bank_Name_lbl_DblClick(Cancel As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_Name_lbl_DblClick"

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

Private Sub Bank_RoutingNumber_lbl_DblClick(Cancel As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_RoutingNumber_lbl_DblClick"

8110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8120    Exit Sub

ERRH:
8130    Select Case ERR.Number
        Case Else
8140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8150    End Select
8160    Resume EXITP

End Sub

Private Sub Bank_AccountNumber_lbl_DblClick(Cancel As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_AccountNumber_lbl_DblClick"

8210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8220    Exit Sub

ERRH:
8230    Select Case ERR.Number
        Case Else
8240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8250    End Select
8260    Resume EXITP

End Sub

Private Sub journal_USER_lbl_DblClick(Cancel As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_lbl_DblClick"

8310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8320    Exit Sub

ERRH:
8330    Select Case ERR.Number
        Case Else
8340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8350    End Select
8360    Resume EXITP

End Sub

Private Sub journal_USER_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "journal_USER_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8420    intShiftDown = (Shift And acShiftMask) > 0
8430    intAltDown = (Shift And acAltMask) > 0
8440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8470        With Me
8480          intRetVal = 0
8490          lngRecsCur = RecCnt  ' ** Function: Below.
8500          If .CurrentRecord < lngRecsCur Then
8510            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
8520            .FocusHolder.SetFocus
8530            DoEvents
8540            .Journal_ID.SetFocus
8550          Else
8560            DoCmd.SelectObject acForm, .Parent.Name, False
8570            If .Parent.cmdEdit.Enabled = True Then
8580              .Parent.cmdEdit.SetFocus
8590            ElseIf .Parent.cmdDelete.Enabled = True Then
8600              .Parent.cmdDelete.SetFocus
8610            ElseIf .Parent.cmdUpdate.Enabled = True Then
8620              .Parent.cmdUpdate.SetFocus
8630            Else
8640              .Parent.opgView.SetFocus
8650            End If
8660          End If
8670        End With
8680      End Select
8690    End If

        ' ** Shift keys.
8700    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8710      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8720        With Me
8730          intRetVal = 0
8740          .description.SetFocus
8750        End With
8760      End Select
8770    End If

EXITP:
8780    KeyCode = intRetVal
8790    Exit Sub

ERRH:
8800    intRetVal = 0
8810    Select Case ERR.Number
        Case Else
8820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8830    End Select
8840    Resume EXITP

End Sub

Public Sub UpdateChk()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateChk"

8910    With Me
8920      If blnChanged = True Then
8930        If .Parent.cmdUpdate.Enabled = False Then
8940          .Parent.cmdUpdate.Enabled = True
8950        End If
8960      End If
8970    End With

EXITP:
8980    Exit Sub

ERRH:
8990    Select Case ERR.Number
        Case Else
9000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9010    End Select
9020    Resume EXITP

End Sub

Private Sub CtlArray1_Load()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "CtlArray1_Load"

        Dim lngE As Long

9110    With Me

9120      If lngCtl1s = 0& Or IsEmpty(arr_varCtl1) Then

9130        lngCtl1s = 0&
9140        ReDim arr_varCtl1(C_ELEMS, 0)

9150        lngCtl1s = lngCtl1s + 1&
9160        lngE = lngCtl1s - 1&
9170        ReDim Preserve arr_varCtl1(C_ELEMS, lngE)
9180        arr_varCtl1(C_CNAM, lngE) = "FocusHolder2"
9190        arr_varCtl1(C_LFT, lngE) = .FocusHolder2.Left
9200        arr_varCtl1(C_LBL, lngE) = Null
9210        arr_varCtl1(C_LBL2, lngE) = Null
9220        arr_varCtl1(C_LIN, lngE) = Null
9230        arr_varCtl1(C_DMLB, lngE) = Null
9240        arr_varCtl1(C_DML2, lngE) = Null
9250        arr_varCtl1(C_DMLN, lngE) = Null

9260        lngCtl1s = lngCtl1s + 1&
9270        lngE = lngCtl1s - 1&
9280        ReDim Preserve arr_varCtl1(C_ELEMS, lngE)
9290        arr_varCtl1(C_CNAM, lngE) = "journal_USER"
9300        arr_varCtl1(C_LFT, lngE) = .journal_USER.Left
9310        arr_varCtl1(C_LBL, lngE) = "journal_USER_lbl"
9320        arr_varCtl1(C_LBL2, lngE) = "journal_USER_lbl2"
9330        arr_varCtl1(C_LIN, lngE) = "journal_USER_lbl_line"
9340        arr_varCtl1(C_DMLB, lngE) = "journal_USER_lbl_dim_hi"
9350        arr_varCtl1(C_DML2, lngE) = "journal_USER_lbl2_dim_hi"
9360        arr_varCtl1(C_DMLN, lngE) = "journal_USER_lbl_line_dim_hi"

9370        lngCtl1s = lngCtl1s + 1&
9380        lngE = lngCtl1s - 1&
9390        ReDim Preserve arr_varCtl1(C_ELEMS, lngE)
9400        arr_varCtl1(C_CNAM, lngE) = "Bank_AccountNumber"
9410        arr_varCtl1(C_LFT, lngE) = .Bank_AccountNumber.Left
9420        arr_varCtl1(C_LBL, lngE) = "Bank_AccountNumber_lbl"
9430        arr_varCtl1(C_LBL2, lngE) = "Bank_AccountNumber_lbl2"
9440        arr_varCtl1(C_LIN, lngE) = "Bank_AccountNumber_lbl_line"
9450        arr_varCtl1(C_DMLB, lngE) = "Bank_AccountNumber_lbl_dim_hi"
9460        arr_varCtl1(C_DML2, lngE) = "Bank_AccountNumber_lbl2_dim_hi"
9470        arr_varCtl1(C_DMLN, lngE) = "Bank_AccountNumber_lbl_line_dim_hi"

9480        lngCtl1s = lngCtl1s + 1&
9490        lngE = lngCtl1s - 1&
9500        ReDim Preserve arr_varCtl1(C_ELEMS, lngE)
9510        arr_varCtl1(C_CNAM, lngE) = "Bank_RoutingNumber"
9520        arr_varCtl1(C_LFT, lngE) = .Bank_RoutingNumber.Left
9530        arr_varCtl1(C_LBL, lngE) = "Bank_RoutingNumber_lbl"
9540        arr_varCtl1(C_LBL2, lngE) = "Bank_RoutingNumber_lbl2"
9550        arr_varCtl1(C_LIN, lngE) = "Bank_RoutingNumber_lbl_line"
9560        arr_varCtl1(C_DMLB, lngE) = "Bank_RoutingNumber_lbl_dim_hi"
9570        arr_varCtl1(C_DML2, lngE) = "Bank_RoutingNumber_lbl2_dim_hi"
9580        arr_varCtl1(C_DMLN, lngE) = "Bank_RoutingNumber_lbl_line_dim_hi"

9590        lngCtl1s = lngCtl1s + 1&
9600        lngE = lngCtl1s - 1&
9610        ReDim Preserve arr_varCtl1(C_ELEMS, lngE)
9620        arr_varCtl1(C_CNAM, lngE) = "Bank_Name"
9630        arr_varCtl1(C_LFT, lngE) = .Bank_Name.Left
9640        arr_varCtl1(C_LBL, lngE) = "Bank_Name_lbl"
9650        arr_varCtl1(C_LBL2, lngE) = Null
9660        arr_varCtl1(C_LIN, lngE) = "Bank_Name_lbl_line"
9670        arr_varCtl1(C_DMLB, lngE) = "Bank_Name_lbl_dim_hi"
9680        arr_varCtl1(C_DML2, lngE) = Null
9690        arr_varCtl1(C_DMLN, lngE) = "Bank_Name_lbl_line_dim_hi"

9700        lngCtl1s = lngCtl1s + 1&
9710        lngE = lngCtl1s - 1&
9720        ReDim Preserve arr_varCtl1(C_ELEMS, lngE)
9730        arr_varCtl1(C_CNAM, lngE) = "description"
9740        arr_varCtl1(C_LFT, lngE) = .description.Left
9750        arr_varCtl1(C_LBL, lngE) = "description_lbl"
9760        arr_varCtl1(C_LBL2, lngE) = "description_msg_lbl"
9770        arr_varCtl1(C_LIN, lngE) = "description_lbl_line"
9780        arr_varCtl1(C_DMLB, lngE) = "description_lbl_dim_hi"
9790        arr_varCtl1(C_DML2, lngE) = Null
9800        arr_varCtl1(C_DMLN, lngE) = "description_lbl_line_dim_hi"

9810        lngCtl1s = lngCtl1s + 1&
9820        lngE = lngCtl1s - 1&
9830        ReDim Preserve arr_varCtl1(C_ELEMS, lngE)
9840        arr_varCtl1(C_CNAM, lngE) = "JrnlMemo_Memo"
9850        arr_varCtl1(C_LFT, lngE) = .JrnlMemo_Memo.Left
9860        arr_varCtl1(C_LBL, lngE) = "JrnlMemo_Memo_lbl"
9870        arr_varCtl1(C_LBL2, lngE) = "JrnlMemo_Memo_msg_lbl"
9880        arr_varCtl1(C_LIN, lngE) = "JrnlMemo_Memo_lbl_line"
9890        arr_varCtl1(C_DMLB, lngE) = "JrnlMemo_Memo_lbl_dim_hi"
9900        arr_varCtl1(C_DML2, lngE) = Null
9910        arr_varCtl1(C_DMLN, lngE) = "JrnlMemo_Memo_lbl_line_dim_hi"

9920        lngCtl1s = lngCtl1s + 1&
9930        lngE = lngCtl1s - 1&
9940        ReDim Preserve arr_varCtl1(C_ELEMS, lngE)
9950        arr_varCtl1(C_CNAM, lngE) = "RecurringItem"
9960        arr_varCtl1(C_LFT, lngE) = .RecurringItem.Left
9970        arr_varCtl1(C_LBL, lngE) = "RecurringItem_lbl"
9980        arr_varCtl1(C_LBL2, lngE) = Null
9990        arr_varCtl1(C_LIN, lngE) = "RecurringItem_lbl_line"
10000       arr_varCtl1(C_DMLB, lngE) = "RecurringItem_lbl_dim_hi"
10010       arr_varCtl1(C_DML2, lngE) = Null
10020       arr_varCtl1(C_DMLN, lngE) = "RecurringItem_lbl_line_dim_hi"

10030     End If  ' ** lngCtl1s.

10040   End With

EXITP:
10050   Exit Sub

ERRH:
10060   Select Case ERR.Number
        Case Else
10070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10080   End Select
10090   Resume EXITP

End Sub

Private Sub CtlArray2_Load()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "CtlArray2_Load"

        Dim lngX As Long, lngY As Long, lngE As Long

10110   With Me

10120     If lngCtl2s = 0& Or IsEmpty(arr_varCtl2) Then

            ' ** Make sure the first one is loaded.
10130       CtlArray1_Load  ' ** Procedure: Above.

10140       lngCtl2s = 0&
10150       ReDim arr_varCtl2(C_ELEMS, 0)

10160       For lngX = 0& To (lngCtl1s - 1&)
10170         lngCtl2s = lngCtl2s + 1&
10180         lngE = lngCtl2s - 1&
10190         ReDim Preserve arr_varCtl2(C_ELEMS, lngE)
10200         For lngY = 0& To C_ELEMS
10210           arr_varCtl2(lngY, lngE) = arr_varCtl1(lngY, lngX)
10220         Next  ' ** lngY.
10230       Next  ' ** lngX.

10240       lngCtl2s = lngCtl2s + 1&
10250       lngE = lngCtl2s - 1&
10260       ReDim Preserve arr_varCtl2(C_ELEMS, lngE)
10270       arr_varCtl2(C_CNAM, lngE) = "curr_id"
10280       arr_varCtl2(C_LFT, lngE) = .curr_id.Left
10290       arr_varCtl2(C_LBL, lngE) = "curr_id_lbl"
10300       arr_varCtl2(C_LBL2, lngE) = Null
10310       arr_varCtl2(C_LIN, lngE) = "curr_id_lbl_line"
10320       arr_varCtl2(C_DMLB, lngE) = "curr_id_lbl_dim_hi"
10330       arr_varCtl2(C_DML2, lngE) = Null
10340       arr_varCtl2(C_DMLN, lngE) = "curr_id_lbl_line_dim_hi"

10350       lngCtl2s = lngCtl2s + 1&
10360       lngE = lngCtl2s - 1&
10370       ReDim Preserve arr_varCtl2(C_ELEMS, lngE)
10380       arr_varCtl2(C_CNAM, lngE) = "amount"
10390       arr_varCtl2(C_LFT, lngE) = .amount.Left
10400       arr_varCtl2(C_LBL, lngE) = "amount_lbl"
10410       arr_varCtl2(C_LBL2, lngE) = "amount_lbl2"
10420       arr_varCtl2(C_LIN, lngE) = "amount_lbl_line"
10430       arr_varCtl2(C_DMLB, lngE) = "amount_lbl_dim_hi"
10440       arr_varCtl2(C_DML2, lngE) = "amount_lbl2_dim_hi"
10450       arr_varCtl2(C_DMLN, lngE) = "amount_lbl_line_dim_hi"

10460       lngCtl2s = lngCtl2s + 1&
10470       lngE = lngCtl2s - 1&
10480       ReDim Preserve arr_varCtl2(C_ELEMS, lngE)
10490       arr_varCtl2(C_CNAM, lngE) = "pcash"
10500       arr_varCtl2(C_LFT, lngE) = .PCash.Left
10510       arr_varCtl2(C_LBL, lngE) = "pcash_lbl"
10520       arr_varCtl2(C_LBL2, lngE) = "pcash_lbl2"
10530       arr_varCtl2(C_LIN, lngE) = "pcash_lbl_line"
10540       arr_varCtl2(C_DMLB, lngE) = "pcash_lbl_dim_hi"
10550       arr_varCtl2(C_DML2, lngE) = "pcash_lbl2_dim_hi"
10560       arr_varCtl2(C_DMLN, lngE) = "pcash_lbl_line_dim_hi"

10570       lngCtl2s = lngCtl2s + 1&
10580       lngE = lngCtl2s - 1&
10590       ReDim Preserve arr_varCtl2(C_ELEMS, lngE)
10600       arr_varCtl2(C_CNAM, lngE) = "icash"
10610       arr_varCtl2(C_LFT, lngE) = .ICash.Left
10620       arr_varCtl2(C_LBL, lngE) = "icash_lbl"
10630       arr_varCtl2(C_LBL2, lngE) = "icash_lbl2"
10640       arr_varCtl2(C_LIN, lngE) = "icash_lbl_line"
10650       arr_varCtl2(C_DMLB, lngE) = "icash_lbl_dim_hi"
10660       arr_varCtl2(C_DML2, lngE) = "icash_lbl2_dim_hi"
10670       arr_varCtl2(C_DMLN, lngE) = "icash_lbl_line_dim_hi"

10680       lngCtl2s = lngCtl2s + 1&
10690       lngE = lngCtl2s - 1&
10700       ReDim Preserve arr_varCtl2(C_ELEMS, lngE)
10710       arr_varCtl2(C_CNAM, lngE) = "transdate"
10720       arr_varCtl2(C_LFT, lngE) = .transdate.Left
10730       arr_varCtl2(C_LBL, lngE) = "transdate_lbl"
10740       arr_varCtl2(C_LBL2, lngE) = "transdate_lbl2"
10750       arr_varCtl2(C_LIN, lngE) = "transdate_lbl_line"
10760       arr_varCtl2(C_DMLB, lngE) = "transdate_lbl_dim_hi"
10770       arr_varCtl2(C_DML2, lngE) = "transdate_lbl2_dim_hi"
10780       arr_varCtl2(C_DMLN, lngE) = "transdate_lbl_line_dim_hi"

10790       lngCtl2s = lngCtl2s + 1&
10800       lngE = lngCtl2s - 1&
10810       ReDim Preserve arr_varCtl2(C_ELEMS, lngE)
10820       arr_varCtl2(C_CNAM, lngE) = "shortname"
10830       arr_varCtl2(C_LFT, lngE) = .shortname.Left
10840       arr_varCtl2(C_LBL, lngE) = "shortname_lbl"
10850       arr_varCtl2(C_LBL2, lngE) = Null
10860       arr_varCtl2(C_LIN, lngE) = "shortname_lbl_line"
10870       arr_varCtl2(C_DMLB, lngE) = "shortname_lbl_dim_hi"
10880       arr_varCtl2(C_DML2, lngE) = Null
10890       arr_varCtl2(C_DMLN, lngE) = "shortname_lbl_line_dim_hi"

10900       lngCtl2s = lngCtl2s + 1&
10910       lngE = lngCtl2s - 1&
10920       ReDim Preserve arr_varCtl2(C_ELEMS, lngE)
10930       arr_varCtl2(C_CNAM, lngE) = "accountno"
10940       arr_varCtl2(C_LFT, lngE) = .accountno.Left
10950       arr_varCtl2(C_LBL, lngE) = "accountno_lbl"
10960       arr_varCtl2(C_LBL2, lngE) = "accountno_lbl2"
10970       arr_varCtl2(C_LIN, lngE) = "accountno_lbl_line"
10980       arr_varCtl2(C_DMLB, lngE) = "accountno_lbl_dim_hi"
10990       arr_varCtl2(C_DML2, lngE) = "accountno_lbl2_dim_hi"
11000       arr_varCtl2(C_DMLN, lngE) = "accountno_lbl_line_dim_hi"

11010     End If  ' ** lngCtl2s.

11020   End With

EXITP:
11030   Exit Sub

ERRH:
11040   Select Case ERR.Number
        Case Else
11050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11060   End Select
11070   Resume EXITP

End Sub

Public Sub IncludeCurrency(blnInclude As Boolean)
' ** Uses lngCtl1s, arr_varCtl1().

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim lngX As Long

11110   With Me
11120     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
11130       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
11140     End If
11150     blnSortHere = False: blnResort = False
11160     Select Case blnInclude
          Case True
11170       .Width = (lngForm_Width + lngCurrID_Width)
11180       For lngX = 0& To (lngCtl1s - 1&)
              ' ** Array starts at the right and moves left.
11190         If IsNull(arr_varCtl1(C_LBL, lngX)) = False Then
11200           If .Sort_line.Left = arr_varCtl1(C_LFT, lngX) Then blnSortHere = True
11210         End If
11220         .Controls(arr_varCtl1(C_CNAM, lngX)).Left = (arr_varCtl1(C_LFT, lngX) + lngCurrID_Width)
11230         If IsNull(arr_varCtl1(C_LBL, lngX)) = False Then
11240           .Controls(arr_varCtl1(C_LBL, lngX)).Left = (arr_varCtl1(C_LFT, lngX) + lngCurrID_Width)
11250           If IsNull(arr_varCtl1(C_DMLB, lngX)) = False Then
11260             .Controls(arr_varCtl1(C_DMLB, lngX)).Left = ((arr_varCtl1(C_LFT, lngX) + lngCurrID_Width) + lngTpp)
11270           End If
11280         End If
11290         If IsNull(arr_varCtl1(C_LBL2, lngX)) = False Then
11300           .Controls(arr_varCtl1(C_LBL2, lngX)).Left = (arr_varCtl1(C_LFT, lngX) + lngCurrID_Width)
11310           If IsNull(arr_varCtl1(C_DML2, lngX)) = False Then
11320             .Controls(arr_varCtl1(C_DML2, lngX)).Left = ((arr_varCtl1(C_LFT, lngX) + lngCurrID_Width) + lngTpp)
11330           End If
11340         End If
11350         If IsNull(arr_varCtl1(C_LIN, lngX)) = False Then
11360           .Controls(arr_varCtl1(C_LIN, lngX)).Left = (arr_varCtl1(C_LFT, lngX) + lngCurrID_Width)
11370           If IsNull(arr_varCtl1(C_DMLN, lngX)) = False Then
11380             .Controls(arr_varCtl1(C_DMLN, lngX)).Left = ((arr_varCtl1(C_LFT, lngX) + lngCurrID_Width) + lngTpp)
11390           End If
11400         End If
11410         If blnSortHere = True Then
11420           .Sort_line.Left = .Controls(arr_varCtl1(C_LBL, lngX)).Left
11430           .Sort_lbl.Left = ((.Controls(arr_varCtl1(C_LBL, lngX)).Left + .Controls(arr_varCtl1(C_LBL, lngX)).Width) - lngSortLbl_Width)
11440         End If
11450         blnSortHere = False
11460       Next  ' ** lngX.
11470       .curr_id.Visible = True
            'CURRENCY LBL SHOULD BE OFFSET!
11480       lngRecsCur = RecCnt  ' ** Function: Below.
11490       If lngRecsCur = 0& Then
11500         .curr_id_lbl.ForeColor = WIN_CLR_DISF
11510         .curr_id_lbl_dim_hi.Visible = True
11520         .curr_id_lbl_line.BorderColor = WIN_CLR_DISR
11530         .curr_id_lbl_line_dim_hi.Visible = True
11540       Else
11550         .curr_id_lbl.ForeColor = CLR_DKGRY2
11560         .curr_id_lbl_dim_hi.Visible = False
11570         .curr_id_lbl_line.BorderColor = CLR_DKGRY
11580         .curr_id_lbl_line_dim_hi.Visible = False
11590       End If
11600       .curr_id_lbl.Visible = True
11610       .curr_id_lbl_line.Visible = True
11620     Case False
11630 On Error Resume Next
11640       .FocusHolder.SetFocus
11650 On Error GoTo ERRH
11660       If .Sort_line.Left = .curr_id_lbl_line.Left Then blnSortHere = True
11670       .curr_id.Visible = False
11680       .curr_id_lbl.Visible = False
11690       .curr_id_lbl_dim_hi.Visible = False
11700       .curr_id_lbl_line.Visible = False
11710       .curr_id_lbl_line_dim_hi.Visible = False
11720       If blnSortHere = True Then
11730         blnResort = True
11740       End If
11750       blnSortHere = False
11760       For lngX = (lngCtl1s - 1&) To 0& Step -1&
              ' ** Array starts at the right and moves left.
11770         If IsNull(arr_varCtl2(C_LBL, lngX)) = False Then
11780           If .Sort_line.Left = arr_varCtl1(C_LFT, lngX) Then blnSortHere = True
11790         End If
11800         .Controls(arr_varCtl1(C_CNAM, lngX)).Left = arr_varCtl1(C_LFT, lngX)
11810         If IsNull(arr_varCtl1(C_LBL, lngX)) = False Then
11820           .Controls(arr_varCtl1(C_LBL, lngX)).Left = arr_varCtl1(C_LFT, lngX)
11830           If IsNull(arr_varCtl1(C_DMLB, lngX)) = False Then
11840             .Controls(arr_varCtl1(C_DMLB, lngX)).Left = (arr_varCtl1(C_LFT, lngX) + lngTpp)
11850           End If
11860         End If
11870         If IsNull(arr_varCtl1(C_LBL2, lngX)) = False Then
11880           .Controls(arr_varCtl1(C_LBL2, lngX)).Left = arr_varCtl1(C_LFT, lngX)
11890           If IsNull(arr_varCtl1(C_DML2, lngX)) = False Then
11900             .Controls(arr_varCtl1(C_DML2, lngX)).Left = (arr_varCtl1(C_LFT, lngX) + lngTpp)
11910           End If
11920         End If
11930         If IsNull(arr_varCtl1(C_LIN, lngX)) = False Then
11940           .Controls(arr_varCtl1(C_LIN, lngX)).Left = arr_varCtl1(C_LFT, lngX)
11950           If IsNull(arr_varCtl1(C_DMLN, lngX)) = False Then
11960             .Controls(arr_varCtl1(C_DMLN, lngX)).Left = (arr_varCtl1(C_LFT, lngX) + lngTpp)
11970           End If
11980         End If
11990         If blnSortHere = True Then
12000           .Sort_line.Left = .Controls(arr_varCtl1(C_LBL, lngX)).Left
12010           .Sort_lbl.Left = ((.Controls(arr_varCtl1(C_LBL, lngX)).Left + .Controls(arr_varCtl1(C_LBL, lngX)).Width) - lngSortLbl_Width)
12020         End If
12030         blnSortHere = False
12040       Next  ' ** lngX.
12050       .Width = lngForm_Width
12060     End Select
12070     If blnResort = True Then
12080       blnResort = False
12090       .Sort_line.Visible = False
12100       .Sort_lbl.Visible = False
12110       SortNow "Form_Load"  ' ** Procedure: Below.
12120     End If
12130   End With

EXITP:
12140   Exit Sub

ERRH:
12150   Select Case ERR.Number
        Case Else
12160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12170   End Select
12180   Resume EXITP

End Sub

Public Sub ShowVoid(blnShow As Boolean)
' ** Uses lngCtl2s, arr_varCtl2().

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowVoid"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim blnFound As Boolean, blnCurrency As Boolean
        Dim lngX As Long

12210   With Me
12220     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
12230       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
12240     End If
12250     blnSortHere = False: blnResort = False
12260     CtlArray2_Load  ' ** Procedure: Above.
12270     Select Case blnShow
          Case True
12280       If .ppj_void.Visible = False Then
12290         .Width = (lngForm_Width + lngVoid_Width)
12300         For lngX = 0& To (lngCtl2s - 1&)
                ' ** This will move them from their current position, whether or not curr_id is visible.
12310           If IsNull(arr_varCtl2(C_LBL, lngX)) = False Then
12320             If .Sort_line.Left = .Controls(arr_varCtl2(C_LBL, lngX)).Left Then blnSortHere = True
12330           End If
                'CURRENCY EXCEPTION!
12340           .Controls(arr_varCtl2(C_CNAM, lngX)).Left = (.Controls(arr_varCtl2(C_CNAM, lngX)).Left + lngVoid_Width)
12350           If IsNull(arr_varCtl2(C_LBL, lngX)) = False Then
12360             .Controls(arr_varCtl2(C_LBL, lngX)).Left = (.Controls(arr_varCtl2(C_LBL, lngX)).Left + lngVoid_Width)
12370             If IsNull(arr_varCtl2(C_DMLB, lngX)) = False Then
12380               .Controls(arr_varCtl2(C_DMLB, lngX)).Left = ((.Controls(arr_varCtl2(C_LBL, lngX)).Left + lngVoid_Width) + lngTpp)
12390             End If
12400           End If
12410           If IsNull(arr_varCtl2(C_LBL2, lngX)) = False Then
12420             .Controls(arr_varCtl2(C_LBL2, lngX)).Left = (.Controls(arr_varCtl2(C_LBL2, lngX)).Left + lngVoid_Width)
12430             If IsNull(arr_varCtl2(C_DML2, lngX)) = False Then
12440               .Controls(arr_varCtl2(C_DML2, lngX)).Left = ((.Controls(arr_varCtl2(C_LBL2, lngX)).Left + lngVoid_Width) + lngTpp)
12450             End If
12460           End If
12470           If IsNull(arr_varCtl2(C_LIN, lngX)) = False Then
12480             .Controls(arr_varCtl2(C_LIN, lngX)).Left = (.Controls(arr_varCtl2(C_LIN, lngX)).Left + lngVoid_Width)
12490             If IsNull(arr_varCtl2(C_DMLN, lngX)) = False Then
12500               .Controls(arr_varCtl2(C_DMLN, lngX)).Left = ((.Controls(arr_varCtl2(C_LIN, lngX)).Left + lngVoid_Width) + lngTpp)
12510             End If
12520           End If
12530           If blnSortHere = True Then
12540             .Sort_line.Left = .Controls(arr_varCtl2(C_LBL, lngX)).Left
12550             .Sort_lbl.Left = ((.Controls(arr_varCtl2(C_LBL, lngX)).Left + .Controls(arr_varCtl2(C_LBL, lngX)).Width) - lngSortLbl_Width)
12560           End If
12570           blnSortHere = False
12580         Next  ' ** lngX.
12590         .ppj_void.Visible = True
12600         .ppj_void_lbl.Visible = True
12610         .ppj_void_lbl_line.Visible = True
12620         .ppj_void_box1.Visible = True
12630         .ppj_void_box2.Visible = True
12640       End If
12650     Case False
12660       If .ppj_void.Visible = True Then
12670         .FocusHolder.SetFocus
12680         DoEvents
12690         If .Sort_line.Left = .ppj_void_lbl_line.Left Then blnSortHere = True
12700         .ppj_void.Visible = False
12710         .ppj_void_lbl.Visible = False
12720         .ppj_void_lbl_line.Visible = False
12730         .ppj_void_box1.Visible = False
12740         .ppj_void_box2.Visible = False
12750         If blnSortHere = True Then
12760           blnResort = True
12770         End If
12780         blnSortHere = False
12790         blnFound = False: blnCurrency = False
12800         For lngX = (lngCtl2s - 1&) To 0& Step -1&
12810           If arr_varCtl2(C_CNAM, lngX) = "curr_id" Then
12820             blnFound = True
12830             If .Controls(arr_varCtl2(C_CNAM, lngX)).Visible = True Then
12840               blnCurrency = True
12850               If .Sort_line.Left = .Controls(arr_varCtl2(C_LIN, lngX)).Left Then blnSortHere = True
12860             End If
12870             .Controls(arr_varCtl2(C_CNAM, lngX)).Left = lngCurrID_Left
12880             .Controls(arr_varCtl2(C_LBL, lngX)).Left = lngCurrIDLbl_Left
12890             .Controls(arr_varCtl2(C_DMLB, lngX)).Left = (lngCurrIDLbl_Left + lngTpp)
12900             .Controls(arr_varCtl2(C_LIN, lngX)).Left = lngCurrID_Left
12910             .Controls(arr_varCtl2(C_DMLN, lngX)).Left = (lngCurrID_Left + lngTpp)
12920             If blnSortHere = True Then
12930               .Sort_line.Left = .Controls(arr_varCtl2(C_LIN, lngX)).Left
12940               .Sort_lbl.Left = ((.Controls(arr_varCtl2(C_LBL, lngX)).Left + .Controls(arr_varCtl2(C_LBL, lngX)).Width) - lngSortLbl_Width)
12950             End If
12960             blnSortHere = False
12970           Else
12980             Select Case blnFound
                  Case True
                    ' ** Controls to right of curr_id may move differently.
12990               Select Case blnCurrency
                    Case True
                      ' ** Add currency width.
13000                 If IsNull(arr_varCtl2(C_LBL, lngX)) = False Then
13010                   If .Sort_line.Left = .Controls(arr_varCtl2(C_LBL, lngX)).Left Then blnSortHere = True
13020                 End If
13030                 .Controls(arr_varCtl2(C_CNAM, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngCurrID_Width)
13040                 If IsNull(arr_varCtl2(C_LBL, lngX)) = False Then
13050                   .Controls(arr_varCtl2(C_LBL, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngCurrID_Width)
13060                   If IsNull(arr_varCtl2(C_DMLB, lngX)) = False Then
13070                     .Controls(arr_varCtl2(C_DMLB, lngX)).Left = ((arr_varCtl2(C_LFT, lngX) + lngCurrID_Width) + lngTpp)
13080                   End If
13090                 End If
13100                 If IsNull(arr_varCtl2(C_LBL2, lngX)) = False Then
13110                   .Controls(arr_varCtl2(C_LBL2, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngCurrID_Width)
13120                   If IsNull(arr_varCtl2(C_DML2, lngX)) = False Then
13130                     .Controls(arr_varCtl2(C_DML2, lngX)).Left = ((arr_varCtl2(C_LFT, lngX) + lngCurrID_Width) + lngTpp)
13140                   End If
13150                 End If
13160                 If IsNull(arr_varCtl2(C_LIN, lngX)) = False Then
13170                   .Controls(arr_varCtl2(C_LIN, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngCurrID_Width)
13180                   If IsNull(arr_varCtl2(C_DMLN, lngX)) = False Then
13190                     .Controls(arr_varCtl2(C_DMLN, lngX)).Left = ((arr_varCtl2(C_LFT, lngX) + lngCurrID_Width) + lngTpp)
13200                   End If
13210                 End If
13220               Case False
                      ' ** Their original positions are fine.
13230                 If IsNull(arr_varCtl2(C_LBL, lngX)) = False Then
13240                   If .Sort_line.Left = .Controls(arr_varCtl2(C_LBL, lngX)).Left Then blnSortHere = True
13250                 End If
13260                 .Controls(arr_varCtl2(C_CNAM, lngX)).Left = arr_varCtl2(C_LFT, lngX)
13270                 If IsNull(arr_varCtl2(C_LBL, lngX)) = False Then
13280                   .Controls(arr_varCtl2(C_LBL, lngX)).Left = arr_varCtl2(C_LFT, lngX)
13290                   If IsNull(arr_varCtl2(C_DMLB, lngX)) = False Then
13300                     .Controls(arr_varCtl2(C_DMLB, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngTpp)
13310                   End If
13320                 End If
13330                 If IsNull(arr_varCtl2(C_LBL2, lngX)) = False Then
13340                   .Controls(arr_varCtl2(C_LBL2, lngX)).Left = arr_varCtl2(C_LFT, lngX)
13350                   If IsNull(arr_varCtl2(C_DML2, lngX)) = False Then
13360                     .Controls(arr_varCtl2(C_DML2, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngTpp)
13370                   End If
13380                 End If
13390                 If IsNull(arr_varCtl2(C_LIN, lngX)) = False Then
13400                   .Controls(arr_varCtl2(C_LIN, lngX)).Left = arr_varCtl2(C_LFT, lngX)
13410                   If IsNull(arr_varCtl2(C_DMLN, lngX)) = False Then
13420                     .Controls(arr_varCtl2(C_DMLN, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngTpp)
13430                   End If
13440                 End If
13450               End Select
13460             Case False
                    ' ** Controls to left of curr_id move normally.
13470               If IsNull(arr_varCtl2(C_LBL, lngX)) = False Then
13480                 If .Sort_line.Left = .Controls(arr_varCtl2(C_LBL, lngX)).Left Then blnSortHere = True
13490               End If
13500               .Controls(arr_varCtl2(C_CNAM, lngX)).Left = arr_varCtl2(C_LFT, lngX)
13510               If IsNull(arr_varCtl2(C_LBL, lngX)) = False Then
13520                 .Controls(arr_varCtl2(C_LBL, lngX)).Left = arr_varCtl2(C_LFT, lngX)
13530                 If IsNull(arr_varCtl2(C_DMLB, lngX)) = False Then
13540                   .Controls(arr_varCtl2(C_DMLB, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngTpp)
13550                 End If
13560               End If
13570               If IsNull(arr_varCtl2(C_LBL2, lngX)) = False Then
13580                 .Controls(arr_varCtl2(C_LBL2, lngX)).Left = arr_varCtl2(C_LFT, lngX)
13590                 If IsNull(arr_varCtl2(C_DML2, lngX)) = False Then
13600                   .Controls(arr_varCtl2(C_DML2, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngTpp)
13610                 End If
13620               End If
13630               If IsNull(arr_varCtl2(C_LIN, lngX)) = False Then
13640                 .Controls(arr_varCtl2(C_LIN, lngX)).Left = arr_varCtl2(C_LFT, lngX)
13650                 If IsNull(arr_varCtl2(C_DMLN, lngX)) = False Then
13660                   .Controls(arr_varCtl2(C_DMLN, lngX)).Left = (arr_varCtl2(C_LFT, lngX) + lngTpp)
13670                 End If
13680               End If
13690             End Select
13700           End If
13710           If blnSortHere = True Then
13720             .Sort_line.Left = .Controls(arr_varCtl2(C_LBL, lngX)).Left
13730             .Sort_lbl.Left = ((.Controls(arr_varCtl2(C_LBL, lngX)).Left + .Controls(arr_varCtl2(C_LBL, lngX)).Width) - lngSortLbl_Width)
13740           End If
13750           blnSortHere = False
13760         Next  ' ** lngX.
13770         If blnResort = True Then
13780           blnResort = False
13790           .Sort_line.Visible = False
13800           .Sort_lbl.Visible = False
13810           SortNow "Form_Load"  ' ** Procedure: Below.
13820         End If
              ' ** Adjust form's width.
13830         Select Case blnCurrency
              Case True
13840           .Width = (lngForm_Width + lngCurrID_Width)
13850         Case False
13860           .Width = lngForm_Width
13870         End Select
13880       End If
13890     End Select
13900   End With

EXITP:
13910   Exit Sub

ERRH:
13920   Select Case ERR.Number
        Case Else
13930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13940   End Select
13950   Resume EXITP

End Sub

Public Sub Sub1_Disable(blnAble As Boolean)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "Sub1_Disable"
14010   With Me
14020     Select Case blnAble
          Case True
14030       .ppj_pospay_lbl.ForeColor = CLR_DKGRY2
14040       .ppj_pospay_lbl_dim_hi.Visible = False
14050       .ppj_pospay_lbl_line.BorderColor = CLR_DKGRY
14060       .ppj_pospay_lbl_line_dim_hi.Visible = False
14070       .Journal_ID_lbl.ForeColor = CLR_DKGRY2
14080       .Journal_ID_lbl_dim_hi.Visible = False
14090       .Journal_ID_lbl2.ForeColor = CLR_DKGRY2
14100       .Journal_ID_lbl2_dim_hi.Visible = False
14110       .Journal_ID_lbl_line.BorderColor = CLR_DKGRY
14120       .Journal_ID_lbl_line_dim_hi.Visible = False
14130       .journaltype_lbl.ForeColor = CLR_DKGRY2
14140       .journaltype_lbl_dim_hi.Visible = False
14150       .journaltype_lbl2.ForeColor = CLR_DKGRY2
14160       .journaltype_lbl2_dim_hi.Visible = False
14170       .journaltype_lbl_line.BorderColor = CLR_DKGRY
14180       .journaltype_lbl_line_dim_hi.Visible = False
14190       .PrintCheck_lbl.ForeColor = CLR_DKGRY2
14200       .PrintCheck_lbl_dim_hi.Visible = False
14210       .PrintCheck_lbl2.ForeColor = CLR_DKGRY2
14220       .PrintCheck_lbl2_dim_hi.Visible = False
14230       .PrintCheck_lbl_line.BorderColor = CLR_DKGRY
14240       .PrintCheck_lbl_line_dim_hi.Visible = False
14250       .CheckNum_lbl.ForeColor = CLR_DKGRY2
14260       .CheckNum_lbl_dim_hi.Visible = False
14270       .CheckNum_lbl2.ForeColor = CLR_DKGRY2
14280       .CheckNum_lbl2_dim_hi.Visible = False
14290       .CheckNum_lbl_line.BorderColor = CLR_DKGRY
14300       .CheckNum_lbl_line_dim_hi.Visible = False
14310       If .ppj_void.Visible = True Then
14320         .ppj_void_lbl.ForeColor = CLR_DKGRY2
14330         .ppj_void_lbl_dim_hi.Visible = False
14340         .ppj_void_lbl_line.BorderColor = CLR_DKGRY
14350         .ppj_void_lbl_line_dim_hi.Visible = False
14360       End If
14370       .accountno_lbl.ForeColor = CLR_DKGRY2
14380       .accountno_lbl_dim_hi.Visible = False
14390       .accountno_lbl2.ForeColor = CLR_DKGRY2
14400       .accountno_lbl2_dim_hi.Visible = False
14410       .accountno_lbl_line.BorderColor = CLR_DKGRY
14420       .accountno_lbl_line_dim_hi.Visible = False
14430       .shortname_lbl.ForeColor = CLR_DKGRY2
14440       .shortname_lbl_dim_hi.Visible = False
14450       .shortname_lbl_line.BorderColor = CLR_DKGRY
14460       .shortname_lbl_Line_dim_hi.Visible = False
14470       .transdate_lbl.ForeColor = CLR_DKGRY2
14480       .transdate_lbl_dim_hi.Visible = False
14490       .transdate_lbl2.ForeColor = CLR_DKGRY2
14500       .transdate_lbl2_dim_hi.Visible = False
14510       .transdate_lbl_line.BorderColor = CLR_DKGRY
14520       .transdate_lbl_line_dim_hi.Visible = False
14530       .icash_lbl.ForeColor = CLR_DKGRY2
14540       .icash_lbl_dim_hi.Visible = False
14550       .icash_lbl2.ForeColor = CLR_DKGRY2
14560       .icash_lbl2_dim_hi.Visible = False
14570       .icash_lbl_line.BorderColor = CLR_DKGRY
14580       .icash_lbl_line_dim_hi.Visible = False
14590       .pcash_lbl.ForeColor = CLR_DKGRY2
14600       .pcash_lbl_dim_hi.Visible = False
14610       .pcash_lbl2.ForeColor = CLR_DKGRY2
14620       .pcash_lbl2_dim_hi.Visible = False
14630       .pcash_lbl_line.BorderColor = CLR_DKGRY
14640       .pcash_lbl_line_dim_hi.Visible = False
14650       .amount_lbl.ForeColor = CLR_DKGRY2
14660       .amount_lbl_dim_hi.Visible = False
14670       .amount_lbl2.ForeColor = CLR_DKGRY2
14680       .amount_lbl2_dim_hi.Visible = False
14690       .amount_lbl_line.BorderColor = CLR_DKGRY
14700       .amount_lbl_line_dim_hi.Visible = False
14710       If .curr_id.Visible = True Then
14720         .curr_id_lbl.ForeColor = CLR_DKGRY2
14730         .curr_id_lbl_dim_hi.Visible = False
14740         .curr_id_lbl_line.BorderColor = CLR_DKGRY
14750         .curr_id_lbl_line_dim_hi.Visible = False
14760       End If
14770       .RecurringItem_lbl.ForeColor = CLR_DKGRY2
14780       .RecurringItem_lbl_dim_hi.Visible = False
14790       .RecurringItem_lbl_line.BorderColor = CLR_DKGRY
14800       .RecurringItem_lbl_line_dim_hi.Visible = False
14810       .JrnlMemo_Memo_lbl.ForeColor = CLR_DKGRY2
14820       .JrnlMemo_Memo_lbl_dim_hi.Visible = False
14830       .JrnlMemo_Memo_lbl_line.BorderColor = CLR_DKGRY
14840       .JrnlMemo_Memo_lbl_line_dim_hi.Visible = False
14850       .description_lbl.ForeColor = CLR_DKGRY2
14860       .description_lbl_dim_hi.Visible = False
14870       .description_lbl_line.BorderColor = CLR_DKGRY
14880       .description_lbl_line_dim_hi.Visible = False
14890       .Bank_Name_lbl.ForeColor = CLR_DKGRY2
14900       .Bank_Name_lbl_dim_hi.Visible = False
14910       .Bank_Name_lbl_line.BorderColor = CLR_DKGRY
14920       .Bank_Name_lbl_line_dim_hi.Visible = False
14930       .Bank_RoutingNumber_lbl.ForeColor = CLR_DKGRY2
14940       .Bank_RoutingNumber_lbl_dim_hi.Visible = False
14950       .Bank_RoutingNumber_lbl2.ForeColor = CLR_DKGRY2
14960       .Bank_RoutingNumber_lbl2_dim_hi.Visible = False
14970       .Bank_RoutingNumber_lbl_line.BorderColor = CLR_DKGRY
14980       .Bank_RoutingNumber_lbl_line_dim_hi.Visible = False
14990       .Bank_AccountNumber_lbl.ForeColor = CLR_DKGRY2
15000       .Bank_AccountNumber_lbl_dim_hi.Visible = False
15010       .Bank_AccountNumber_lbl2.ForeColor = CLR_DKGRY2
15020       .Bank_AccountNumber_lbl2_dim_hi.Visible = False
15030       .Bank_AccountNumber_lbl_line.BorderColor = CLR_DKGRY
15040       .Bank_AccountNumber_lbl_line_dim_hi.Visible = False
15050       .journal_USER_lbl.ForeColor = CLR_DKGRY2
15060       .journal_USER_lbl_dim_hi.Visible = False
15070       .journal_USER_lbl2.ForeColor = CLR_DKGRY2
15080       .journal_USER_lbl2_dim_hi.Visible = False
15090       .journal_USER_lbl_line.BorderColor = CLR_DKGRY
15100       .journal_USER_lbl_line_dim_hi.Visible = False
15110     Case False
15120       .ppj_pospay_lbl.ForeColor = WIN_CLR_DISF
15130       .ppj_pospay_lbl_dim_hi.Visible = True
15140       .ppj_pospay_lbl_line.BorderColor = WIN_CLR_DISR
15150       .ppj_pospay_lbl_line_dim_hi.Visible = True
15160       .Journal_ID_lbl.ForeColor = WIN_CLR_DISF
15170       .Journal_ID_lbl_dim_hi.Visible = True
15180       .Journal_ID_lbl2.ForeColor = WIN_CLR_DISF
15190       .Journal_ID_lbl2_dim_hi.Visible = True
15200       .Journal_ID_lbl_line.BorderColor = WIN_CLR_DISR
15210       .Journal_ID_lbl_line_dim_hi.Visible = True
15220       .journaltype_lbl.ForeColor = WIN_CLR_DISF
15230       .journaltype_lbl_dim_hi.Visible = True
15240       .journaltype_lbl2.ForeColor = WIN_CLR_DISF
15250       .journaltype_lbl2_dim_hi.Visible = True
15260       .journaltype_lbl_line.BorderColor = WIN_CLR_DISR
15270       .journaltype_lbl_line_dim_hi.Visible = True
15280       .PrintCheck_lbl.ForeColor = WIN_CLR_DISF
15290       .PrintCheck_lbl_dim_hi.Visible = True
15300       .PrintCheck_lbl2.ForeColor = WIN_CLR_DISF
15310       .PrintCheck_lbl2_dim_hi.Visible = True
15320       .PrintCheck_lbl_line.BorderColor = WIN_CLR_DISR
15330       .PrintCheck_lbl_line_dim_hi.Visible = True
15340       .CheckNum_lbl.ForeColor = WIN_CLR_DISF
15350       .CheckNum_lbl_dim_hi.Visible = True
15360       .CheckNum_lbl2.ForeColor = WIN_CLR_DISF
15370       .CheckNum_lbl2_dim_hi.Visible = True
15380       .CheckNum_lbl_line.BorderColor = WIN_CLR_DISR
15390       .CheckNum_lbl_line_dim_hi.Visible = True
15400       If .ppj_void.Visible = True Then
15410         .ppj_void_lbl.ForeColor = WIN_CLR_DISF
15420         .ppj_void_lbl_dim_hi.Visible = True
15430         .ppj_void_lbl_line.BorderColor = WIN_CLR_DISR
15440         .ppj_void_lbl_line_dim_hi.Visible = True
15450       End If
15460       .accountno_lbl.ForeColor = WIN_CLR_DISF
15470       .accountno_lbl_dim_hi.Visible = True
15480       .accountno_lbl2.ForeColor = WIN_CLR_DISF
15490       .accountno_lbl2_dim_hi.Visible = True
15500       .accountno_lbl_line.BorderColor = WIN_CLR_DISR
15510       .accountno_lbl_line_dim_hi.Visible = True
15520       .shortname_lbl.ForeColor = WIN_CLR_DISF
15530       .shortname_lbl_dim_hi.Visible = True
15540       .shortname_lbl_line.BorderColor = WIN_CLR_DISR
15550       .shortname_lbl_Line_dim_hi.Visible = True
15560       .transdate_lbl.ForeColor = WIN_CLR_DISF
15570       .transdate_lbl_dim_hi.Visible = True
15580       .transdate_lbl2.ForeColor = WIN_CLR_DISF
15590       .transdate_lbl2_dim_hi.Visible = True
15600       .transdate_lbl_line.BorderColor = WIN_CLR_DISR
15610       .transdate_lbl_line_dim_hi.Visible = True
15620       .icash_lbl.ForeColor = WIN_CLR_DISF
15630       .icash_lbl_dim_hi.Visible = True
15640       .icash_lbl2.ForeColor = WIN_CLR_DISF
15650       .icash_lbl2_dim_hi.Visible = True
15660       .icash_lbl_line.BorderColor = WIN_CLR_DISR
15670       .icash_lbl_line_dim_hi.Visible = True
15680       .pcash_lbl.ForeColor = WIN_CLR_DISF
15690       .pcash_lbl_dim_hi.Visible = True
15700       .pcash_lbl2.ForeColor = WIN_CLR_DISF
15710       .pcash_lbl2_dim_hi.Visible = True
15720       .pcash_lbl_line.BorderColor = WIN_CLR_DISR
15730       .pcash_lbl_line_dim_hi.Visible = True
15740       .amount_lbl.ForeColor = WIN_CLR_DISF
15750       .amount_lbl_dim_hi.Visible = True
15760       .amount_lbl2.ForeColor = WIN_CLR_DISF
15770       .amount_lbl2_dim_hi.Visible = True
15780       .amount_lbl_line.BorderColor = WIN_CLR_DISR
15790       .amount_lbl_line_dim_hi.Visible = True
15800       If .curr_id.Visible = True Then
15810         .curr_id_lbl.ForeColor = WIN_CLR_DISF
15820         .curr_id_lbl_dim_hi.Visible = True
15830         .curr_id_lbl_line.BorderColor = WIN_CLR_DISR
15840         .curr_id_lbl_line_dim_hi.Visible = True
15850       End If
15860       .RecurringItem_lbl.ForeColor = WIN_CLR_DISF
15870       .RecurringItem_lbl_dim_hi.Visible = True
15880       .RecurringItem_lbl_line.BorderColor = WIN_CLR_DISR
15890       .RecurringItem_lbl_line_dim_hi.Visible = True
15900       .JrnlMemo_Memo_lbl.ForeColor = WIN_CLR_DISF
15910       .JrnlMemo_Memo_lbl_dim_hi.Visible = True
15920       .JrnlMemo_Memo_lbl_line.BorderColor = WIN_CLR_DISR
15930       .JrnlMemo_Memo_lbl_line_dim_hi.Visible = True
15940       .description_lbl.ForeColor = WIN_CLR_DISF
15950       .description_lbl_dim_hi.Visible = True
15960       .description_lbl_line.BorderColor = WIN_CLR_DISR
15970       .description_lbl_line_dim_hi.Visible = True
15980       .Bank_Name_lbl.ForeColor = WIN_CLR_DISF
15990       .Bank_Name_lbl_dim_hi.Visible = True
16000       .Bank_Name_lbl_line.BorderColor = WIN_CLR_DISR
16010       .Bank_Name_lbl_line_dim_hi.Visible = True
16020       .Bank_RoutingNumber_lbl.ForeColor = WIN_CLR_DISF
16030       .Bank_RoutingNumber_lbl_dim_hi.Visible = True
16040       .Bank_RoutingNumber_lbl2.ForeColor = WIN_CLR_DISF
16050       .Bank_RoutingNumber_lbl2_dim_hi.Visible = True
16060       .Bank_RoutingNumber_lbl_line.BorderColor = WIN_CLR_DISR
16070       .Bank_RoutingNumber_lbl_line_dim_hi.Visible = True
16080       .Bank_AccountNumber_lbl.ForeColor = WIN_CLR_DISF
16090       .Bank_AccountNumber_lbl_dim_hi.Visible = True
16100       .Bank_AccountNumber_lbl2.ForeColor = WIN_CLR_DISF
16110       .Bank_AccountNumber_lbl2_dim_hi.Visible = True
16120       .Bank_AccountNumber_lbl_line.BorderColor = WIN_CLR_DISR
16130       .Bank_AccountNumber_lbl_line_dim_hi.Visible = True
16140       .journal_USER_lbl.ForeColor = WIN_CLR_DISF
16150       .journal_USER_lbl_dim_hi.Visible = True
16160       .journal_USER_lbl2.ForeColor = WIN_CLR_DISF
16170       .journal_USER_lbl2_dim_hi.Visible = True
16180       .journal_USER_lbl_line.BorderColor = WIN_CLR_DISR
16190       .journal_USER_lbl_line_dim_hi.Visible = True
16200     End Select
16210   End With

EXITP:
16220   Exit Sub

ERRH:
16230   Select Case ERR.Number
        Case Else
16240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16250   End Select
16260   Resume EXITP

End Sub

Public Function Recalc_NoRNAN() As Long

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "Recalc_NoRNAN"

        Dim varTmp00 As Variant
        Dim lngRetVal As Long

16310   With Me
16320     lngRetVal = 0&
16330     varTmp00 = .Bank_NoRNAN_Tot  ' ** Sum of a Boolean field.
16340     If IsNull(varTmp00) = False Then
16350       lngRetVal = Abs(varTmp00)
16360     End If
16370   End With

EXITP:
16380   Recalc_NoRNAN = lngRetVal
16390   Exit Function

ERRH:
16400   lngRetVal = 0&
16410   Select Case ERR.Number
        Case Else
16420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16430   End Select
16440   Resume EXITP

End Function

Public Function Recalc_CheckNum_NoNum() As Long

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "Recalc_CheckNum_NoNum"

        Dim varTmp00 As Variant
        Dim lngRetVal As Long

16510   With Me
16520     lngRetVal = 0&
16530     varTmp00 = .CheckNum_NoNum_Tot  ' ** Sum of a Boolean field.
16540     If IsNull(varTmp00) = False Then
16550       lngRetVal = Abs(varTmp00)
16560     End If
16570   End With

EXITP:
16580   Recalc_CheckNum_NoNum = lngRetVal
16590   Exit Function

ERRH:
16600   lngRetVal = 0&
16610   Select Case ERR.Number
        Case Else
16620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16630   End Select
16640   Resume EXITP

End Function

Public Function RecCnt() As Long

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

16710   lngRetVal = 0&

16720   With Me
16730     Set rst = .RecordsetClone
16740     With rst
16750       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
16760       Else
16770         .MoveLast
16780         lngRetVal = .RecordCount
16790       End If
16800       .Close
16810     End With
16820   End With

EXITP:
16830   Set rst = Nothing
16840   RecCnt = lngRetVal
16850   Exit Function

ERRH:
16860   lngRetVal = 0&
16870   Select Case ERR.Number
        Case Else
16880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16890   End Select
16900   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

17010   Select Case intWhere
        Case acCmdRecordsGoToFirst
17020     DoCmd.RunCommand acCmdRecordsGoToFirst
17030   Case acCmdRecordsGoToLast
17040     DoCmd.RunCommand acCmdRecordsGoToLast
17050   Case acCmdRecordsGoToPrevious
17060     DoCmd.RunCommand acCmdRecordsGoToPrevious
17070   Case acCmdRecordsGoToNext
17080     DoCmd.RunCommand acCmdRecordsGoToNext
17090   Case Else
17100     If IsMissing(varID) = False Then
17110       With Me
17120         Set rst = .RecordsetClone
17130         With rst
17140           .FindFirst "[ppj_id] = " & CStr(varID)
17150           If .NoMatch = False Then
17160             Me.Bookmark = .Bookmark
17170           End If
17180           .Close
17190         End With
17200       End With
17210     End If
17220   End Select

EXITP:
17230   Set rst = Nothing
17240   Exit Sub

ERRH:
17250   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
17260   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
17270   Case Else
17280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17290   End Select
17300   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String

        Const strStdAsc As String = ", [transdate], [accountno]"

17410   With Me
17420     If .Journal_ID_lbl_dim_hi.Visible = False Then
17430       If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
17440         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
17450       End If
17460       .Controls(strSortLbl).Visible = False
17470       .Controls(strSortLine).Visible = False
17480       .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
17490       If strProc = "Form_Load" Then
17500         strCalled = "Journal_ID"
17510         strSortNow = strSortOrig
17520         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
17530         lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - (lngSortLbl_Width)) + (8& * lngTpp))
17540         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
17550         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
17560         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
17570         .Controls(strSortLbl).Top = lngSortLbl_Top
17580         .Controls(strSortLbl).Left = lngSortLbl_Left
17590         .Controls(strSortLbl).Caption = strArwUp
17600         .Controls(strSortLine).Top = lngSortLine_Top
17610         .Controls(strSortLine).Left = lngSortLine_Left
17620         .Controls(strSortLine).Width = lngSortLine_Width
17630         .Controls(strSortLbl).ForeColor = CLR_DKBLU
17640       Else
17650         strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
17660         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
17670         lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - (lngSortLbl_Width))
17680         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
17690         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
17700         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
17710         .Controls(strSortLbl).Top = lngSortLbl_Top
17720         .Controls(strSortLbl).Left = lngSortLbl_Left
17730         .Controls(strSortLine).Top = lngSortLine_Top
17740         .Controls(strSortLine).Left = lngSortLine_Left
17750         .Controls(strSortLine).Width = lngSortLine_Width
17760         If strCalled = "Journal_ID" Then
17770           lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
17780           .Controls(strSortLbl).Left = lngSortLbl_Left
17790           strSortAsc = strSortOrig
17800           If strSortNow = strSortAsc Then
17810             strSortNow = "[Journal_ID] DESC"
17820             .Controls(strSortLbl).Caption = strArwDn
17830             .Controls(strSortLbl).ForeColor = CLR_DKRED
17840           Else
17850             strSortNow = strSortAsc
17860             .Controls(strSortLbl).Caption = strArwUp
17870             .Controls(strSortLbl).ForeColor = CLR_DKBLU
17880           End If
17890         Else
17900           .Controls(strSortLbl).Caption = strArwUp
17910           .Controls(strSortLbl).ForeColor = CLR_DKBLU
17920           Select Case strCalled
                Case "ppj_pospay"
17930             strSortAsc = "[ppj_pospay]" & strStdAsc
17940             lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
17950             .Controls(strSortLbl).Left = lngSortLbl_Left
17960           Case "journaltype"
17970             strSortAsc = "[journaltype]" & strStdAsc
17980             lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
17990             lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
18000             lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
18010             .Controls(strSortLbl).Left = lngSortLbl_Left
18020             .Controls(strSortLine).Left = lngSortLine_Left
18030             .Controls(strSortLine).Width = lngSortLine_Width
18040           Case "PrintCheck"
18050             strSortAsc = "[PrintCheck]" & strStdAsc
18060             lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
18070             .Controls(strSortLbl).Left = lngSortLbl_Left
18080           Case "CheckNum"
18090             strSortAsc = "[CheckNum]" & strStdAsc
18100           Case "ppj_void"
18110             strSortAsc = "[ppj_void]" & strStdAsc
18120             lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
18130             .Controls(strSortLbl).Left = lngSortLbl_Left
18140           Case "accountno"
18150             strSortAsc = "[accountno], [transdate], [RecurringItem]"
18160           Case "shortname"
18170             strSortAsc = "[shortname], [transdate], [RecurringItem]"
18180           Case "transdate"
18190             strSortAsc = "[transdate]" & strStdAsc
18200           Case "icash"
18210             strSortAsc = "[icash]" & strStdAsc
18220           Case "pcash"
18230             strSortAsc = "[pcash]" & strStdAsc
18240           Case "amount"
18250             strSortAsc = "[amount]" & strStdAsc
18260           Case "curr_id"
18270             strSortAsc = "[curr_code]" & strStdAsc
18280           Case "RecurringItem"
18290             strSortAsc = "[RecurringItem]" & strStdAsc
18300           Case "JrnlMemo_Memo"
18310             strSortAsc = "[JrnlMemo_Memo]" & strStdAsc
18320           Case "description"
18330             strSortAsc = "[description]" & strStdAsc
18340           Case "Bank_Name"
18350             strSortAsc = "[Bank_Name]" & strStdAsc
18360           Case "Bank_RoutingNumber"
18370             strSortAsc = "[Bank_RoutingNumber]" & strStdAsc
18380             lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
18390             .Controls(strSortLbl).Left = lngSortLbl_Left
18400           Case "Bank_AccountNumber"
18410             strSortAsc = "[Bank_AccountNumber]" & strStdAsc
18420             lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
18430             .Controls(strSortLbl).Left = lngSortLbl_Left
18440           Case "journal_USER"
18450             strSortAsc = "[journal_USER]" & strStdAsc
18460           End Select
18470           If strSortNow = strSortAsc Then
18480             intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
18490             Select Case intCnt
                  Case 1
18500               strTmp01 = strSortAsc & " DESC"
18510             Case 2
18520               intPos01 = InStr(strSortAsc, ",")
18530               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
18540               strTmp02 = Mid(strSortAsc, intPos01) & " DESC"
18550               strTmp01 = strTmp01 & strTmp02
18560             Case 3
18570               intPos01 = InStr(strSortAsc, ",")
18580               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
18590               strTmp03 = Mid(strSortAsc, intPos01)
18600               intPos01 = InStr(2, strTmp03, ",")
18610               strTmp02 = Left(strTmp03, (intPos01 - 1)) & " DESC"
18620               strTmp03 = Mid(strTmp03, intPos01) & " DESC"
18630               strTmp01 = strTmp01 & strTmp02 & strTmp03
18640             Case 4
18650               intPos01 = InStr(strSortAsc, ",")
18660               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
18670               strTmp03 = Mid(strSortAsc, intPos01)
18680               intPos01 = InStr(2, strTmp03, ",")
18690               strTmp02 = Left(strTmp03, (intPos01 - 1)) & " DESC"
18700               strTmp01 = strTmp01 & strTmp02
18710               strTmp03 = Mid(strTmp03, intPos01)
18720               intPos01 = InStr(2, strTmp03, ",")
18730               strTmp02 = Left(strTmp03, (intPos01 - 1)) & " DESC"
18740               strTmp03 = Mid(strTmp03, intPos01) & " DESC"
18750               strTmp01 = strTmp01 & strTmp02 & strTmp03
18760             End Select
18770             strSortNow = strTmp01
18780             .Controls(strSortLbl).Caption = strArwDn
18790             .Controls(strSortLbl).ForeColor = CLR_DKRED
18800           Else
18810             strSortNow = strSortAsc
18820           End If
18830         End If
18840       End If
18850       .Controls(strSortLbl).Visible = True
18860       .Controls(strSortLine).Visible = True
18870       .OrderBy = strSortNow
18880       .OrderByOn = True
18890     End If
18900   End With

        ' ** Potential Sorts:
        ' **   [Journal_ID]
        ' **   [ppj_pospay], [transdate], [accountno]
        ' **   [journaltype], [transdate], [accountno]
        ' **   [PrintCheck], [transdate], [accountno]
        ' **   [CheckNum], [transdate], [accountno]
        'Check_Number
        ' **   [ppj_void], [transdate], [accountno]
        'Issue_Type
        ' **   [accountno], [transdate], [RecurringItem]
        ' **   [shortname], [transdate], [RecurringItem]
        ' **   [transdate], [transdate], [accountno]
        'Issue_Date
        ' **   [icash], [transdate], [accountno]
        ' **   [pcash], [transdate], [accountno]
        ' **   [amount], [transdate], [accountno]
        'Amount
        ' **   [curr_code], [transdate], [accountno]
        ' **   [RecurringItem], [transdate], [accountno]
        'Payee_Name
        ' **   [JrnlMemo_Memo], [transdate], [accountno]
        ' **   [description], [transdate], [accountno]
        ' **   [Bank_Name], [transdate], [accountno]
        ' **   [Bank_RoutingNumber], [transdate], [accountno]
        'ABA_TRC
        ' **   [Bank_AccountNumber], [transdate], [accountno]
        'Account
        ' **   [journal_USER], [transdate], [accountno]

        'Debit_Credit
        'Issue_Action

EXITP:
18910   Exit Sub

ERRH:
18920   Select Case ERR.Number
        Case Else
18930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18940   End Select
18950   Resume EXITP

End Sub

Public Function SortNow_Get() As String

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim blnIsDesc As Boolean
        Dim strTmp01 As String
        Dim intPos01 As Integer
        Dim strRetVal As String

        Const strStdAsc As String = ", [Issue_Date], [Account]"
        Const strStdAscD As String = ", [Issue_Date] DESC, [Account] DESC"

19010   blnIsDesc = False

19020   If strSortNow = vbNullString Then
19030     strRetVal = strSortOrig
19040   Else
19050     strRetVal = strSortNow
19060   End If

19070   If InStr(strRetVal, " DESC") > 0 Then
19080     blnIsDesc = True
19090   End If

19100   intPos01 = InStr(strRetVal, "]")
19110   strTmp01 = Left(strRetVal, intPos01)
19120   Select Case strTmp01
        Case "[CheckNum]"
          ' **   [CheckNum], [transdate], [accountno]
19130     Select Case blnIsDesc
          Case True
19140       strRetVal = "[Check_Number] DESC" & strStdAscD
19150     Case False
19160       strRetVal = "[Check_Number]" & strStdAsc
19170     End Select
19180   Case "[ppj_void]"
          ' **   [ppj_void], [transdate], [accountno]
19190     Select Case blnIsDesc
          Case True
19200       strRetVal = "[Issue_Type] DESC" & strStdAscD
19210     Case False
19220       strRetVal = "[Issue_Type]" & strStdAsc
19230     End Select
19240   Case "[transdate]"
          ' **   [transdate], [transdate], [accountno]
19250     Select Case blnIsDesc
          Case True
19260       strRetVal = "[Issue_Date] DESC" & strStdAscD
19270     Case False
19280       strRetVal = "[Issue_Date]" & strStdAsc
19290     End Select
19300   Case "[amount]"
          ' **   [amount], [transdate], [accountno]
19310     Select Case blnIsDesc
          Case True
19320       strRetVal = "[Amount] DESC" & strStdAscD
19330     Case False
19340       strRetVal = "[Amount]" & strStdAsc
19350     End Select
19360   Case "[RecurringItem]"
          ' **   [RecurringItem], [transdate], [accountno]
19370     Select Case blnIsDesc
          Case True
19380       strRetVal = "[Payee_Name] DESC" & strStdAscD
19390     Case False
19400       strRetVal = "[Payee_Name]" & strStdAsc
19410     End Select
19420   Case "[Bank_RoutingNumber]"
          ' **   [Bank_RoutingNumber], [transdate], [accountno]
19430     Select Case blnIsDesc
          Case True
19440       strRetVal = "[ABA_TRC] DESC" & strStdAscD
19450     Case False
19460       strRetVal = "[ABA_TRC]" & strStdAsc
19470     End Select
19480   Case "[Bank_AccountNumber]"
          ' **   [Bank_AccountNumber], [transdate], [accountno]
19490     Select Case blnIsDesc
          Case True
19500       strRetVal = "[Account] DESC" & strStdAscD
19510     Case False
19520       strRetVal = "[Account]" & strStdAsc
19530     End Select
19540   Case Else
19550     Select Case strTmp01
          Case "[accountno]", [shortname]
            ' **   [accountno], [transdate], [RecurringItem]
            ' **   [shortname], [transdate], [RecurringItem]
19560       Select Case blnIsDesc
            Case True
19570         strRetVal = "[Account] DESC, [Issue_Date] DESC, [Payee_Name] DESC"
19580       Case False
19590         strRetVal = "[Account], [Issue_Date], [Payee_Name]"
19600       End Select
19610     Case "[icash]", "[pcash]"
            ' **   [icash], [transdate], [accountno]
            ' **   [pcash], [transdate], [accountno]
19620       Select Case blnIsDesc
            Case True
19630         strRetVal = "[Amount] DESC" & strStdAscD
19640       Case False
19650         strRetVal = "[Amount]" & strStdAsc
19660       End Select
19670     Case "[JrnlMemo_Memo]", "[description]"
            ' **   [JrnlMemo_Memo], [transdate], [accountno]
            ' **   [description], [transdate], [accountno]
19680       Select Case blnIsDesc
            Case True
19690         strRetVal = "[Payee_Name] DESC" & strStdAscD
19700       Case False
19710         strRetVal = "[Payee_Name]" & strStdAsc
19720       End Select
19730     Case "[Bank_Name]"
            ' **   [Bank_Name], [transdate], [accountno]
19740       Select Case blnIsDesc
            Case True
19750         strRetVal = "[ABA_TRC] DESC" & strStdAscD
19760       Case False
19770         strRetVal = "[ABA_TRC]" & strStdAsc
19780       End Select
19790     Case "[Journal_ID]", "[ppj_pospay]", "[journaltype]", "[PrintCheck]", "[curr_code]"
            ' **   [Journal_ID]
            ' **   [ppj_pospay], [transdate], [accountno]
            ' **   [journaltype], [transdate], [accountno]
            ' **   [PrintCheck], [transdate], [accountno]
19800       Select Case blnIsDesc
            Case True
19810         strRetVal = "[Issue_Date] DESC, [Account] DESC"
19820       Case False
19830         strRetVal = "[Issue_Date], [Account]"
19840       End Select
19850     End Select
19860   End Select

        ' ** Fields not in subform:
        ' **   [Debit_Credit]
        ' **   [Issue_Action]

        ' ** Fields not on report:
        ' **   [Journal_ID]
        ' **   [ppj_pospay], [transdate], [accountno]
        ' **   [journaltype], [transdate], [accountno]
        ' **   [PrintCheck], [transdate], [accountno]
        ' **   [accountno]
        ' **   [shortname]
        ' **   [icash], [transdate], [accountno]
        ' **   [pcash], [transdate], [accountno]
        ' **   [curr_code], [transdate], [accountno]
        ' **   [JrnlMemo_Memo], [transdate], [accountno]
        ' **   [description], [transdate], [accountno]
        ' **   [Bank_Name], [transdate], [accountno]

EXITP:
19870   SortNow_Get = strRetVal
19880   Exit Function

ERRH:
19890   strRetVal = RET_ERR
19900   Select Case ERR.Number
        Case Else
19910     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19920   End Select
19930   Resume EXITP

End Function
