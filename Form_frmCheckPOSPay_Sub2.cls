VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckPOSPay_Sub2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckPOSPay_Sub2"

'VGC 03/25/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Issue Date:       A {Issue_Date on frmCheckPOSPay}
' **   Check Number:     B {ppd_checknum on frmCheckPOSPay_Sub3}  ' ** Not visible when here.
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
' **   Void:             V {ppd_void on frmCheckPOSPay_Sub3}  ' ** Not visible when here.
' **   Exit:             X {cmdClose on frmCheckPOSPay}
' **   Payee:            Y {RecurringItem on frmCheckPOSPay_Sub1}

'I CAN'T FIGURE OUT WHAT TO GIVE THESE!
' **   Description:      P {pp_description}
' **   Payee:            P {ppd_payee on frmCheckPOSPay_Sub3}  ' ** Not visible when here.

' ** Shortcut Ctrl keys responsive from this form:
' **   Next:             N {MoveRec}
' **   Previous:         P {MoveRec}
' **   Save:             S {cmdSave on frmCheckPOSPay_Sub1}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[pp_id] DESC"
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

440       .pp_id.ForeColor = CLR_DISABLED_FG
450       .pp_id.BackColor = CLR_DISABLED_BG
460       .pp_date.ForeColor = CLR_DISABLED_FG
470       .pp_date.BackColor = CLR_DISABLED_BG
480       .pp_checks.ForeColor = CLR_DISABLED_FG
490       .pp_checks.BackColor = CLR_DISABLED_BG
500       .pp_description.ForeColor = CLR_DISABLED_FG
510       .pp_description.BackColor = CLR_DISABLED_BG
520       .pp_pathfile.ForeColor = CLR_DISABLED_FG
530       .pp_pathfile.BackColor = CLR_DISABLED_BG
540       .Username.ForeColor = CLR_DISABLED_FG
550       .Username.BackColor = CLR_DISABLED_BG
560       .pp_datemodified.ForeColor = CLR_DISABLED_FG
570       .pp_datemodified.BackColor = CLR_DISABLED_BG

580       SortNow THIS_PROC  ' ** Procedure: Below.

590     End With

EXITP:
600     Exit Sub

ERRH:
610     DoCmd.Hourglass False
620     Select Case ERR.Number
        Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Sub

Private Sub Form_Current()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

710     With Me
720       .Parent.pp_id = .pp_id
730       Select Case IsNull(.pp_id)
          Case True
740         .Parent.pp_date = Null
750         .Parent.pp_description = Null
760       Case False
770         .Parent.pp_date = .pp_date
780         .Parent.pp_description = .pp_description
790       End Select
800     End With

EXITP:
810     Exit Sub

ERRH:
820     DoCmd.Hourglass False
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Check Number:     K {CheckNum on frmCheckPOSPay_Sub1}
        ' **   Check Memo:       M {JrnlMemo_Memo on frmCheckPOSPay_Sub1}
        ' **   Comments:         O {description on frmCheckPOSPay_Sub1}
        ' **   Currency:         R {curr_id on frmCheckPOSPay_Sub1}
        ' **   Posting Date:     S {transdate on frmCheckPOSPay_Sub1}
        ' **   Exit:             X {cmdClose on frmCheckPOSPay}
        ' **   Payee:            Y {RecurringItem on frmCheckPOSPay_Sub1}

        'I CAN'T FIGURE OUT WHAT TO GIVE THIS!
        ' **   Description:      P {pp_description}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next:             N {MoveRec}
        ' **   Previous:         P {MoveRec}
        ' **   Save:             S {cmdSave on frmCheckPOSPay_Sub1}

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyUp
970         With Me
980           intRetVal = 0
990           MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1000        End With
1010      Case vbKeyDown
1020        With Me
1030          intRetVal = 0
1040          MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1050        End With
1060      End Select
1070    End If

        ' ** Alt keys.
1080    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1090      Select Case intRetVal
          Case vbKeyK
1100        With Me
1110          intRetVal = 0
1120          DoCmd.SelectObject acForm, .Parent.Name, False
1130          .Parent.frmCheckPOSPay_Sub1.SetFocus
1140          .Parent.frmCheckPOSPay_Sub1.Form.CheckNum.SetFocus
1150        End With
1160      Case vbKeyM
1170        With Me
1180          intRetVal = 0
1190          DoCmd.SelectObject acForm, .Parent.Name, False
1200          .Parent.frmCheckPOSPay_Sub1.SetFocus
1210          .Parent.frmCheckPOSPay_Sub1.Form.JrnlMemo_Memo.SetFocus
1220        End With
1230      Case vbKeyO
1240        With Me
1250          intRetVal = 0
1260          DoCmd.SelectObject acForm, .Parent.Name, False
1270          .Parent.frmCheckPOSPay_Sub1.SetFocus
1280          .Parent.frmCheckPOSPay_Sub1.Form.description.SetFocus
1290        End With
'1300      Case vbKeyP
'1310        With Me
'1320          intRetVal = 0
'1330          .pp_description.SetFocus
'1340        End With
1300      Case vbKeyR
1310        With Me
1320          intRetVal = 0
1330          DoCmd.SelectObject acForm, .Parent.Name, False
1340          .Parent.frmCheckPOSPay_Sub1.SetFocus
1350          .Parent.frmCheckPOSPay_Sub1.Form.curr_id.SetFocus
1360        End With
1370      Case vbKeyS
1380        With Me
1390          intRetVal = 0
1400          DoCmd.SelectObject acForm, .Parent.Name, False
1410          .Parent.frmCheckPOSPay_Sub1.SetFocus
1420          .Parent.frmCheckPOSPay_Sub1.Form.transdate.SetFocus
1430        End With
1440      Case vbKeyX
1450        With Me
1460          intRetVal = 0
1470          .Parent.cmdClose_Click  ' ** Form Procedure: frmCheckPOSPay.
1480        End With
1490      Case vbKeyY
1500        With Me
1510          intRetVal = 0
1520          DoCmd.SelectObject acForm, .Parent.Name, False
1530          .Parent.frmCheckPOSPay_Sub1.SetFocus
1540          .Parent.frmCheckPOSPay_Sub1.Form.RecurringItem.SetFocus
1550        End With
1560      End Select
1570    End If

        ' ** Ctrl keys.
1580    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1590      Select Case intRetVal
          Case vbKeyN
1600        With Me
1610          intRetVal = 0
1620          MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1630        End With
1640      Case vbKeyP
1650        With Me
1660          intRetVal = 0
1670          MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1680        End With
1690      Case vbKeyS
1700        With Me
1710          intRetVal = 0
              'DoCmd.SelectObject acForm, .Parent.Name, False
              '.Parent.frmCheckPOSPay_Sub1.SetFocus
              '.Parent.frmCheckPOSPay_Sub1.Form.cmdSave_Click  ' ** Form Procedure: frmCheckPOSPay_Sub1.
1720        End With
1730      Case vbKeyTab, vbKeyReturn
1740        With Me
1750          intRetVal = 0
1760          DoCmd.SelectObject acForm, .Parent.Name, False
1770          .Parent.opgView.SetFocus
1780        End With
1790      Case vbKeyPageUp
1800        intRetVal = 0
1810        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1820      Case vbKeyPageDown
1830        intRetVal = 0
1840        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1850      End Select
1860    End If

        ' ** Ctrl-Shift keys.
1870    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1880      Select Case intRetVal
          Case vbKeyF
1890        With Me
1900          intRetVal = 0
1910          DoCmd.SelectObject acForm, .Parent.Name, False
1920          .Parent.FocusHolder.SetFocus
1930        End With
1940      Case vbKeyTab, vbKeyReturn
1950        With Me
1960          intRetVal = 0
1970          DoCmd.SelectObject acForm, .Parent.Name, False
1980          If .Parent.cmdUpdate.Enabled = True Then
1990            .Parent.cmdUpdate.SetFocus
2000          ElseIf .Parent.cmdDelete.Enabled = True Then
2010            .Parent.cmdDelete.SetFocus
2020          ElseIf .Parent.cmdEdit.Enabled = True Then
2030            .Parent.cmdEdit.SetFocus
2040          ElseIf .Parent.cmdGeneratePOSPay.Enabled = True Then
2050            .Parent.cmdGeneratePOSPay.SetFocus
2060          Else
2070            lngRecsCur = .Parent.frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
2080            If lngRecsCur > 0& Then
2090              .Parent.frmCheckPOSPay_Sub1.SetFocus
2100            Else
2110              .Parent.cmdClose.SetFocus
2120            End If
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
2210    Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
2220    Case Else
2230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2240    End Select
2250    Resume EXITP

End Sub

Private Sub pp_id_lbl_DblClick(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "pp_id_lbl_DblClick"

2310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub pp_id_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "pp_id_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyTab
2470        With Me
2480          intRetVal = 0
2490          .pp_date.SetFocus
2500        End With
2510      End Select
2520    End If

        ' ** Shift keys.
2530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2540      Select Case intRetVal
          Case vbKeyTab
2550        With Me
2560          intRetVal = 0
2570          If .CurrentRecord > 1 Then
2580            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2590            .FocusHolder2.SetFocus
2600            DoEvents
2610            .pp_datemodified.SetFocus
2620          Else
2630            DoCmd.SelectObject acForm, .Parent.Name, False
2640            If .Parent.cmdUpdate.Enabled = True Then
2650              .Parent.cmdUpdate.SetFocus
2660            ElseIf .Parent.cmdDelete.Enabled = True Then
2670              .Parent.cmdDelete.SetFocus
2680            ElseIf .Parent.cmdEdit.Enabled = True Then
2690              .Parent.cmdEdit.SetFocus
2700            ElseIf .Parent.cmdGeneratePOSPay.Enabled = True Then
2710              .Parent.cmdGeneratePOSPay.SetFocus
2720            Else
2730              lngRecsCur = .Parent.frmCheckPOSPay_Sub1.Form.RecCnt  ' ** Form Function: frmCheckPOSPay_Sub1.
2740              If lngRecsCur > 0& Then
2750                .Parent.frmCheckPOSPay_Sub1.SetFocus
2760              Else
2770                .Parent.cmdClose.SetFocus
2780              End If
2790            End If
2800          End If
2810        End With
2820      End Select
2830    End If

EXITP:
2840    KeyCode = intRetVal
2850    Exit Sub

ERRH:
2860    intRetVal = 0
2870    Select Case ERR.Number
        Case Else
2880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2890    End Select
2900    Resume EXITP

End Sub

Private Sub pp_date_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "pp_date_lbl_DblClick"

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

Private Sub pp_checks_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "pp_checks_lbl_DblClick"

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

Private Sub pp_description_lbl_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "pp_description_lbl_DblClick"

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

Private Sub pp_pathfile_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "pp_pathfile_lbl_DblClick"

3310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub Username_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Username_lbl_DblClick"

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

Private Sub pp_datemodified_lbl_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "pp_datemodified_lbl_DblClick"

3510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub pp_datemodified_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "pp_datemodified_KeyDown"

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
          Case vbKeyTab
3670        With Me
3680          intRetVal = 0
3690          lngRecsCur = RecCnt  ' ** Function: Below.
3700          If .CurrentRecord < lngRecsCur Then
3710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3720            .pp_id.SetFocus
3730          Else
3740            DoCmd.SelectObject acForm, .Parent.Name, False
3750            .Parent.opgView.SetFocus
3760          End If
3770        End With
3780      End Select
3790    End If

        ' ** Shift keys.
3800    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3810      Select Case intRetVal
          Case vbKeyTab
3820        With Me
3830          intRetVal = 0
3840          .Username.SetFocus
3850        End With
3860      End Select
3870    End If

EXITP:
3880    KeyCode = intRetVal
3890    Exit Sub

ERRH:
3900    intRetVal = 0
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Sub

Public Sub Sub2_Disable(blnAble As Boolean)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Sub2_Disable"

4010    With Me
4020      Select Case blnAble
          Case True
4030        .pp_id_lbl.ForeColor = CLR_DKGRY2
4040        .pp_id_lbl_dim_hi.Visible = False
4050        .pp_id_lbl_line.BorderColor = CLR_DKGRY
4060        .pp_id_lbl_line_dim_hi.Visible = False
4070        .pp_date_lbl.ForeColor = CLR_DKGRY2
4080        .pp_date_lbl_dim_hi.Visible = False
4090        .pp_date_lbl_line.BorderColor = CLR_DKGRY
4100        .pp_date_lbl_line_dim_hi.Visible = False
4110        .pp_checks_lbl.ForeColor = CLR_DKGRY2
4120        .pp_checks_lbl_dim_hi.Visible = False
4130        .pp_checks_lbl2.ForeColor = CLR_DKGRY2
4140        .pp_checks_lbl2_dim_hi.Visible = False
4150        .pp_checks_lbl_line.BorderColor = CLR_DKGRY
4160        .pp_checks_lbl_line_dim_hi.Visible = False
4170        .pp_description_lbl.ForeColor = CLR_DKGRY2
4180        .pp_description_lbl_dim_hi.Visible = False
4190        .pp_description_lbl_line.BorderColor = CLR_DKGRY
4200        .pp_description_lbl_line_dim_hi.Visible = False
4210        .pp_pathfile_lbl.ForeColor = CLR_DKGRY2
4220        .pp_pathfile_lbl_dim_hi.Visible = False
4230        .pp_pathfile_lbl_line.BorderColor = CLR_DKGRY
4240        .pp_pathfile_lbl_line_dim_hi.Visible = False
4250        .Username_lbl.ForeColor = CLR_DKGRY2
4260        .Username_lbl_dim_hi.Visible = False
4270        .Username_lbl_line.BorderColor = CLR_DKGRY
4280        .Username_lbl_line_dim_hi.Visible = False
4290        .pp_datemodified_lbl.ForeColor = CLR_DKGRY2
4300        .pp_datemodified_lbl_dim_hi.Visible = False
4310        .pp_datemodified_lbl_line.BorderColor = CLR_DKGRY
4320        .pp_datemodified_lbl_line_dim_hi.Visible = False
4330      Case False
4340        .pp_id_lbl.ForeColor = WIN_CLR_DISF
4350        .pp_id_lbl_dim_hi.Visible = True
4360        .pp_id_lbl_line.BorderColor = WIN_CLR_DISR
4370        .pp_id_lbl_line_dim_hi.Visible = True
4380        .pp_date_lbl.ForeColor = WIN_CLR_DISF
4390        .pp_date_lbl_dim_hi.Visible = True
4400        .pp_date_lbl_line.BorderColor = WIN_CLR_DISR
4410        .pp_date_lbl_line_dim_hi.Visible = True
4420        .pp_checks_lbl.ForeColor = WIN_CLR_DISF
4430        .pp_checks_lbl_dim_hi.Visible = True
4440        .pp_checks_lbl2.ForeColor = WIN_CLR_DISF
4450        .pp_checks_lbl2_dim_hi.Visible = True
4460        .pp_checks_lbl_line.BorderColor = WIN_CLR_DISR
4470        .pp_checks_lbl_line_dim_hi.Visible = True
4480        .pp_description_lbl.ForeColor = WIN_CLR_DISF
4490        .pp_description_lbl_dim_hi.Visible = True
4500        .pp_description_lbl_line.BorderColor = WIN_CLR_DISR
4510        .pp_description_lbl_line_dim_hi.Visible = True
4520        .pp_pathfile_lbl.ForeColor = WIN_CLR_DISF
4530        .pp_pathfile_lbl_dim_hi.Visible = True
4540        .pp_pathfile_lbl_line.BorderColor = WIN_CLR_DISR
4550        .pp_pathfile_lbl_line_dim_hi.Visible = True
4560        .Username_lbl.ForeColor = WIN_CLR_DISF
4570        .Username_lbl_dim_hi.Visible = True
4580        .Username_lbl_line.BorderColor = WIN_CLR_DISR
4590        .Username_lbl_line_dim_hi.Visible = True
4600        .pp_datemodified_lbl.ForeColor = WIN_CLR_DISF
4610        .pp_datemodified_lbl_dim_hi.Visible = True
4620        .pp_datemodified_lbl_line.BorderColor = WIN_CLR_DISR
4630        .pp_datemodified_lbl_line_dim_hi.Visible = True
4640      End Select
4650    End With

EXITP:
4660    Exit Sub

ERRH:
4670    Select Case ERR.Number
        Case Else
4680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4690    End Select
4700    Resume EXITP

End Sub

Public Function RecCnt() As Long

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

4810    lngRetVal = 0&

4820    With Me
4830      Set rst = .RecordsetClone
4840      With rst
4850        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
4860        Else
4870          .MoveLast
4880          lngRetVal = .RecordCount
4890        End If
4900        .Close
4910      End With
4920    End With

EXITP:
4930    Set rst = Nothing
4940    RecCnt = lngRetVal
4950    Exit Function

ERRH:
4960    lngRetVal = 0&
4970    Select Case ERR.Number
        Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

5110    Select Case intWhere
        Case acCmdRecordsGoToFirst
5120      DoCmd.RunCommand acCmdRecordsGoToFirst
5130    Case acCmdRecordsGoToLast
5140      DoCmd.RunCommand acCmdRecordsGoToLast
5150    Case acCmdRecordsGoToPrevious
5160      DoCmd.RunCommand acCmdRecordsGoToPrevious
5170    Case acCmdRecordsGoToNext
5180      DoCmd.RunCommand acCmdRecordsGoToNext
5190    Case Else
5200      If IsMissing(varID) = False Then
5210        With Me
5220          Set rst = .RecordsetClone
5230          With rst
5240            .FindFirst "[pp_id] = " & CStr(varID)
5250            If .NoMatch = False Then
5260              Me.Bookmark = .Bookmark
5270            End If
5280            .Close
5290          End With
5300        End With
5310      End If
5320    End Select

EXITP:
5330    Set rst = Nothing
5340    Exit Sub

ERRH:
5350    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5360    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5370    Case Else
5380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5390    End Select
5400    Resume EXITP

End Sub

Public Sub MoveRecX(intWhere As AcCommand)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRecX"

        Dim rst As DAO.Recordset
        Dim lngPPID As Long

5510    With Me
5520      If IsNull(.pp_id) = False Then
5530        lngPPID = .pp_id
5540        Set rst = .RecordsetClone
5550        rst.MoveFirst
5560        If rst![pp_id] = lngPPID Then
5570          Select Case intWhere
              Case acCmdRecordsGoToPrevious
                ' ** Nothing; can't move!
5580          Case acCmdRecordsGoToNext
5590            rst.MoveNext
5600            .Bookmark = rst.Bookmark
5610          End Select
5620        Else
5630          rst.FindFirst "[pp_id] = " & CStr(lngPPID)
5640          If rst.NoMatch = False Then
5650            Select Case intWhere
                Case acCmdRecordsGoToPrevious
5660              rst.MovePrevious
5670              .Bookmark = rst.Bookmark
5680            Case acCmdRecordsGoToNext
5690              rst.MoveNext
5700              .Bookmark = rst.Bookmark
5710            End Select
5720          End If
5730        End If
5740        rst.Close
5750      End If
5760    End With

EXITP:
5770    Set rst = Nothing
5780    Exit Sub

ERRH:
5790    Select Case ERR.Number
        Case 3021  ' **No current record.
          ' ** Ignore.
5800    Case Else
5810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5820    End Select
5830    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String

        Const strStdAsc As String = ", [pp_date], [pp_description]"

5910    With Me
5920      If .pp_id_lbl_dim_hi.Visible = False Then
5930        If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5940          lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5950        End If
5960        .Controls(strSortLbl).Visible = False
5970        .Controls(strSortLine).Visible = False
5980        .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
5990        If strProc = "Form_Load" Then
6000          strCalled = "pp_id"
6010          strSortNow = strSortOrig
6020          lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6030          lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - (lngSortLbl_Width))
6040          lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6050          lngSortLine_Left = .Controls(strCalled & "_lbl").Left
6060          lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6070          .Controls(strSortLbl).Top = lngSortLbl_Top
6080          .Controls(strSortLbl).Left = lngSortLbl_Left
6090          .Controls(strSortLbl).Caption = strArwDn
6100          .Controls(strSortLine).Top = lngSortLine_Top
6110          .Controls(strSortLine).Left = lngSortLine_Left
6120          .Controls(strSortLine).Width = lngSortLine_Width
6130          .Controls(strSortLbl).ForeColor = CLR_DKRED
6140        Else
6150          strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
6160          lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6170          lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - (lngSortLbl_Width))
6180          lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6190          lngSortLine_Left = .Controls(strCalled & "_lbl").Left
6200          lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6210          .Controls(strSortLbl).Top = lngSortLbl_Top
6220          .Controls(strSortLbl).Left = lngSortLbl_Left
6230          .Controls(strSortLine).Top = lngSortLine_Top
6240          .Controls(strSortLine).Left = lngSortLine_Left
6250          .Controls(strSortLine).Width = lngSortLine_Width
6260          If strCalled = "pp_id" Then
6270            strSortAsc = strSortOrig
6280            If strSortNow = strSortAsc Then
6290              strSortNow = "[pp_id]"
6300              .Controls(strSortLbl).Caption = strArwUp
6310              .Controls(strSortLbl).ForeColor = CLR_DKBLU
6320            Else
6330              strSortNow = strSortAsc
6340              .Controls(strSortLbl).Caption = strArwDn
6350              .Controls(strSortLbl).ForeColor = CLR_DKRED
6360            End If
6370          Else
6380            .Controls(strSortLbl).Caption = strArwUp
6390            .Controls(strSortLbl).ForeColor = CLR_DKBLU
6400            Select Case strCalled
                Case "pp_date"
6410              strSortAsc = "[pp_date]" & strStdAsc
6420              lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
6430              .Controls(strSortLbl).Left = lngSortLbl_Left
6440            Case "pp_checks"
6450              strSortAsc = "[pp_checks]" & strStdAsc
6460              lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
6470              .Controls(strSortLbl).Left = lngSortLbl_Left
6480            Case "pp_description"
6490              strSortAsc = "[pp_description]" & strStdAsc
6500            Case "pp_pathfile"
6510              strSortAsc = "[pp_pathfile]"
6520            Case "Username"
6530              strSortAsc = "[Username]"
6540              lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
6550              .Controls(strSortLbl).Left = lngSortLbl_Left
6560            Case "pp_datemodified"
6570              strSortAsc = "[pp_datemodified]" & strStdAsc
6580            End Select
6590            If strSortNow = strSortAsc Then
6600              intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
6610              Select Case intCnt
                  Case 1
6620                strTmp01 = strSortAsc & " DESC"
6630              Case 2
6640                intPos01 = InStr(strSortAsc, ",")
6650                strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6660                strTmp02 = Mid(strSortAsc, intPos01) & " DESC"
6670                strTmp01 = strTmp01 & strTmp02
6680              Case 3
6690                intPos01 = InStr(strSortAsc, ",")
6700                strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6710                strTmp03 = Mid(strSortAsc, intPos01)
6720                intPos01 = InStr(2, strTmp03, ",")
6730                strTmp02 = Left(strTmp03, (intPos01 - 1)) & " DESC"
6740                strTmp03 = Mid(strTmp03, intPos01) & " DESC"
6750                strTmp01 = strTmp01 & strTmp02 & strTmp03
6760              Case 4
6770                intPos01 = InStr(strSortAsc, ",")
6780                strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6790                strTmp03 = Mid(strSortAsc, intPos01)
6800                intPos01 = InStr(2, strTmp03, ",")
6810                strTmp02 = Left(strTmp03, (intPos01 - 1)) & " DESC"
6820                strTmp01 = strTmp01 & strTmp02
6830                strTmp03 = Mid(strTmp03, intPos01)
6840                intPos01 = InStr(2, strTmp03, ",")
6850                strTmp02 = Left(strTmp03, (intPos01 - 1)) & " DESC"
6860                strTmp03 = Mid(strTmp03, intPos01) & " DESC"
6870                strTmp01 = strTmp01 & strTmp02 & strTmp03
6880              End Select
6890              strSortNow = strTmp01
6900              .Controls(strSortLbl).Caption = strArwDn
6910              .Controls(strSortLbl).ForeColor = CLR_DKRED
6920            Else
6930              strSortNow = strSortAsc
6940            End If
6950          End If
6960        End If
6970        .Controls(strSortLbl).Visible = True
6980        .Controls(strSortLine).Visible = True
6990        .OrderBy = strSortNow
7000        .OrderByOn = True
7010      End If
7020    End With

EXITP:
7030    Exit Sub

ERRH:
7040    Select Case ERR.Number
        Case Else
7050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7060    End Select
7070    Resume EXITP

End Sub
