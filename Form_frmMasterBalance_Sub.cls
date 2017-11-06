VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMasterBalance_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMasterBalance_Sub"

'VGC 10/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Show Acct Master: A {chkShowAcctMast on frmMasterBalance}
' **   Close:            C {cmdClose on frmMasterBalance}
' **   Include Closed:   D {chkIncludeClosed}
' **   Excel:            E {cmdExcel on frmMasterBalance}
' **   Disc Only:        H {chkDiscrepancies on frmMasterBalance}
' **   Fix:              I {fix}
' **   Name:             M {shortname}
' **   Select None:      N {cmdSelectNone on frmMasterBalance}
' **   Word:             O {cmdWord on frmMasterBalance}
' **   Print:            P {cmdPrint on frmMasterBalance}
' **   Show Acct Sort:   R {chkShowAcctSort on frmMasterBalance}
' **   Select All:       S {cmdSelectAll on frmMasterBalance}
' **   Update:           U {cmdUpdate on frmMasterBalance}
' **   Preview:          W {cmdPreview on frmMasterBalance}
' **   Exit:             X {cmdClose on frmMasterBalance}
' **   Save Size/Pos:    Z {chkSaveSizePos on frmMasterBalance}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:           F5 {RecalcTots on frmMasterBalance}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[account_sort]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngForm_Width As Long, lngOffset As Long
Private lngICash_Left As Long, lngPCash_Left As Long, lngCost_Left As Long
Private lngICashLbl_Left As Long, lngPCashLbl_Left As Long, lngCostLbl_Left As Long
Private lngICashL_Left As Long, lngPCashL_Left As Long, lngCostL_Left As Long
Private lngICashLLbl_Left As Long, lngPCashLLbl_Left As Long, lngCostLLbl_Left As Long
Private lngIx_Left As Long, lngPx_Left As Long, lngCx_Left As Long
Private lngAccountNo_Left As Long, lngShortName_Left As Long, lngShortName_Width As Long, lngFix_Left As Long
Private lngAccountNoLbl_Left As Long, lngShortNameLbl_Left As Long, lngFixLbl_Left As Long
Private lngIsClosedBG_Width As Long

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngTpp As Long, lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private strCallingForm As String, lngRecsCur As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmMasterBalance" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

            ' ** Just get this out of the way.
190         .form_width_line.Left = 0&

200         lngForm_Width = .WhatsThisButton
210         lngOffset = (.accountno.Left - .account_sort.Left)
220         lngAccountNo_Left = .accountno.Left
230         lngAccountNoLbl_Left = .accountno_lbl.Left
240         lngShortName_Left = .shortname.Left
250         lngShortName_Width = .shortname.Width
260         lngShortNameLbl_Left = .shortname_lbl.Left
270         lngFix_Left = .fix.Left
280         lngFixLbl_Left = .Fix_lbl.Left
290         lngICash_Left = .ICash.Left
300         lngICashLbl_Left = .icash_lbl.Left
310         lngPCash_Left = .PCash.Left
320         lngPCashLbl_Left = .pcash_lbl.Left
330         lngCost_Left = .Cost.Left
340         lngCostLbl_Left = .cost_lbl.Left
350         lngICashL_Left = .icash_L.Left
360         lngICashLLbl_Left = .icash_L_lbl.Left
370         lngPCashL_Left = .pcash_L.Left
380         lngPCashLLbl_Left = .pcash_L_lbl.Left
390         lngCostL_Left = .cost_L.Left
400         lngCostLLbl_Left = .cost_L_lbl.Left
410         lngIx_Left = .Ix.Left
420         lngPx_Left = .Px.Left
430         lngCx_Left = .Cx.Left
440         lngIsClosedBG_Width = .IsClosed_bg1.Width

450         lngSortLbl_Width = .Sort_lbl.Width

460       Else
470         Cancel = -1
480       End If
490     End With

500     If Cancel = -1 Then
510       Beep
520       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
530     End If

EXITP:
540     Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me
620       .account_sort.ForeColor = CLR_DISABLED_FG
630       .account_sort.BackColor = CLR_DISABLED_BG
640       .accountno.ForeColor = CLR_DISABLED_FG
650       .accountno.BackColor = CLR_DISABLED_BG
660       .shortname.ForeColor = CLR_DISABLED_FG
670       .shortname.BackColor = CLR_DISABLED_BG
680       .ICash.ForeColor = CLR_DISABLED_FG
690       .ICash.BackColor = CLR_DISABLED_BG
700       .icash_L.ForeColor = CLR_DISABLED_FG
710       .icash_L.BackColor = CLR_DISABLED_BG
720       .PCash.ForeColor = CLR_DISABLED_FG
730       .PCash.BackColor = CLR_DISABLED_BG
740       .pcash_L.ForeColor = CLR_DISABLED_FG
750       .pcash_L.BackColor = CLR_DISABLED_BG
760       .Cost.ForeColor = CLR_DISABLED_FG
770       .Cost.BackColor = CLR_DISABLED_BG
780       .cost_L.ForeColor = CLR_DISABLED_FG
790       .cost_L.BackColor = CLR_DISABLED_BG
800       .Ix.ForeColor = CLR_RED
810       .Px.ForeColor = CLR_RED
820       .Cx.ForeColor = CLR_RED
830     End With

840     SortNow THIS_PROC

EXITP:
850     Exit Sub

ERRH:
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub Form_Current()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

910     With Me
920       Select Case .IsClosed
          Case True
930         .account_sort.BackColor = CLR_LTTEAL3
940         .accountno.BackColor = CLR_LTTEAL3
950         .shortname.BackColor = CLR_LTTEAL3
960         .ICash.BackColor = CLR_LTTEAL3
970         .icash_L.BackColor = CLR_LTTEAL3
980         .Ix.BackColor = CLR_LTTEAL3
990         .PCash.BackColor = CLR_LTTEAL3
1000        .pcash_L.BackColor = CLR_LTTEAL3
1010        .Px.BackColor = CLR_LTTEAL3
1020        .Cost.BackColor = CLR_LTTEAL3
1030        .cost_L.BackColor = CLR_LTTEAL3
1040        .Cx.BackColor = CLR_LTTEAL3
1050      Case False
1060        .account_sort.BackColor = CLR_DISABLED_BG
1070        .accountno.BackColor = CLR_DISABLED_BG
1080        .shortname.BackColor = CLR_DISABLED_BG
1090        .ICash.BackColor = CLR_DISABLED_BG
1100        .icash_L.BackColor = CLR_DISABLED_BG
1110        .Ix.BackColor = CLR_DISABLED_BG
1120        .PCash.BackColor = CLR_DISABLED_BG
1130        .pcash_L.BackColor = CLR_DISABLED_BG
1140        .Px.BackColor = CLR_DISABLED_BG
1150        .Cost.BackColor = CLR_DISABLED_BG
1160        .cost_L.BackColor = CLR_DISABLED_BG
1170        .Cx.BackColor = CLR_DISABLED_BG
1180      End Select
1190    End With

EXITP:
1200    Exit Sub

ERRH:
1210    Select Case ERR.Number
        Case Else
1220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1230    End Select
1240    Resume EXITP

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

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:           F5 {RecalcTots on frmMasterBalance}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Fix:              I {fix}
        ' **   Name:             M {shortname}
        ' **   Exit:             X {cmdClose on frmMasterBalance}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyF5
1370        With Me
1380          intRetVal = 0
1390          .Parent.RecalcTots  ' ** Form Procedure: frmMasterBalance.
1400        End With
1410      Case vbKeyDown
1420        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1430      Case vbKeyUp
1440        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1450      Case vbKeyEscape
1460        With Me
1470          intRetVal = 0
1480          .Parent.cmdClose_Click  ' ** Form Procedure: frmMasterBalance.
1490        End With
1500      End Select
1510    End If

        ' ** Alt keys.
1520    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1530      Select Case intRetVal
          Case vbKeyI
1540        With Me
1550          intRetVal = 0
1560          .fix.SetFocus
1570          .fix = Not (.fix)
1580        End With
1590      Case vbKeyM
1600        With Me
1610          intRetVal = 0
1620          .shortname.SetFocus
1630        End With
1640      Case vbKeyX
1650        With Me
1660          intRetVal = 0
1670          .Parent.cmdClose_Click  ' ** Form Procedure: frmMasterBalance.
1680        End With
1690      End Select
1700    End If

        ' ** Ctrl keys.
1710    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1720      Select Case intRetVal
          Case vbKeyG
1730        With Me
1740          intRetVal = 0
1750          If .Parent.chkPageOf.Enabled = True Then
1760            .Parent.chkPageOf.SetFocus
1770            .Parent.chkPageOf = (Not .Parent.chkPageOf)
1780            .Parent.chkPageOf_AfterUpdate  ' ** Form Procedure: frmMasterBalance.
1790          End If
1800        End With
1810      Case vbKeyS
1820        intRetVal = 0
1830        cmdSave_Click  ' ** Procedure: Below.
1840      Case vbKeyTab, vbKeyReturn
1850        With Me
1860          intRetVal = 0
1870          DoCmd.SelectObject acForm, .Parent.Name, False
1880          If .Parent.cmdSelectAll.Enabled = True Then
1890            .Parent.cmdSelectAll.SetFocus
1900          ElseIf .Parent.chkDiscrepancies.Enabled = True And .Parent.chkDiscrepancies.Visible = True Then
1910            .Parent.chkDiscrepancies.SetFocus
1920          ElseIf .Parent.chkIncludeClosed.Enabled = True And .Parent.chkIncludeClosed.Visible = True Then
1930            .Parent.chkIncludeClosed.SetFocus
1940          ElseIf .Parent.chkShowAcctSort.Enabled And .Parent.chkShowAcctSort.Visible = True Then
1950            .Parent.chkShowAcctSort.SetFocus
1960          ElseIf .Parent.chkShowAcctMast.Enabled = True And .Parent.chkShowAcctMast.Visible = True Then
1970            .Parent.chkShowAcctMast.SetFocus
1980          Else
1990            .Parent.chkPageOf.SetFocus
2000          End If
2010        End With
2020      Case vbKeyPageUp, vbKeyHome
2030        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2040      Case vbKeyPageDown, vbKeyEnd
2050        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2060      End Select
2070    End If

        ' ** Ctrl-Shift keys.
2080    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2090      Select Case intRetVal
          Case vbKeyF
2100        With Me
2110          intRetVal = 0
2120          DoCmd.SelectObject acForm, .Parent.Name, False
2130          .Parent.FocusHolder.SetFocus
2140        End With
2150      Case vbKeyTab, vbKeyReturn
2160        With Me
2170          intRetVal = 0
2180          DoCmd.SelectObject acForm, .Parent.Name, False
2190          .Parent.cmdClose.SetFocus
2200        End With
2210      End Select
2220    End If

EXITP:
2230    KeyCode = intRetVal
2240    Exit Sub

ERRH:
2250    intRetVal = 0
2260    Select Case ERR.Number
        Case Else
2270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2280    End Select
2290    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2310    With Me
2320      If .Filter <> vbNullString Then
2330        .Filter = vbNullString
2340      End If
2350    End With

EXITP:
2360    Exit Sub

ERRH:
2370    Select Case ERR.Number
        Case Else
2380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2390    End Select
2400    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2510    DoCmd.RunCommand acCmdSelectRecord
2520    DoCmd.RunCommand acCmdSaveRecord

EXITP:
2530    Exit Sub

ERRH:
2540    Select Case ERR.Number
        Case Else
2550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2560    End Select
2570    Resume EXITP

End Sub

Private Sub account_sort_lbl_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "account_sort_lbl_DblClick"

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

Private Sub account_sort_GotFocus()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "account_sort_GotFocus"

2710    With Me
2720      .account_sort.SelLength = 0
2730      .account_sort.SelStart = 99
2740    End With

EXITP:
2750    Exit Sub

ERRH:
2760    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
2770    Case Else
2780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2790    End Select
2800    Resume EXITP

End Sub

Private Sub account_sort_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "account_sort_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2920    intShiftDown = (Shift And acShiftMask) > 0
2930    intAltDown = (Shift And acAltMask) > 0
2940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2970        With Me
2980          .accountno.SetFocus
2990        End With
3000      End Select
3010    End If

        ' ** Shift keys.
3020    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3030      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3040        With Me
3050          intRetVal = 0
3060          If .CurrentRecord > 1 Then
3070            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: below.
3080            .Cx.SetFocus
3090          Else
3100            DoCmd.SelectObject acForm, .Parent.Name, False
3110            .Parent.cmdClose.SetFocus
3120          End If
3130        End With
3140      End Select
3150    End If

EXITP:
3160    KeyCode = intRetVal
3170    Exit Sub

ERRH:
3180    intRetVal = 0
3190    Select Case ERR.Number
        Case Else
3200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3210    End Select
3220    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

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

Private Sub accountno_GotFocus()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

3410    With Me
3420      .accountno.SelLength = 0
3430      .accountno.SelStart = 99
3440    End With

EXITP:
3450    Exit Sub

ERRH:
3460    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
3470    Case Else
3480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3490    End Select
3500    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

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

Private Sub Fix_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Fix_lbl_DblClick"

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

Private Sub fix_BeforeUpdate(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "fix_BeforeUpdate"

3810    With Me
3820      Select Case .fix
          Case True
3830        If IsNull(.Ix) = True And IsNull(.Px) = True And IsNull(.Cx) = True Then
3840          Beep
3850          MsgBox "This account has no discrepancies to fix!", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
3860          Cancel = -1
3870        End If
3880      Case False
            ' ** Nothing.
3890      End Select
3900    End With

EXITP:
3910    Exit Sub

ERRH:
3920    Select Case ERR.Number
        Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Private Sub fix_AfterUpdate()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "fix_AfterUpdate"

        Dim strAccountNo As String

4010    With Me
4020      strAccountNo = .accountno
4030      cmdSave_Click  ' ** Procedure: Above.
4040      DoEvents  ' ** It seems just really hard to get that RecsSum_Fix to recalculate!
4050      .Requery
4060      DoEvents
4070      MoveRec 0, strAccountNo  ' ** Procedure: Below.
4080      DoEvents
4090      .Parent.ResetTots True  ' ** Form Procedure: frmMasterBalance.
4100      DoEvents
4110      .fix.SetFocus
4120    End With

EXITP:
4130    Exit Sub

ERRH:
4140    Select Case ERR.Number
        Case Else
4150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4160    End Select
4170    Resume EXITP

End Sub

Private Sub icash_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "ICash_lbl_DblClick"

4210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub icash_L_lbl_DblClick(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "icash_L_lbl_DblClick"

4310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub Ix_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Ix_lbl_DblClick"

4410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub pcash_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "PCash_lbl_DblClick"

4510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub pcash_L_lbl_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "pcash_L_lbl_DblClick"

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

Private Sub Px_lbl_DblClick(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Px_lbl_DblClick"

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

Private Sub cost_lbl_DblClick(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

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

Private Sub cost_L_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_L_lbl_DblClick"

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

Private Sub Cx_lbl_DblClick(Cancel As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Cx_lbl_DblClick"

5010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5020    Exit Sub

ERRH:
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub Cx_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "Cx_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5170        With Me
5180          intRetVal = 0
5190          lngRecsCur = RecCnt  ' ** Function: Below.
5200          If .CurrentRecord < lngRecsCur Then
5210            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5220            .account_sort.SetFocus
5230          Else
5240            DoCmd.SelectObject acForm, .Parent.Name, False
5250            If .Parent.cmdSelectAll.Enabled = True Then
5260              .Parent.cmdSelectAll.SetFocus
5270            ElseIf .Parent.chkDiscrepancies.Enabled = True And .Parent.chkDiscrepancies.Visible = True Then
5280              .Parent.chkDiscrepancies.SetFocus
5290            ElseIf .Parent.chkIncludeClosed.Enabled = True And .Parent.chkIncludeClosed.Visible = True Then
5300              .Parent.chkIncludeClosed.SetFocus
5310            ElseIf .Parent.chkShowAcctSort.Enabled And .Parent.chkShowAcctSort.Visible = True Then
5320              .Parent.chkShowAcctSort.SetFocus
5330            ElseIf .Parent.chkShowAcctMast.Enabled = True And .Parent.chkShowAcctMast.Visible = True Then
5340              .Parent.chkShowAcctMast.SetFocus
5350            Else
5360              .Parent.chkPageOf.SetFocus
5370            End If
5380          End If
5390        End With
5400      End Select
5410    End If

        ' ** Shift keys.
5420    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5430      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5440        With Me
5450          intRetVal = 0
5460          .cost_L.SetFocus
5470        End With
5480      End Select
5490    End If

EXITP:
5500    KeyCode = intRetVal
5510    Exit Sub

ERRH:
5520    intRetVal = 0
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub

Public Sub SetAcctMastCols(blnWithAcctMast As Boolean)
' ** Without the Master values, Ix, Px, and Cx are meaningless.

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "SetAcctMastCols"

        Dim lngSortOffset As Long
        Dim lngTmp01 As Long

5610    With Me
5620      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5630        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5640      End If
5650      Select Case .Parent.chkShowAcctSort
          Case True
5660        lngSortOffset = 0&
5670      Case False
5680        lngSortOffset = (lngAccountNo_Left - .accountno.Left)
5690      End Select
5700      Select Case blnWithAcctMast
          Case True
            ' ** With account master values.
5710        .shortname.Width = lngShortName_Width  ' ** Reset shortname width when adding or deleting columns.
5720        .shortname_box.Width = .shortname.Width
5730        .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
5740        .shortname_lbl.Width = .shortname.Width
5750        .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
5760        lngTmp01 = ((.shortname_lbl.Left + .shortname_lbl.Width) + (4& * lngTpp))
5770        If .Fix_lbl.Left <> lngTmp01 Then
5780          .Fix_lbl.Left = lngTmp01
5790          .fix.Left = (.Fix_lbl.Left + (lngFix_Left - lngFixLbl_Left))
5800          .Fix_lbl_line.Left = .Fix_lbl.Left
5810        End If
5820        .ICash.Visible = True
5830        .icash_box.Visible = True
5840        .icash_IsClosed_bg.Visible = True
5850        .icash_lbl.Visible = True
5860        .icash_lbl2.Visible = True
5870        .icash_lbl_line.Visible = True
5880        .Ix.Visible = True
5890        .Ix_box.Visible = True
5900        .Ix_IsClosed_bg.Visible = True
5910        .Ix_lbl.Visible = True
5920        .Ix_lbl_line.Visible = True
5930        .PCash.Visible = True
5940        .pcash_box.Visible = True
5950        .pcash_IsClosed_bg.Visible = True
5960        .pcash_lbl.Visible = True
5970        .pcash_lbl2.Visible = True
5980        .pcash_lbl_line.Visible = True
5990        .Px.Visible = True
6000        .Px_box.Visible = True
6010        .Px_IsClosed_bg.Visible = True
6020        .Px_lbl.Visible = True
6030        .Px_lbl_line.Visible = True
6040        .Cost.Visible = True
6050        .cost_box.Visible = True
6060        .cost_IsClosed_bg.Visible = True
6070        .cost_lbl.Visible = True
6080        .cost_lbl2.Visible = True
6090        .cost_lbl_line.Visible = True
6100        .Cx.Visible = True
6110        .Cx_box.Visible = True
6120        .Cx_IsClosed_bg.Visible = True
6130        .Cx_lbl.Visible = True
6140        .Cx_lbl_line.Visible = True
6150        .ICash.Left = (lngICash_Left - lngSortOffset)
6160        .icash_box.Left = .ICash.Left
6170        .icash_IsClosed_bg.Left = (.icash_box.Left - lngTpp)
6180        .icash_lbl.Left = (lngICashLbl_Left - lngSortOffset)
6190        .icash_lbl2.Left = (lngICashLbl_Left - lngSortOffset)
6200        .icash_lbl_line.Left = .icash_lbl.Left
6210        .icash_L.Left = (lngICashL_Left - lngSortOffset)
6220        .icash_L_box.Left = .icash_L.Left
6230        .icash_L_IsClosed_bg.Left = (.icash_L_box.Left - lngTpp)
6240        .icash_L_lbl.Left = (lngICashLLbl_Left - lngSortOffset)
6250        .icash_L_lbl2.Left = (lngICashLLbl_Left - lngSortOffset)
6260        .icash_L_lbl_line.Left = .icash_L_lbl.Left
6270        .Ix.Left = (lngIx_Left - lngSortOffset)
6280        .Ix_box.Left = .Ix.Left
6290        .Ix_IsClosed_bg.Left = (.Ix_box.Left - lngTpp)
6300        .Ix_lbl.Left = (lngIx_Left - lngSortOffset)
6310        .Ix_lbl_line.Left = .Ix_lbl.Left
6320        .PCash.Left = (lngPCash_Left - lngSortOffset)
6330        .pcash_box.Left = .PCash.Left
6340        .pcash_IsClosed_bg.Left = (.pcash_box.Left - lngTpp)
6350        .pcash_lbl.Left = (lngPCashLbl_Left - lngSortOffset)
6360        .pcash_lbl2.Left = (lngPCashLbl_Left - lngSortOffset)
6370        .pcash_lbl_line.Left = .pcash_lbl.Left
6380        .pcash_L.Left = (lngPCashL_Left - lngSortOffset)
6390        .pcash_L_box.Left = .pcash_L.Left
6400        .pcash_L_IsClosed_bg.Left = (.pcash_L_box.Left - lngTpp)
6410        .pcash_L_lbl.Left = (lngPCashLLbl_Left - lngSortOffset)
6420        .pcash_L_lbl2.Left = (lngPCashLLbl_Left - lngSortOffset)
6430        .pcash_L_lbl_line.Left = .pcash_L_lbl.Left
6440        .Px.Left = (lngPx_Left - lngSortOffset)
6450        .Px_box.Left = .Px.Left
6460        .Px_IsClosed_bg.Left = (.Px_box.Left - lngTpp)
6470        .Px_lbl.Left = (lngPx_Left - lngSortOffset)
6480        .Px_lbl_line.Left = .Px_lbl.Left
6490        .Cost.Left = (lngCost_Left - lngSortOffset)
6500        .cost_box.Left = .Cost.Left
6510        .cost_IsClosed_bg.Left = (.cost_box.Left - lngTpp)
6520        .cost_lbl.Left = (lngCostLbl_Left - lngSortOffset)
6530        .cost_lbl2.Left = (lngCostLbl_Left - lngSortOffset)
6540        .cost_lbl_line.Left = .cost_lbl.Left
6550        .cost_L.Left = (lngCostL_Left - lngSortOffset)
6560        .cost_L_box.Left = .cost_L.Left
6570        .cost_L_IsClosed_bg.Left = (.cost_L_box.Left - lngTpp)
6580        .cost_L_lbl.Left = (lngCostLLbl_Left - lngSortOffset)
6590        .cost_L_lbl2.Left = (lngCostLLbl_Left - lngSortOffset)
6600        .cost_L_lbl_line.Left = .cost_L_lbl.Left
6610        .Cx.Left = (lngCx_Left - lngSortOffset)
6620        .Cx_box.Left = .Cx.Left
6630        .Cx_IsClosed_bg.Left = (.Cx_box.Left - lngTpp)
6640        .Cx_lbl.Left = (lngCx_Left - lngSortOffset)
6650        .Cx_lbl_line.Left = .Cx_lbl.Left
6660        lngTmp01 = ((.Cx.Left + .Cx.Width) + (7& * lngTpp))
6670        .IsClosed_bg1.Width = lngTmp01
6680        .IsClosed_bg2.Width = lngTmp01
6690      Case False
            ' ** Without account master values.
6700        .FocusHolder.SetFocus
6710        .shortname.Width = lngShortName_Width  ' ** Reset shortname width when adding or deleting columns.
6720        .shortname_box.Width = .shortname.Width
6730        .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
6740        .shortname_lbl.Width = .shortname.Width
6750        .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
6760        lngTmp01 = ((.shortname_lbl.Left + .shortname_lbl.Width) + (4& * lngTpp))
6770        If .Fix_lbl.Left <> lngTmp01 Then
6780          .Fix_lbl.Left = lngTmp01
6790          .fix.Left = (.Fix_lbl.Left + (lngFix_Left - lngFixLbl_Left))
6800          .Fix_lbl_line.Left = .Fix_lbl.Left
6810        End If
6820        .ICash.Visible = False
6830        .icash_box.Visible = False
6840        .icash_IsClosed_bg.Visible = False
6850        .icash_lbl.Visible = False
6860        .icash_lbl2.Visible = False
6870        .icash_lbl_line.Visible = False
6880        .icash_L.Left = (lngICash_Left - lngSortOffset)
6890        .icash_L_box.Left = .icash_L.Left
6900        .icash_L_IsClosed_bg.Left = (.icash_L_box.Left - lngTpp)
6910        .icash_L_lbl.Left = ((lngICashLLbl_Left - (lngICashL_Left - lngICash_Left)) - lngSortOffset)
6920        .icash_L_lbl2.Left = ((lngICashLLbl_Left - (lngICashL_Left - lngICash_Left)) - lngSortOffset)
6930        .icash_L_lbl_line.Left = .icash_L_lbl.Left
6940        .Ix.Visible = False
6950        .Ix_box.Visible = False
6960        .Ix_IsClosed_bg.Visible = False
6970        .Ix_lbl.Visible = False
6980        .Ix_lbl_line.Visible = False
6990        .PCash.Visible = False
7000        .pcash_box.Visible = False
7010        .pcash_IsClosed_bg.Visible = False
7020        .pcash_lbl.Visible = False
7030        .pcash_lbl2.Visible = False
7040        .pcash_lbl_line.Visible = False
7050        .pcash_L.Left = (lngICashL_Left - lngSortOffset)
7060        .pcash_L_box.Left = .pcash_L.Left
7070        .pcash_L_IsClosed_bg.Left = (.pcash_L_box.Left - lngTpp)
7080        .pcash_L_lbl.Left = ((lngPCashLLbl_Left - (lngPCashL_Left - lngICashL_Left)) - lngSortOffset)
7090        .pcash_L_lbl2.Left = ((lngPCashLLbl_Left - (lngPCashL_Left - lngICashL_Left)) - lngSortOffset)
7100        .pcash_L_lbl_line.Left = .pcash_L_lbl.Left
7110        .Px.Visible = False
7120        .Px_box.Visible = False
7130        .Px_IsClosed_bg.Visible = False
7140        .Px_lbl.Visible = False
7150        .Px_lbl_line.Visible = False
7160        .Cost.Visible = False
7170        .cost_box.Visible = False
7180        .cost_IsClosed_bg.Visible = False
7190        .cost_lbl.Visible = False
7200        .cost_lbl2.Visible = False
7210        .cost_lbl_line.Visible = False
7220        .cost_L.Left = (lngIx_Left - lngSortOffset)
7230        .cost_L_box.Left = .cost_L.Left
7240        .cost_L_IsClosed_bg.Left = (.cost_L_box.Left - lngTpp)
7250        .cost_L_lbl.Left = ((lngCostLLbl_Left - (lngCostL_Left - lngIx_Left)) - lngSortOffset)
7260        .cost_L_lbl2.Left = ((lngCostLLbl_Left - (lngCostL_Left - lngIx_Left)) - lngSortOffset)
7270        .cost_L_lbl_line.Left = .cost_L_lbl.Left
7280        .Cx.Visible = False
7290        .Cx_box.Visible = False
7300        .Cx_IsClosed_bg.Visible = False
7310        .Cx_lbl.Visible = False
7320        .Cx_lbl_line.Visible = False
7330        .ICash.Left = 0&
7340        .icash_box.Left = 0&
7350        .icash_IsClosed_bg.Left = 0&
7360        .icash_lbl.Left = 0&
7370        .icash_lbl2.Left = 0&
7380        .icash_lbl_line.Left = 0&
7390        .PCash.Left = 0&
7400        .pcash_box.Left = 0&
7410        .pcash_IsClosed_bg.Left = 0&
7420        .pcash_lbl.Left = 0&
7430        .pcash_lbl2.Left = 0&
7440        .pcash_lbl_line.Left = 0&
7450        .Cost.Left = 0&
7460        .cost_box.Left = 0&
7470        .cost_IsClosed_bg.Left = 0&
7480        .cost_lbl.Left = 0&
7490        .cost_lbl2.Left = 0&
7500        .cost_lbl_line.Left = 0&
7510        .Ix.Left = 0&
7520        .Ix_box.Left = 0&
7530        .Ix_IsClosed_bg.Left = 0&
7540        .Ix_lbl.Left = 0&
7550        .Ix_lbl_line.Left = 0&
7560        .Px.Left = 0&
7570        .Px_box.Left = 0&
7580        .Px_IsClosed_bg.Left = 0&
7590        .Px_lbl.Left = 0&
7600        .Px_lbl_line.Left = 0&
7610        .Cx.Left = 0&
7620        .Cx_box.Left = 0&
7630        .Cx_IsClosed_bg.Left = 0&
7640        .Cx_lbl.Left = 0&
7650        .Cx_lbl_line.Left = 0&
7660        lngTmp01 = ((.cost_L.Left + .cost_L.Width) + (7& * lngTpp))
7670        .IsClosed_bg1.Width = lngTmp01
7680        .IsClosed_bg2.Width = lngTmp01
7690      End Select
7700    End With

EXITP:
7710    Exit Sub

ERRH:
7720    Select Case ERR.Number
        Case Else
7730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7740    End Select
7750    Resume EXITP

End Sub

Public Sub SetAcctSortCols(blnWithAcctSort As Boolean, blnDiscrepancies As Boolean, strCallingForm As String)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "SetAcctSortCols"

        Dim strAccountNo As String, strFilter As String
        Dim lngTmp01 As Long

7810    With Me
7820      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7830        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7840      End If
7850      Select Case blnWithAcctSort
          Case True
            ' ** With account_sort.
7860        .Filter = vbNullString
7870        .FilterOn = False
7880        DoEvents
7890        .account_sort.Visible = True
7900        .account_sort_box.Visible = True
7910        .account_sort_IsClosed_bg.Visible = True
7920        .account_sort_lbl.Visible = True
7930        .account_sort_lbl2.Visible = True
7940        .account_sort_lbl_line.Visible = True
7950        Select Case .Parent.chkShowAcctMast
            Case True
7960          .Ix.Visible = True
7970          .Ix_box.Visible = True
7980          .Ix_IsClosed_bg.Visible = True
7990          .Ix_lbl.Visible = True
8000          .Ix_lbl_line.Visible = True
8010          .Px.Visible = True
8020          .Px_box.Visible = True
8030          .Px_IsClosed_bg.Visible = True
8040          .Px_lbl.Visible = True
8050          .Px_lbl_line.Visible = True
8060          .Cx.Visible = True
8070          .Cx_box.Visible = True
8080          .Cx_IsClosed_bg.Visible = True
8090          .Cx_lbl.Visible = True
8100          .Cx_lbl_line.Visible = True
8110          .accountno.Left = lngAccountNo_Left
8120          .accountno_box.Left = .accountno.Left
8130          .accountno_IsClosed_bg.Left = (.accountno_box.Left - lngTpp)
8140          .accountno_lbl.Left = lngAccountNoLbl_Left
8150          .accountno_lbl_line.Left = .accountno_lbl.Left
8160          .shortname.Width = lngShortName_Width  ' ** Reset shortname width when adding or deleting columns.
8170          .shortname_box.Width = .shortname.Width
8180          .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
8190          .shortname_lbl.Width = .shortname.Width
8200          .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
8210          .shortname.Left = lngShortName_Left
8220          .shortname_box.Left = .shortname.Left
8230          .shortname_IsClosed_bg.Left = (.shortname_box.Left - lngTpp)
8240          .shortname_lbl.Left = lngShortNameLbl_Left
8250          .shortname_lbl_line.Left = .shortname_lbl.Left
8260          .fix.Left = lngFix_Left
8270          .Fix_lbl.Left = lngFixLbl_Left
8280          .Fix_lbl_line.Left = .Fix_lbl.Left
8290          .ICash.Left = lngICash_Left
8300          .icash_box.Left = .ICash.Left
8310          .icash_IsClosed_bg.Left = (.icash_box.Left - lngTpp)
8320          .icash_lbl.Left = lngICashLbl_Left
8330          .icash_lbl2.Left = lngICashLbl_Left
8340          .icash_lbl_line.Left = .icash_lbl.Left
8350          .icash_L.Left = lngICashL_Left
8360          .icash_L_box.Left = .icash_L.Left
8370          .icash_L_IsClosed_bg.Left = (.icash_L_box.Left - lngTpp)
8380          .icash_L_lbl.Left = lngICashLLbl_Left
8390          .icash_L_lbl2.Left = lngICashLLbl_Left
8400          .icash_L_lbl_line.Left = .icash_L_lbl.Left
8410          .Ix.Left = lngIx_Left
8420          .Ix_box.Left = .Ix.Left
8430          .Ix_IsClosed_bg.Left = .Ix_box.Left
8440          .Ix_lbl.Left = lngIx_Left
8450          .Ix_lbl_line.Left = .Ix_lbl.Left
8460          .PCash.Left = lngPCash_Left
8470          .pcash_box.Left = .PCash.Left
8480          .pcash_IsClosed_bg.Left = (.pcash_box.Left - lngTpp)
8490          .pcash_lbl.Left = lngPCashLbl_Left
8500          .pcash_lbl2.Left = lngPCashLbl_Left
8510          .pcash_lbl_line.Left = .pcash_lbl.Left
8520          .pcash_L.Left = lngPCashL_Left
8530          .pcash_L_box.Left = .pcash_L.Left
8540          .pcash_L_IsClosed_bg.Left = (.pcash_L_box.Left - lngTpp)
8550          .pcash_L_lbl.Left = lngPCashLLbl_Left
8560          .pcash_L_lbl2.Left = lngPCashLLbl_Left
8570          .pcash_L_lbl_line.Left = .pcash_L_lbl.Left
8580          .Px.Left = lngPx_Left
8590          .Px_box.Left = .Px.Left
8600          .Px_IsClosed_bg.Left = .Px_box.Left
8610          .Px_lbl.Left = lngPx_Left
8620          .Px_lbl_line.Left = .Px_lbl.Left
8630          .Cost.Left = lngCost_Left
8640          .cost_box.Left = .Cost.Left
8650          .cost_IsClosed_bg.Left = (.cost_box.Left - lngTpp)
8660          .cost_lbl.Left = lngCostLbl_Left
8670          .cost_lbl2.Left = lngCostLbl_Left
8680          .cost_lbl_line.Left = .cost_lbl.Left
8690          .cost_L.Left = lngCostL_Left
8700          .cost_L_box.Left = .cost_L.Left
8710          .cost_L_IsClosed_bg.Left = (.cost_L_box.Left - lngTpp)
8720          .cost_L_lbl.Left = lngCostLLbl_Left
8730          .cost_L_lbl2.Left = lngCostLLbl_Left
8740          .cost_L_lbl_line.Left = .cost_L_lbl.Left
8750          .Cx.Left = lngCx_Left
8760          .Cx_box.Left = .Cx.Left
8770          .Cx_IsClosed_bg.Left = .Cx_box.Left
8780          .Cx_lbl.Left = lngCx_Left
8790          .Cx_lbl_line.Left = .Cx_lbl.Left
8800          .IsClosed_bg1.Width = lngIsClosedBG_Width
8810          .IsClosed_bg2.Width = lngIsClosedBG_Width
8820        Case False
8830          .FocusHolder.SetFocus
8840          .Ix.Visible = False
8850          .Ix_box.Visible = False
8860          .Ix_IsClosed_bg.Visible = False
8870          .Ix_lbl.Visible = False
8880          .Ix_lbl_line.Visible = False
8890          .Px.Visible = False
8900          .Px_box.Visible = False
8910          .Px_IsClosed_bg.Visible = False
8920          .Px_lbl.Visible = False
8930          .Px_lbl_line.Visible = False
8940          .Cx.Visible = False
8950          .Cx_box.Visible = False
8960          .Cx_IsClosed_bg.Visible = False
8970          .Cx_lbl.Visible = False
8980          .Cx_lbl_line.Visible = False
8990          .accountno.Left = lngAccountNo_Left
9000          .accountno_box.Left = .accountno.Left
9010          .accountno_IsClosed_bg.Left = (.accountno_box.Left - lngTpp)
9020          .accountno_lbl.Left = lngAccountNoLbl_Left
9030          .accountno_lbl_line.Left = .accountno_lbl.Left
9040          .shortname.Left = lngShortName_Left
9050          .shortname_box.Left = .shortname.Left
9060          .shortname_IsClosed_bg.Left = (.shortname_box.Left - lngTpp)
9070          .shortname_lbl.Left = lngShortNameLbl_Left
9080          .shortname_lbl_line.Left = .shortname_lbl.Left
9090          .shortname.Width = lngShortName_Width  ' ** Reset shortname width when adding or deleting columns.
9100          .shortname_box.Width = .shortname.Width
9110          .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
9120          .shortname_lbl.Width = .shortname.Width
9130          .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
9140          .fix.Left = lngFix_Left
9150          .Fix_lbl.Left = lngFixLbl_Left
9160          .Fix_lbl_line.Left = .Fix_lbl.Left
9170          .icash_L.Left = lngICash_Left
9180          .icash_L_box.Left = .icash_L.Left
9190          .icash_L_IsClosed_bg.Left = (.icash_L_box.Left - lngTpp)
9200          .icash_L_lbl.Left = lngICashLbl_Left
9210          .icash_L_lbl2.Left = lngICashLbl_Left
9220          .icash_L_lbl_line.Left = .icash_L_lbl.Left
9230          .pcash_L.Left = lngICashL_Left
9240          .pcash_L_box.Left = .pcash_L.Left
9250          .pcash_L_IsClosed_bg.Left = (.pcash_L_box.Left - lngTpp)
9260          .pcash_L_lbl.Left = lngICashLLbl_Left
9270          .pcash_L_lbl2.Left = lngICashLLbl_Left
9280          .pcash_L_lbl_line.Left = .pcash_L_lbl.Left
9290          .cost_L.Left = lngIx_Left
9300          .cost_L_box.Left = .cost_L.Left
9310          .cost_L_IsClosed_bg.Left = (.cost_L_box.Left - lngTpp)
9320          .cost_L_lbl.Left = (lngCostLLbl_Left - (lngCostL_Left - lngIx_Left))
9330          .cost_L_lbl2.Left = (lngCostLLbl_Left - (lngCostL_Left - lngIx_Left))
9340          .cost_L_lbl_line.Left = .cost_L_lbl.Left
9350          lngTmp01 = ((.cost_L.Left + .cost_L.Width) + (7& * lngTpp))
9360          .IsClosed_bg1.Width = lngTmp01
9370          .IsClosed_bg2.Width = lngTmp01
9380        End Select
9390        DoEvents
9400        SortNow "accountno_lbl_DblClick"  ' ** Procedure: Below.
9410        SortNow "account_sort_lbl_DblClick"  ' ** Procedure: Below.
9420      Case False
            ' ** Without account_sort.
9430        strFilter = vbNullString
9440        Select Case strCallingForm
            Case "frmAccountProfile"
9450          strAccountNo = Forms(strCallingForm).frmAccountProfile_Sub.Form.accountno
9460          strFilter = "[accountno] = '" & strAccountNo & "'"
9470          If .Parent.chkDiscrepancies = True Then
9480            strFilter = strFilter & " And ([Ix] = 'X' Or [Px] = 'X' Or [Cx] = 'X')"
9490          End If
9500          If .Parent.chkIncludeClosed = False Then
9510            strFilter = strFilter & " And [IsClosed] = False"
9520          End If
9530          .Filter = strFilter
9540          .FilterOn = True
9550        Case "frmMenu_Maintenance"
9560          If .Parent.chkDiscrepancies = True Then
9570            strFilter = "([Ix] = 'X' Or [Px] = 'X' Or [Cx] = 'X')"
9580          End If
9590          If .Parent.chkIncludeClosed = False Then
9600            If strFilter <> vbNullString Then strFilter = strFilter & " And "
9610            strFilter = strFilter & "[IsClosed] = False"
9620          End If
9630          .Filter = strFilter
9640          If strFilter = vbNullString Then
9650            .FilterOn = False
9660          Else
9670            .FilterOn = True
9680          End If
9690        End Select
9700        DoEvents
9710  On Error Resume Next
9720        .accountno.SetFocus
9730  On Error GoTo ERRH
9740        .FocusHolder.SetFocus
9750        .account_sort.Visible = False
9760        .account_sort_box.Visible = False
9770        .account_sort_IsClosed_bg.Visible = False
9780        .account_sort_lbl.Visible = False
9790        .account_sort_lbl2.Visible = False
9800        .account_sort_lbl_line.Visible = False
9810        Select Case .Parent.chkShowAcctMast
            Case True
9820          .Ix.Visible = True
9830          .Ix_box.Visible = True
9840          .Ix_IsClosed_bg.Visible = True
9850          .Ix_lbl.Visible = True
9860          .Ix_lbl_line.Visible = True
9870          .Px.Visible = True
9880          .Px_box.Visible = True
9890          .Px_IsClosed_bg.Visible = True
9900          .Px_lbl.Visible = True
9910          .Px_lbl_line.Visible = True
9920          .Cx.Visible = True
9930          .Cx_box.Visible = True
9940          .Cx_IsClosed_bg.Visible = True
9950          .Cx_lbl.Visible = True
9960          .Cx_lbl_line.Visible = True
9970          .accountno.Left = (lngAccountNo_Left - lngOffset)
9980          .accountno_box.Left = .accountno.Left
9990          .accountno_IsClosed_bg.Left = (.accountno_box.Left - lngTpp)
10000         .accountno_lbl.Left = (lngAccountNoLbl_Left - lngOffset)
10010         .accountno_lbl_line.Left = .accountno_lbl.Left
10020         .shortname.Left = (lngShortName_Left - lngOffset)
10030         .shortname_box.Left = .shortname.Left
10040         .shortname_IsClosed_bg.Left = (.shortname_box.Left - lngTpp)
10050         .shortname_lbl.Left = (lngShortNameLbl_Left - lngOffset)
10060         .shortname_lbl_line.Left = .shortname_lbl.Left
10070         .shortname.Width = lngShortName_Width  ' ** Reset shortname width when adding or deleting columns.
10080         .shortname_box.Width = .shortname.Width
10090         .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
10100         .shortname_lbl.Width = .shortname.Width
10110         .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
10120         .fix.Left = (lngFix_Left - lngOffset)
10130         .Fix_lbl.Left = (lngFixLbl_Left - lngOffset)
10140         .Fix_lbl_line.Left = .Fix_lbl.Left
10150         .ICash.Left = (lngICash_Left - lngOffset)
10160         .icash_box.Left = .ICash.Left
10170         .icash_IsClosed_bg.Left = (.icash_box.Left - lngTpp)
10180         .icash_lbl.Left = (lngICashLbl_Left - lngOffset)
10190         .icash_lbl2.Left = (lngICashLbl_Left - lngOffset)
10200         .icash_lbl_line.Left = .icash_lbl.Left
10210         .icash_L.Left = (lngICashL_Left - lngOffset)
10220         .icash_L_box.Left = .icash_L.Left
10230         .icash_L_IsClosed_bg.Left = (.icash_L_box.Left - lngTpp)
10240         .icash_L_lbl.Left = (lngICashLLbl_Left - lngOffset)
10250         .icash_L_lbl2.Left = (lngICashLLbl_Left - lngOffset)
10260         .icash_L_lbl_line.Left = .icash_L_lbl.Left
10270         .Ix.Left = (lngIx_Left - lngOffset)
10280         .Ix_box.Left = .Ix.Left
10290         .Ix_IsClosed_bg.Left = (.Ix_box.Left - lngTpp)
10300         .Ix_lbl.Left = (lngIx_Left - lngOffset)
10310         .Ix_lbl_line.Left = .Ix_lbl.Left
10320         .PCash.Left = (lngPCash_Left - lngOffset)
10330         .pcash_box.Left = .PCash.Left
10340         .pcash_IsClosed_bg.Left = (.pcash_box.Left - lngTpp)
10350         .pcash_lbl.Left = (lngPCashLbl_Left - lngOffset)
10360         .pcash_lbl2.Left = (lngPCashLbl_Left - lngOffset)
10370         .pcash_lbl_line.Left = .pcash_lbl.Left
10380         .pcash_L.Left = (lngPCashL_Left - lngOffset)
10390         .pcash_L_box.Left = .pcash_L.Left
10400         .pcash_L_IsClosed_bg.Left = (.pcash_L_box.Left - lngTpp)
10410         .pcash_L_lbl.Left = (lngPCashLLbl_Left - lngOffset)
10420         .pcash_L_lbl2.Left = (lngPCashLLbl_Left - lngOffset)
10430         .pcash_L_lbl_line.Left = .pcash_L_lbl.Left
10440         .Px.Left = (lngPx_Left - lngOffset)
10450         .Px_box.Left = .Px.Left
10460         .Px_IsClosed_bg.Left = (.Px_box.Left - lngTpp)
10470         .Px_lbl.Left = (lngPx_Left - lngOffset)
10480         .Px_lbl_line.Left = .Px_lbl.Left
10490         .Cost.Left = (lngCost_Left - lngOffset)
10500         .cost_box.Left = .Cost.Left
10510         .cost_IsClosed_bg.Left = (.cost_box.Left - lngTpp)
10520         .cost_lbl.Left = (lngCostLbl_Left - lngOffset)
10530         .cost_lbl2.Left = (lngCostLbl_Left - lngOffset)
10540         .cost_lbl_line.Left = .cost_lbl.Left
10550         .cost_L.Left = (lngCostL_Left - lngOffset)
10560         .cost_L_box.Left = .cost_L.Left
10570         .cost_L_IsClosed_bg.Left = (.cost_L_box.Left - lngTpp)
10580         .cost_L_lbl.Left = (lngCostLLbl_Left - lngOffset)
10590         .cost_L_lbl2.Left = (lngCostLLbl_Left - lngOffset)
10600         .cost_L_lbl_line.Left = .cost_L_lbl.Left
10610         .Cx.Left = (lngCx_Left - lngOffset)
10620         .Cx_box.Left = .Cx.Left
10630         .Cx_IsClosed_bg.Left = (.Cx_box.Left - lngTpp)
10640         .Cx_lbl.Left = (lngCx_Left - lngOffset)
10650         .Cx_lbl_line.Left = .Cx_lbl.Left
10660         lngTmp01 = ((.Cx.Left + .Cx.Width) + (7& * lngTpp))
10670         .IsClosed_bg1.Width = lngTmp01
10680         .IsClosed_bg2.Width = lngTmp01
10690       Case False
10700         .FocusHolder.SetFocus
10710         .Ix.Visible = False
10720         .Ix_box.Visible = False
10730         .Ix_IsClosed_bg.Visible = False
10740         .Ix_lbl.Visible = False
10750         .Ix_lbl_line.Visible = False
10760         .Px.Visible = False
10770         .Px_box.Visible = False
10780         .Px_IsClosed_bg.Visible = False
10790         .Px_lbl.Visible = False
10800         .Px_lbl_line.Visible = False
10810         .Cx.Visible = False
10820         .Cx_box.Visible = False
10830         .Cx_IsClosed_bg.Visible = False
10840         .Cx_lbl.Visible = False
10850         .Cx_lbl_line.Visible = False
10860         .accountno.Left = (lngAccountNo_Left - lngOffset)
10870         .accountno_box.Left = .accountno.Left
10880         .accountno_IsClosed_bg.Left = (.accountno_box.Left - lngTpp)
10890         .accountno_lbl.Left = (lngAccountNoLbl_Left - lngOffset)
10900         .accountno_lbl_line.Left = .accountno_lbl.Left
10910         .shortname.Left = (lngShortName_Left - lngOffset)
10920         .shortname_box.Left = .shortname.Left
10930         .shortname_IsClosed_bg.Left = (.shortname_box.Left - lngTpp)
10940         .shortname_lbl.Left = (lngShortNameLbl_Left - lngOffset)
10950         .shortname_lbl_line.Left = .shortname_lbl.Left
10960         .shortname.Width = lngShortName_Width  ' ** Reset shortname width when adding or deleting columns.
10970         .shortname_box.Width = .shortname.Width
10980         .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
10990         .shortname_lbl.Width = .shortname.Width
11000         .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
11010         .fix.Left = (lngFix_Left - lngOffset)
11020         .Fix_lbl.Left = (lngFixLbl_Left - lngOffset)
11030         .Fix_lbl_line.Left = .Fix_lbl.Left
11040         .icash_L.Left = (lngICash_Left - lngOffset)
11050         .icash_L_box.Left = .icash_L.Left
11060         .icash_L_IsClosed_bg.Left = (.icash_L_box.Left - lngTpp)
11070         .icash_L_lbl.Left = (lngICashLbl_Left - lngOffset)
11080         .icash_L_lbl2.Left = (lngICashLbl_Left - lngOffset)
11090         .icash_L_lbl_line.Left = .icash_L_lbl.Left
11100         .pcash_L.Left = (lngICashL_Left - lngOffset)
11110         .pcash_L_box.Left = .pcash_L.Left
11120         .pcash_L_IsClosed_bg.Left = (.pcash_L_box.Left - lngTpp)
11130         .pcash_L_lbl.Left = (lngICashLLbl_Left - lngOffset)
11140         .pcash_L_lbl2.Left = (lngICashLLbl_Left - lngOffset)
11150         .pcash_L_lbl_line.Left = .pcash_L_lbl.Left
11160         .cost_L.Left = (lngIx_Left - lngOffset)
11170         .cost_L_box.Left = .cost_L.Left
11180         .cost_L_IsClosed_bg.Left = (.cost_L_box.Left - lngTpp)
11190         .cost_L_lbl.Left = ((lngCostLLbl_Left - (lngCostL_Left - lngIx_Left)) - lngOffset)
11200         .cost_L_lbl2.Left = ((lngCostLLbl_Left - (lngCostL_Left - lngIx_Left)) - lngOffset)
11210         .cost_L_lbl_line.Left = .cost_L_lbl.Left
11220         lngTmp01 = ((.cost_L.Left + .cost_L.Width) + (7& * lngTpp))
11230         .IsClosed_bg1.Width = lngTmp01
11240         .IsClosed_bg2.Width = lngTmp01
11250       End Select
11260       DoEvents
11270       SortNow "shortname_lbl_DblClick"  ' ** Procedure: Below.
11280       SortNow "accountno_lbl_DblClick"  ' ** Procedure: Below.
11290     End Select
11300     If strCallingForm = "frmMenu_Maintenance" Then
            ' ** I don't know why this filter keeps reappearing!
11310       strFilter = vbNullString
11320       If blnDiscrepancies = True Then
11330         strFilter = "([Ix] = 'X' Or [Px] = 'X' Or [Cx] = 'X')"
11340       End If
11350       If .Parent.chkIncludeClosed = False Then
11360         If strFilter <> vbNullString Then strFilter = strFilter & " And "
11370         strFilter = strFilter & "[IsClosed] = False"
11380       End If
11390       .Filter = strFilter
11400       If strFilter = vbNullString Then
11410         .FilterOn = False
11420       Else
11430         .FilterOn = True
11440       End If
11450     End If
11460   End With

EXITP:
11470   Exit Sub

ERRH:
11480   Select Case ERR.Number
        Case Else
11490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11500   End Select
11510   Resume EXITP

End Sub

Public Sub SetColWidths(lngShort_Width As Long)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "SetColWidths"

        Dim blnSortHere As Boolean
        Dim lngTmp01 As Long

11610   With Me
11620     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
11630       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
11640     End If
11650     blnSortHere = False
11660     If lngShort_Width <> .shortname.Width Then
11670       If lngShort_Width < .shortname.Width Then
11680         If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
11690         lngTmp01 = (.shortname.Width - lngShort_Width)  ' ** Amount to narrow.
11700         .shortname.Width = lngShort_Width
11710         .shortname_box.Width = .shortname.Width
11720         .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
11730         .shortname_lbl.Width = .shortname.Width
11740         .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
11750         If blnSortHere = True Then
11760           .Sort_line.Width = (.shortname_lbl.Width + lngTpp)
11770           .Sort_lbl.Left = ((.shortname_lbl.Left + .shortname_lbl.Width) - lngSortLbl_Width)
11780         End If
11790         blnSortHere = False
11800         If .Sort_line.Left = .Fix_lbl.Left Then blnSortHere = True
11810         .Fix_lbl.Left = (.Fix_lbl.Left - lngTmp01)
11820         .fix.Left = (.Fix_lbl.Left + (lngFix_Left - lngFixLbl_Left))
11830         .Fix_lbl_line.Left = .Fix_lbl.Left
11840         If blnSortHere = True Then
11850           .Sort_line.Left = .Fix_lbl.Left
11860           .Sort_lbl.Left = (((.Fix_lbl.Left + .Fix_lbl.Width) - lngSortLbl_Width) + (8& * lngTpp))
11870         End If
11880         blnSortHere = False
11890         If .ICash.Visible = True Then
11900           If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
11910           .ICash.Left = (.ICash.Left - lngTmp01)
11920           .icash_box.Left = .ICash.Left
11930           .icash_IsClosed_bg.Left = (.ICash.Left - lngTpp)
11940           .icash_lbl.Left = .ICash.Left
11950           .icash_lbl2.Left = .icash_lbl.Left
11960           .icash_lbl_line.Left = .icash_lbl.Left
11970           If blnSortHere = True Then
11980             .Sort_line.Left = .icash_lbl.Left
11990             .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
12000           End If
12010           blnSortHere = False
12020           If .Sort_line.Left = .Ix_lbl.Left Then blnSortHere = True
12030           .Ix.Left = (.Ix.Left - lngTmp01)
12040           .Ix_box.Left = .Ix.Left
12050           .Ix_IsClosed_bg.Left = (.Ix.Left - lngTpp)
12060           .Ix_lbl.Left = .Ix.Left
12070           .Ix_lbl_line.Left = .Ix_lbl.Left
12080           If blnSortHere = True Then
12090             .Sort_line.Left = .Ix_lbl.Left
12100             .Sort_lbl.Left = (((.Ix_lbl.Left + .Ix_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
12110           End If
12120         End If
12130         blnSortHere = False
12140         If .Sort_line.Left = .icash_L_lbl.Left Then blnSortHere = True
12150         .icash_L.Left = (.icash_L.Left - lngTmp01)
12160         .icash_L_box.Left = .icash_L.Left
12170         .icash_L_IsClosed_bg.Left = (.icash_L.Left - lngTpp)
12180         .icash_L_lbl.Left = .icash_L.Left
12190         .icash_L_lbl2.Left = .icash_L_lbl.Left
12200         .icash_L_lbl_line.Left = .icash_L_lbl.Left
12210         If blnSortHere = True Then
12220           .Sort_line.Left = .icash_L_lbl.Left
12230           .Sort_lbl.Left = ((.icash_L_lbl.Left + .icash_L_lbl.Width) - lngSortLbl_Width)
12240         End If
12250         blnSortHere = False
12260         If .PCash.Visible = True Then
12270           If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
12280           .PCash.Left = (.PCash.Left - lngTmp01)
12290           .pcash_box.Left = .PCash.Left
12300           .pcash_IsClosed_bg.Left = (.PCash.Left - lngTpp)
12310           .pcash_lbl.Left = .PCash.Left
12320           .pcash_lbl2.Left = .pcash_lbl.Left
12330           .pcash_lbl_line.Left = .pcash_lbl.Left
12340           If blnSortHere = True Then
12350             .Sort_line.Left = .pcash_lbl.Left
12360             .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
12370           End If
12380           blnSortHere = False
12390           If .Sort_line.Left = .Px_lbl.Left Then blnSortHere = True
12400           .Px.Left = (.Px.Left - lngTmp01)
12410           .Px_box.Left = .Px.Left
12420           .Px_IsClosed_bg.Left = (.Px.Left - lngTpp)
12430           .Px_lbl.Left = .Px.Left
12440           .Px_lbl_line.Left = .Px_lbl.Left
12450           If blnSortHere = True Then
12460             .Sort_line.Left = .Px_lbl.Left
12470             .Sort_lbl.Left = (((.Px_lbl.Left + .Px_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
12480           End If
12490         End If
12500         blnSortHere = False
12510         If .Sort_line.Left = .pcash_L_lbl.Left Then blnSortHere = True
12520         .pcash_L.Left = (.pcash_L.Left - lngTmp01)
12530         .pcash_L_box.Left = .pcash_L.Left
12540         .pcash_L_IsClosed_bg.Left = (.pcash_L.Left - lngTpp)
12550         .pcash_L_lbl.Left = .pcash_L.Left
12560         .pcash_L_lbl2.Left = .pcash_L_lbl.Left
12570         .pcash_L_lbl_line.Left = .pcash_L_lbl.Left
12580         If blnSortHere = True Then
12590           .Sort_line.Left = .pcash_L_lbl.Left
12600           .Sort_lbl.Left = ((.pcash_L_lbl.Left + .pcash_L_lbl.Width) - lngSortLbl_Width)
12610         End If
12620         blnSortHere = False
12630         If .Cost.Visible = True Then
12640           If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
12650           .Cost.Left = (.Cost.Left - lngTmp01)
12660           .cost_box.Left = .Cost.Left
12670           .cost_IsClosed_bg.Left = (.Cost.Left - lngTpp)
12680           .cost_lbl.Left = .Cost.Left
12690           .cost_lbl2.Left = .cost_lbl.Left
12700           .cost_lbl_line.Left = .cost_lbl.Left
12710           If blnSortHere = True Then
12720             .Sort_line.Left = .cost_lbl.Left
12730             .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
12740           End If
12750           blnSortHere = False
12760           If .Sort_line.Left = .Cx_lbl.Left Then blnSortHere = True
12770           .Cx.Left = (.Cx.Left - lngTmp01)
12780           .Cx_box.Left = .Cx.Left
12790           .Cx_IsClosed_bg.Left = (.Cx.Left - lngTpp)
12800           .Cx_lbl.Left = .Cx.Left
12810           .Cx_lbl_line.Left = .Cx_lbl.Left
12820           If blnSortHere = True Then
12830             .Sort_line.Left = .Cx_lbl.Left
12840             .Sort_lbl.Left = (((.Cx_lbl.Left + .Cx_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
12850           End If
12860         End If
12870         blnSortHere = False
12880         If .Sort_line.Left = .cost_L_lbl.Left Then blnSortHere = True
12890         .cost_L.Left = (.cost_L.Left - lngTmp01)
12900         .cost_L_box.Left = .cost_L.Left
12910         .cost_L_IsClosed_bg.Left = (.cost_L.Left - lngTpp)
12920         .cost_L_lbl.Left = .cost_L.Left
12930         .cost_L_lbl2.Left = .cost_L_lbl.Left
12940         .cost_L_lbl_line.Left = .cost_L_lbl.Left
12950         If blnSortHere = True Then
12960           .Sort_line.Left = .cost_L_lbl.Left
12970           .Sort_lbl.Left = ((.cost_L_lbl.Left + .cost_L_lbl.Width) - lngSortLbl_Width)
12980         End If
12990         blnSortHere = False
13000         .IsClosed_bg1.Width = (.IsClosed_bg1.Width - lngTmp01)
13010         .IsClosed_bg2.Width = .IsClosed_bg1.Width
13020         .Width = .Width - lngTmp01
13030       ElseIf lngShort_Width > .shortname.Width Then  ' ** Reverse the order later, if you wish.
              ' ** It may still be narrower than the standard width.
13040         If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
13050         lngTmp01 = (lngShort_Width - .shortname.Width)
13060         .Width = .Width + lngTmp01
13070         .shortname.Width = lngShort_Width
13080         .shortname_box.Width = .shortname.Width
13090         .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
13100         .shortname_lbl.Width = .shortname.Width
13110         .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
13120         If blnSortHere = True Then
13130           .Sort_line.Width = (.shortname_lbl.Width + lngTpp)
13140           .Sort_lbl.Left = ((.shortname_lbl.Left + .shortname_lbl.Width) - lngSortLbl_Width)
13150         End If
13160         blnSortHere = False

13170         If .Sort_line.Left = .Fix_lbl.Left Then blnSortHere = True
13180         .Fix_lbl.Left = (.Fix_lbl.Left + lngTmp01)
13190         .fix.Left = (.Fix_lbl.Left + (lngFix_Left - lngFixLbl_Left))
13200         .Fix_lbl_line.Left = .Fix_lbl.Left
13210         If blnSortHere = True Then
13220           .Sort_line.Left = .Fix_lbl.Left
13230           .Sort_lbl.Left = (((.Fix_lbl.Left + .Fix_lbl.Width) - lngSortLbl_Width) + (8& * lngTpp))
13240         End If
13250         blnSortHere = False
13260         If .ICash.Visible = True Then
13270           If .Sort_line.Left = .icash_lbl.Left Then blnSortHere = True
13280           .ICash.Left = (.ICash.Left + lngTmp01)
13290           .icash_box.Left = .ICash.Left
13300           .icash_IsClosed_bg.Left = (.ICash.Left - lngTpp)
13310           .icash_lbl.Left = .ICash.Left
13320           .icash_lbl2.Left = .icash_lbl.Left
13330           .icash_lbl_line.Left = .icash_lbl.Left
13340           If blnSortHere = True Then
13350             .Sort_line.Left = .icash_lbl.Left
13360             .Sort_lbl.Left = ((.icash_lbl.Left + .icash_lbl.Width) - lngSortLbl_Width)
13370           End If
13380           blnSortHere = False
13390           If .Sort_line.Left = .Ix_lbl.Left Then blnSortHere = True
13400           .Ix.Left = (.Ix.Left + lngTmp01)
13410           .Ix_box.Left = .Ix.Left
13420           .Ix_IsClosed_bg.Left = (.Ix.Left - lngTpp)
13430           .Ix_lbl.Left = .Ix.Left
13440           .Ix_lbl_line.Left = .Ix_lbl.Left
13450           If blnSortHere = True Then
13460             .Sort_line.Left = .Ix_lbl.Left
13470             .Sort_lbl.Left = (((.Ix_lbl.Left + .Ix_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
13480           End If
13490         End If
13500         blnSortHere = False
13510         If .Sort_line.Left = .icash_L_lbl.Left Then blnSortHere = True
13520         .icash_L.Left = (.icash_L.Left + lngTmp01)
13530         .icash_L_box.Left = .icash_L.Left
13540         .icash_L_IsClosed_bg.Left = (.icash_L.Left - lngTpp)
13550         .icash_L_lbl.Left = .icash_L.Left
13560         .icash_L_lbl2.Left = .icash_L_lbl.Left
13570         .icash_L_lbl_line.Left = .icash_L_lbl.Left
13580         If blnSortHere = True Then
13590           .Sort_line.Left = .icash_L_lbl.Left
13600           .Sort_lbl.Left = ((.icash_L_lbl.Left + .icash_L_lbl.Width) - lngSortLbl_Width)
13610         End If
13620         blnSortHere = False
13630         If .PCash.Visible = True Then
13640           If .Sort_line.Left = .pcash_lbl.Left Then blnSortHere = True
13650           .PCash.Left = (.PCash.Left + lngTmp01)
13660           .pcash_box.Left = .PCash.Left
13670           .pcash_IsClosed_bg.Left = (.PCash.Left - lngTpp)
13680           .pcash_lbl.Left = .PCash.Left
13690           .pcash_lbl2.Left = .pcash_lbl.Left
13700           .pcash_lbl_line.Left = .pcash_lbl.Left
13710           If blnSortHere = True Then
13720             .Sort_line.Left = .pcash_lbl.Left
13730             .Sort_lbl.Left = ((.pcash_lbl.Left + .pcash_lbl.Width) - lngSortLbl_Width)
13740           End If
13750           blnSortHere = False
13760           If .Sort_line.Left = .Px_lbl.Left Then blnSortHere = True
13770           .Px.Left = (.Px.Left + lngTmp01)
13780           .Px_box.Left = .Px.Left
13790           .Px_IsClosed_bg.Left = (.Px.Left - lngTpp)
13800           .Px_lbl.Left = .Px.Left
13810           .Px_lbl_line.Left = .Px_lbl.Left
13820           If blnSortHere = True Then
13830             .Sort_line.Left = .Px_lbl.Left
13840             .Sort_lbl.Left = (((.Px_lbl.Left + .Px_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
13850           End If
13860         End If
13870         blnSortHere = False
13880         If .Sort_line.Left = .pcash_L_lbl.Left Then blnSortHere = True
13890         .pcash_L.Left = (.pcash_L.Left + lngTmp01)
13900         .pcash_L_box.Left = .pcash_L.Left
13910         .pcash_L_IsClosed_bg.Left = (.pcash_L.Left - lngTpp)
13920         .pcash_L_lbl.Left = .pcash_L.Left
13930         .pcash_L_lbl2.Left = .pcash_L_lbl.Left
13940         .pcash_L_lbl_line.Left = .pcash_L_lbl.Left
13950         If blnSortHere = True Then
13960           .Sort_line.Left = .pcash_L_lbl.Left
13970           .Sort_lbl.Left = ((.pcash_L_lbl.Left + .pcash_L_lbl.Width) - lngSortLbl_Width)
13980         End If
13990         blnSortHere = False
14000         If .Cost.Visible = True Then
14010           If .Sort_line.Left = .cost_lbl.Left Then blnSortHere = True
14020           .Cost.Left = (.Cost.Left + lngTmp01)
14030           .cost_box.Left = .Cost.Left
14040           .cost_IsClosed_bg.Left = (.Cost.Left - lngTpp)
14050           .cost_lbl.Left = .Cost.Left
14060           .cost_lbl2.Left = .cost_lbl.Left
14070           .cost_lbl_line.Left = .cost_lbl.Left
14080           If blnSortHere = True Then
14090             .Sort_line.Left = .cost_lbl.Left
14100             .Sort_lbl.Left = ((.cost_lbl.Left + .cost_lbl.Width) - lngSortLbl_Width)
14110           End If
14120           blnSortHere = False
14130           If .Sort_line.Left = .Cx_lbl.Left Then blnSortHere = True
14140           .Cx.Left = (.Cx.Left + lngTmp01)
14150           .Cx_box.Left = .Cx.Left
14160           .Cx_IsClosed_bg.Left = (.Cx.Left - lngTpp)
14170           .Cx_lbl.Left = .Cx.Left
14180           .Cx_lbl_line.Left = .Cx_lbl.Left
14190           If blnSortHere = True Then
14200             .Sort_line.Left = .Cx_lbl.Left
14210             .Sort_lbl.Left = (((.Cx_lbl.Left + .Cx_lbl.Width) - lngSortLbl_Width) + (6& * lngTpp))
14220           End If
14230         End If
14240         blnSortHere = False
14250         If .Sort_line.Left = .cost_L_lbl.Left Then blnSortHere = True
14260         .cost_L.Left = (.cost_L.Left + lngTmp01)
14270         .cost_L_box.Left = .cost_L.Left
14280         .cost_L_IsClosed_bg.Left = (.cost_L.Left - lngTpp)
14290         .cost_L_lbl.Left = .cost_L.Left
14300         .cost_L_lbl2.Left = .cost_L_lbl.Left
14310         .cost_L_lbl_line.Left = .cost_L_lbl.Left
14320         If blnSortHere = True Then
14330           .Sort_line.Left = .cost_L_lbl.Left
14340           .Sort_lbl.Left = ((.cost_L_lbl.Left + .cost_L_lbl.Width) - lngSortLbl_Width)
14350         End If
14360         blnSortHere = False
14370         .IsClosed_bg1.Width = (.IsClosed_bg1.Width + lngTmp01)
14380         .IsClosed_bg2.Width = .IsClosed_bg1.Width
14390         .Width = .IsClosed_bg1.Width
14400       End If
14410     End If
14420   End With

EXITP:
14430   Exit Sub

ERRH:
14440   Select Case ERR.Number
        Case Else
14450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14460   End Select
14470   Resume EXITP

End Sub

Public Function RecCnt() As Long

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

14510   lngRetVal = 0&

14520   With Me
14530     Set rst = .RecordsetClone
14540     With rst
14550       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
14560       Else
14570         .MoveLast
14580         lngRetVal = .RecordCount
14590       End If
14600       .Close
14610     End With
14620   End With

EXITP:
14630   Set rst = Nothing
14640   RecCnt = lngRetVal
14650   Exit Function

ERRH:
14660   lngRetVal = 0&
14670   Select Case ERR.Number
        Case Else
14680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14690   End Select
14700   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

14810   Select Case intWhere
        Case acCmdRecordsGoToFirst
14820     DoCmd.RunCommand acCmdRecordsGoToFirst
14830   Case acCmdRecordsGoToLast
14840     DoCmd.RunCommand acCmdRecordsGoToLast
14850   Case acCmdRecordsGoToPrevious
14860     DoCmd.RunCommand acCmdRecordsGoToPrevious
14870   Case acCmdRecordsGoToNext
14880     DoCmd.RunCommand acCmdRecordsGoToNext
14890   Case Else
14900     If IsMissing(varID) = False Then
14910       With Me
14920         Set rst = .RecordsetClone
14930         With rst
14940           .FindFirst "[accountno] = '" & varID & "'"
14950           If .NoMatch = False Then
14960             Me.Bookmark = .Bookmark
14970           End If
14980           .Close
14990         End With
15000       End With
15010     End If
15020   End Select

EXITP:
15030   Set rst = Nothing
15040   Exit Sub

ERRH:
15050   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
15060   Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
15070   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
15080   Case Else
15090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15100   End Select
15110   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim ctl As Access.Control
        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [account_sort]"

15210   With Me
15220     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
15230       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
15240     End If
15250     .Controls(strSortLbl).Visible = False
15260     .Controls(strSortLbl).Width = lngSortLbl_Width
15270     .Controls(strSortLine).Visible = False
15280     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
15290     If strProc = "Form_Load" Then
15300       strCalled = "account_sort"
15310       strSortNow = strSortOrig
15320       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
15330       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
15340       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
15350       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
15360       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
15370       .Controls(strSortLbl).Left = lngSortLbl_Left
15380       .Controls(strSortLbl).Top = lngSortLbl_Top
15390       .Controls(strSortLbl).Caption = strArwUp
15400       .Controls(strSortLbl).ForeColor = CLR_DKBLU
15410       .Controls(strSortLine).Left = lngSortLine_Left
15420       .Controls(strSortLine).Top = lngSortLine_Top
15430       .Controls(strSortLine).Width = lngSortLine_Width
15440     Else
15450       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
15460       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
15470       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
15480       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
15490       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
15500       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
15510       If strCalled = "Cx" Then
15520         .Controls(strSortLbl).Width = (lngSortLbl_Width - (2& * lngTpp))
15530         .Controls(strSortLbl).Left = (lngSortLbl_Left + (7& * lngTpp))
15540       Else
15550         .Controls(strSortLbl).Left = lngSortLbl_Left
15560       End If
15570       .Controls(strSortLbl).Top = lngSortLbl_Top
15580       .Controls(strSortLine).Left = lngSortLine_Left
15590       .Controls(strSortLine).Top = lngSortLine_Top
15600       .Controls(strSortLine).Width = lngSortLine_Width
15610       If strCalled = "account_sort" Then
15620         strSortAsc = strSortOrig
15630         If strSortNow = strSortAsc Then
15640           strSortNow = strSortOrig & " DESC"
15650           .Controls(strSortLbl).Caption = strArwDn
15660           .Controls(strSortLbl).ForeColor = CLR_DKRED
15670         Else
15680           strSortNow = strSortAsc
15690           .Controls(strSortLbl).Caption = strArwUp
15700           .Controls(strSortLbl).ForeColor = CLR_DKBLU
15710         End If
15720       Else
15730         .Controls(strSortLbl).Caption = strArwUp
15740         .Controls(strSortLbl).ForeColor = CLR_DKBLU
15750         Select Case strCalled
              Case "accountno"
15760           strSortAsc = "[account_sort]"
15770         Case "shortname"
15780           strSortAsc = "[shortname]" & strStdAsc
15790         Case "fix"
15800           strSortAsc = "[fix]" & strStdAsc
15810           lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
15820           .Controls(strSortLbl).Left = lngSortLbl_Left
15830         Case "icash"
15840           strSortAsc = "[icash], [pcash], [cost]" & strStdAsc
15850         Case "icash_L"
15860           strSortAsc = "[icash_L], [pcash_L], [cost_L]" & strStdAsc
15870         Case "Ix"
15880           strSortAsc = "[Ix]" & strStdAsc
15890           lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
15900           .Controls(strSortLbl).Left = lngSortLbl_Left
15910         Case "pcash"
15920           strSortAsc = "[pcash], [icash], [cost]" & strStdAsc
15930           lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
15940           .Controls(strSortLbl).Left = lngSortLbl_Left
15950         Case "pcash_L"
15960           strSortAsc = "[pcash_L], [icash_L], [cost_L]" & strStdAsc
15970           lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
15980           .Controls(strSortLbl).Left = lngSortLbl_Left
15990         Case "Px"
16000           strSortAsc = "[Px]" & strStdAsc
16010           lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
16020           .Controls(strSortLbl).Left = lngSortLbl_Left
16030         Case "cost"
16040           strSortAsc = "[cost], [icash], [pcash]" & strStdAsc
16050         Case "cost_L"
16060           strSortAsc = "[cost_L], [icash_L], [pcash_L]" & strStdAsc
16070         Case "Cx"
16080           strSortAsc = "[Cx]" & strStdAsc
16090         End Select
16100         If strSortNow = strSortAsc Then
16110           intCnt = (CharCnt(strSortAsc, ",") + 1)  ' ** Module Function: modStringFuncs.
16120           Select Case intCnt
                Case 1&
16130             strTmp01 = strSortAsc & " DESC"
16140           Case 2&
16150             intPos01 = InStr(strSortNow, ",")
16160             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
16170             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
16180           Case 3&
16190             intPos01 = InStr(strSortAsc, ",")
16200             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
16210             strTmp02 = Mid(strSortAsc, intPos01)
16220             intPos01 = InStr(2, strTmp02, ",")
16230             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16240             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
16250           Case 4&
16260             intPos01 = InStr(strSortAsc, ",")
16270             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
16280             strTmp02 = Mid(strSortAsc, intPos01)
16290             intPos01 = InStr(2, strTmp02, ",")
16300             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16310             strTmp02 = Mid(strTmp02, intPos01)
16320             intPos01 = InStr(2, strTmp02, ",")
16330             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16340             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
16350           Case 5&
16360             intPos01 = InStr(strSortAsc, ",")
16370             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
16380             strTmp02 = Mid(strSortAsc, intPos01)
16390             intPos01 = InStr(2, strTmp02, ",")
16400             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16410             strTmp02 = Mid(strTmp02, intPos01)
16420             intPos01 = InStr(2, strTmp02, ",")
16430             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16440             strTmp02 = Mid(strTmp02, intPos01)
16450             intPos01 = InStr(2, strTmp02, ",")
16460             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16470             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
16480           End Select
16490           strSortNow = strTmp01
16500           .Controls(strSortLbl).Caption = strArwDn
16510           .Controls(strSortLbl).ForeColor = CLR_DKRED
16520         Else
16530           strSortNow = strSortAsc
16540         End If
16550       End If
16560     End If
16570     .Controls(strSortLbl).Visible = True
16580     .Controls(strSortLine).Visible = True
16590     .OrderBy = strSortNow
16600     .OrderByOn = True
16610   End With

        ' ** Potential sorts:
        ' **   [account_sort]
        ' **   [accountno]
        ' **   [shortname], [account_sort]
        ' **   [fix], [account_sort]
        ' **   [icash], [pcash], [cost], [account_sort]
        ' **   [icash_L], [pcash_L], [cost_L], [account_sort]
        ' **   [Ix], [account_sort]
        ' **   [pcash], [icash], [cost], [account_sort]
        ' **   [pcash_L], [icash_L], [cost_L], [account_sort]
        ' **   [Px], [account_sort]
        ' **   [cost], [icash], [pcash], [account_sort]
        ' **   [cost_L], [icash_L], [pcash_L], [account_sort]
        ' **   [Cx], [account_sort]

EXITP:
16620   Set ctl = Nothing
16630   Exit Sub

ERRH:
16640   Select Case ERR.Number
        Case Else
16650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16660   End Select
16670   Resume EXITP

End Sub

Public Function SortNow_Get() As String

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

16710   strRetVal = strSortNow

EXITP:
16720   SortNow_Get = strRetVal
16730   Exit Function

ERRH:
16740   strRetVal = vbNullString
16750   Select Case ERR.Number
        Case Else
16760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16770   End Select
16780   Resume EXITP

End Function
