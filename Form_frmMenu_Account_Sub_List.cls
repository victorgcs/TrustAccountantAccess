VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Account_Sub_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Account_Sub_List"

'VGC 10/11/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add Account:      A {cmdAdd on frmMenu_Account}
' **   Close:            C {cmdClose on frmMenu_Account}
' **   Delete Account:   D {cmdDelete on frmMenu_Account}
' **   Search:           E {cmdSearch on frmMenu_Account}
' **   Hide:             H {cmdHideTransactions on frmMenu_Account}
' **   Inc/Exp Codes:    I {cmdEditIECodes on frmMenu_Account}
' **   Display List:     L {opgSubForm_optList on frmMenu_Account}
' **   Comments:         M {cmdEditComments on frmMenu_Account}
' **   Short Name:       N {shortname}
' **   Display One:      O {opgSubForm_optOne on frmMenu_Account}
' **   Profile:          P {cmdAccountProfile on frmMenu_Account}
' **   Transactions:     R {cmdTransactions on frmMenu_Account}
' **   Assets:           S {cmdAssets on frmMenu_Account}
' **   Tax Codes:        T {cmdEditTaxCodes on frmMenu_Account}
' **   Include Closed:   U {chkIncludeClosed on frmMenu_Account}
' **   Exit:             X {cmdClose on frmMenu_Account}

' ** Shortcut Ctrl keys responsive from this form:
' **   Next Record:      N {MoveRec on frmMenu_Account_Sub_Nav}
' **   Previous Record:  P {MoveRec on frmMenu_Account_Sub_Nav}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Unlock Var:       C {blnFormOpening: Form_KeyDown on frmMenu_Account}
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm on frmMenu_Account}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private lngSortLbl_Width As Long

Private lngForm_Width As Long
Private lngDateClosed_Left As Long, lngShortName_Width As Long

Private blnNotOpen As Boolean
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
150       If strCallingForm = "frmMenu_Account" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

180         .form_width_line.Left = 0&  ' ** Just get this out of the way.

190         lngSortLbl_Width = .Sort_lbl.Width
200         lngForm_Width = .Width
210         lngDateClosed_Left = .dateclosed.Left
220         lngShortName_Width = .shortname.Width

230       Else
240         Cancel = -1
250       End If
260     End With

270     If Cancel = -1 Then
280       Beep
290       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
300     End If

EXITP:
310     Exit Sub

ERRH:
320     Select Case ERR.Number
        Case Else
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
340     End Select
350     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .accountno.ForeColor = CLR_DISABLED_FG
430       .accountno.BackColor = CLR_DISABLED_BG
440       .shortname.ForeColor = CLR_DISABLED_FG
450       .shortname.BackColor = CLR_DISABLED_BG
460       .dateclosed.ForeColor = CLR_DISABLED_FG
470       .dateclosed.BackColor = CLR_DISABLED_BG
480       blnNotOpen = True
490     End With

EXITP:
500     Exit Sub

ERRH:
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Sub

Private Sub Form_Current()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

610     UpdateHiddenCnt  ' ** Function: Below.

620     With Me
630       gstrAccountNo = .accountno
640       .Parent.accountno = .accountno
650       gstrAccountName = .shortname
660       .Parent.shortname = .shortname
670       .Parent.legalname = .legalname
680       If IsNull(.dateclosed) = True Then
690         .Parent.dateclosed = Null
700       Else
710         .Parent.dateclosed = .dateclosed
720       End If
730       .Parent.hidden_trans = .hidden_trans
740       If blnNotOpen = True Then
750         If .Parent.frmMenu_Account_Sub_Nav.Form.accountno <> .accountno Then
760           .Parent.frmMenu_Account_Sub_Nav.Form.MoveRec 0, .accountno  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
770         End If
780       End If
790       Select Case .IsClosed
          Case True
800         .accountno.BackColor = CLR_LTTEAL3
810         .shortname.BackColor = CLR_LTTEAL3
820         .dateclosed.BackColor = CLR_LTTEAL3
830         .dateclosed.Enabled = True
840       Case False
850         .accountno.BackColor = CLR_LTTEAL
860         .shortname.BackColor = CLR_LTTEAL
870         .dateclosed.BackColor = MY_CLR_BGE
880         .dateclosed.Enabled = False
890       End Select
900     End With

EXITP:
910     Exit Sub

ERRH:
920     Select Case ERR.Number
        Case 2424  ' ** The expression you entered has a field, control, or property name that Microsoft Access can't find.
          ' ** Not sure why it said this!
930     Case 2455  ' ** You entered an expression that has an invalid reference to the property Form/Report.
          ' ** Not loaded yet.
940     Case Else
950       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
960     End Select
970     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim Cancel As Integer
        Dim intRetVal As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Short Name:       N {shortname}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Next Record:      N {MoveRec on frmMenu_Account_Sub_Nav}
        ' **   Previous Record:  P {MoveRec on frmMenu_Account_Sub_Nav}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Unlock Var:       C {blnFormOpening: Form_KeyDown on frmMenu_Account}
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm on frmMenu_Account}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyReturn
1070        With Me
1080          intRetVal = 0
1090          If IsNull(.accountno) = False Then
1100            Form_DblClick Cancel  ' ** Procedure: Below.
1110          End If
1120        End With
1130      Case vbKeyUp
1140        With Me
1150          intRetVal = 0
1160          .Parent.frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
1170        End With
1180      Case vbKeyDown
1190        With Me
1200          intRetVal = 0
1210          .Parent.frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
1220        End With
1230      Case vbKeyEscape
1240        With Me
1250          intRetVal = 0
1260          .Parent.cmdClose_Click  ' ** Form Procedure: frmMenu_Account.
1270        End With
1280      End Select
1290    End If

        ' ** Alt keys.
1300    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1310      Select Case intRetVal
          Case vbKeyC
1320        With Me
1330          intRetVal = 0
1340          .Parent.cmdClose_Click  ' ** Form Procedure: frmMenu_Account.
1350        End With
1360      Case vbKeyN
1370        With Me
1380          intRetVal = 0
1390          .shortname.SetFocus
1400        End With
1410      Case vbKeyX
1420        With Me
1430          intRetVal = 0
1440          .Parent.cmdClose_Click  ' ** Form Procedure: frmMenu_Account.
1450        End With
1460      End Select
1470    End If

        ' ** Ctrl keys.
1480    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1490      Select Case intRetVal
          Case vbKeyN
1500        With Me
1510          intRetVal = 0
1520          .Parent.frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
1530        End With
1540      Case vbKeyP
1550        With Me
1560          intRetVal = 0
1570          .Parent.frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
1580        End With
1590      Case vbKeyTab, vbKeyReturn
1600        With Me
1610          intRetVal = 0
1620          DoCmd.SelectObject acForm, .Parent.Name, False
1630          .Parent.cmdSearch.SetFocus
1640        End With
1650      End Select
1660    End If

        ' ** Ctrl-Shift keys.
1670    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1680      Select Case intRetVal
          Case vbKeyC
1690        With Me
              ' ** In case this gets stuck on.
1700          .Parent.FormCloseVar False  ' ** blnFormOpening = False
1710          DoBeeps 2, 300  ' ** Module Function: modWindowFunctions.
1720        End With
1730      Case vbKeyF
1740        With Me
1750          intRetVal = 0
1760          DoCmd.SelectObject acForm, .Parent.Name, False
1770          .Parent.FocusHolder.SetFocus
1780        End With
1790      Case vbKeyT
1800        With Me
1810          intRetVal = 0  ' ** User doesn't matter.
1820          If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
1830            FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
1840          Else
1850            If CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Module Function: modFileUtilities.
1860              With .Parent
1870                Select Case .chkLoadTimer_vis
                    Case True
1880                  .chkLoadTimer.Visible = False
1890                  .chkLoadTimer_lbl1.Visible = False
1900                  .cmdDelete.Caption = "&Delete Account"
1910                  .cmdDelete.StatusBarText = "Delete the current account - Alt+D"
1920                  .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Alt+D"
1930                Case False
1940                  .chkLoadTimer.Visible = True
1950                  .chkLoadTimer_lbl1.Visible = True
1960                  .cmdDelete.Caption = "Delete Account"
1970                  .cmdDelete.StatusBarText = "Delete the current account - Ctrl+D"
1980                  .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Ctrl+D"
1990                End Select
2000              End With
2010            End If
2020          End If
2030          .Parent.chkLoadTimer_vis = .Parent.chkLoadTimer.Visible
2040        End With
2050      Case vbKeyX
2060        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
2070          With Me
2080            intRetVal = 0
2090            .Parent.JustClose_Set True  ' ** Form Procedure: frmMenu_Account.
2100            SetOption_Dev  ' ** Module Function: modUtilities.
2110            DoCmd.Close acForm, strCallingForm
2120          End With
2130        End If
2140      Case vbKeyTab, vbKeyReturn
2150        With Me
2160          intRetVal = 0
2170          DoCmd.SelectObject acForm, .Parent.Name, False
2180          .Parent.cmdAccountNew.SetFocus
2190        End With
2200      End Select
2210    End If

EXITP:
2220    KeyCode = intRetVal
2230    Exit Sub

ERRH:
2240    intRetVal = 0
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Sub

Private Sub Form_DblClick(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_DblClick"

2310    With Me
2320      DoCmd.SelectObject acForm, .Parent.Name, False
2330      .Parent.cmdAccountProfile.SetFocus
2340      .Parent.cmdAccountProfile_Click  ' ** Form Procedure: frmMenu_Account.
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

Private Sub Detail_DblClick(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_DblClick"

2510    Form_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub accountno_GotFocus()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

2610    With Me
2620      .accountno.SelLength = 0
2630      .accountno.SelStart = 20
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

Private Sub accountno_DblClick(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_DblClick"

2710    Form_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
2870        With Me
2880          intRetVal = 0
2890          .shortname.SetFocus
2900        End With
2910      End Select
2920    End If

        ' ** Shift keys.
2930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2950        With Me
2960          intRetVal = 0
2970          If .CurrentRecord > 1 Then
2980            MoveRec acCmdRecordsGoToPrevious
2990            If .dateclosed.Visible = True And .dateclosed.Enabled = True Then
3000              .dateclosed.SetFocus
3010            Else
3020              .shortname.SetFocus
3030            End If
3040          Else
3050            DoCmd.SelectObject acForm, .Parent.Name, False
3060            .Parent.cmdAccountNew.SetFocus
3070          End If
3080        End With
3090      End Select
3100    End If

EXITP:
3110    KeyCode = intRetVal
3120    Exit Sub

ERRH:
3130    intRetVal = 0
3140    Select Case ERR.Number
        Case Else
3150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3160    End Select
3170    Resume EXITP

End Sub

Private Sub shortname_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_DblClick"

3210    Form_DblClick Cancel  ' ** Procedure: Above.

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub shortname_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3320    intShiftDown = (Shift And acShiftMask) > 0
3330    intAltDown = (Shift And acAltMask) > 0
3340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3370        With Me
3380          intRetVal = 0
3390          If .dateclosed.Visible = True And .dateclosed.Enabled = True Then
3400            .dateclosed.SetFocus
3410          Else
3420            lngRecsCur = RecCnt  ' ** Function: Below.
3430            If .CurrentRecord < lngRecsCur Then
3440              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3450              .accountno.SetFocus
3460            Else
3470              DoCmd.SelectObject acForm, .Parent.Name, False
3480              .Parent.cmdSearch.SetFocus
3490            End If
3500          End If
3510        End With
3520      End Select
3530    End If

        ' ** Shift keys.
3540    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3550      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3560        With Me
3570          intRetVal = 0
3580          .accountno.SetFocus
3590        End With
3600      End Select
3610    End If

EXITP:
3620    KeyCode = intRetVal
3630    Exit Sub

ERRH:
3640    intRetVal = 0
3650    Select Case ERR.Number
        Case Else
3660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3670    End Select
3680    Resume EXITP

End Sub

Private Sub dateclosed_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "dateclosed_DblClick"

3710    Form_DblClick Cancel  ' ** Procedure: Above.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub dateclosed_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "dateclosed_KeyDown"

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
3890          lngRecsCur = RecCnt  ' ** Function: Below.
3900          If .CurrentRecord < lngRecsCur Then
3910            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3920            .accountno.SetFocus
3930          Else
3940            DoCmd.SelectObject acForm, .Parent.Name, False
3950            .Parent.cmdSearch.SetFocus
3960          End If
3970        End With
3980      End Select
3990    End If

        ' ** Shift keys.
4000    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4010      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4020        With Me
4030          intRetVal = 0
4040          .shortname.SetFocus
4050        End With
4060      End Select
4070    End If

EXITP:
4080    KeyCode = intRetVal
4090    Exit Sub

ERRH:
4100    intRetVal = 0
4110    Select Case ERR.Number
        Case Else
4120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4130    End Select
4140    Resume EXITP

End Sub

Private Function UpdateHiddenCnt() As Boolean

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateHiddenCnt"

        Dim blnRetVal As Boolean

4210    blnRetVal = True

4220    With Me
4230      If gblnAdmin = True Then
4240        If .hidden_trans > 0& Then
4250          .Parent.cmdHideTransactions_lbl.Caption = CStr(.hidden_trans)
4260        Else
4270          .Parent.cmdHideTransactions_lbl.Caption = vbNullString
4280        End If
4290      End If
4300    End With

EXITP:
4310    UpdateHiddenCnt = blnRetVal
4320    Exit Function

ERRH:
4330    blnRetVal = False
4340    Select Case ERR.Number
        Case 2424  ' ** The expression you entered has a field, control, or property name that Microsoft Access can't find.
          ' ** Not sure why it said this!
4350    Case Else
4360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4370    End Select
4380    Resume EXITP

End Function

Public Sub SubResize_List(lngWidth_Diff As Long)
' ** It won't ever get narrower.

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "SubResize_List"

        Dim blnSortHere As Boolean

4410    With Me

4420      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowsFunctions.
4430        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
4440      End If

4450      blnSortHere = False
4460      If lngWidth_Diff > 0& Then
4470        .Width = (lngForm_Width + lngWidth_Diff)
4480        If .Sort_line.Left = .dateclosed_lbl.Left Then blnSortHere = True
4490        .dateclosed.Left = (lngDateClosed_Left + lngWidth_Diff)
4500        .dateclosed_IsClosed_bg.Left = (.dateclosed.Left - lngTpp)
4510        .dateclosed_lbl.Left = .dateclosed.Left
4520        .dateclosed_lbl_line.Left = .dateclosed_lbl.Left
4530        If blnSortHere = True Then
4540          .Sort_line.Left = .dateclosed_lbl.Left
4550          .Sort_lbl.Left = ((.dateclosed_lbl.Left + .dateclosed_lbl.Width) - lngSortLbl_Width)
4560        End If
4570        blnSortHere = False
4580        If .Sort_line.Left = .shortname_lbl.Left Then blnSortHere = True
4590        .shortname.Width = (lngShortName_Width + lngWidth_Diff)
4600        .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
4610        .shortname_IsntClosed_bg.Width = .shortname_IsClosed_bg.Width
4620        .shortname_lbl.Width = .shortname.Width
4630        .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
4640        If blnSortHere = True Then
4650          .Sort_line.Width = (.shortname_lbl.Width + lngTpp)
4660          .Sort_lbl.Left = ((.shortname_lbl.Left + .shortname_lbl.Width) - lngSortLbl_Width)
4670        End If
4680        blnSortHere = False
4690        .IsClosed_bg1.Width = (lngForm_Width + lngWidth_Diff)
4700        .IsClosed_bg2.Width = .IsClosed_bg1.Width
4710        .Width = (lngForm_Width + lngWidth_Diff)
4720      Else
4730        If .Width <> lngForm_Width Or .shortname.Width <> lngShortName_Width Then
4740          .Width = lngForm_Width
4750          If .Sort_line.Left = .dateclosed_lbl.Left Then blnSortHere = True
4760          .dateclosed.Left = lngDateClosed_Left
4770          .dateclosed_IsClosed_bg.Left = (.dateclosed.Left - lngTpp)
4780          .dateclosed_lbl.Left = .dateclosed.Left
4790          .dateclosed_lbl_line.Left = .dateclosed_lbl.Left
4800          If blnSortHere = True Then
4810            .Sort_line.Left = .dateclosed_lbl.Left
4820            .Sort_lbl.Left = ((.dateclosed_lbl.Left + .dateclosed_lbl.Width) - lngSortLbl_Width)
4830          End If
4840          blnSortHere = False
4850          .shortname.Width = lngShortName_Width
4860          .shortname_IsClosed_bg.Width = (.shortname.Width + (2& * lngTpp))
4870          .shortname_IsntClosed_bg.Width = .shortname_IsClosed_bg.Width
4880          .shortname_lbl.Width = .shortname.Width
4890          .shortname_lbl_line.Width = (.shortname_lbl.Width + lngTpp)
4900          If blnSortHere = True Then
4910            .Sort_line.Width = (.shortname_lbl.Width + lngTpp)
4920            .Sort_lbl.Left = ((.shortname_lbl.Left + .shortname_lbl.Width) - lngSortLbl_Width)
4930          End If
4940          blnSortHere = False
4950          .IsClosed_bg1.Width = lngForm_Width
4960          .IsClosed_bg2.Width = .IsClosed_bg1.Width
4970          .Width = lngForm_Width
4980        End If
4990      End If

5000    End With

EXITP:
5010    Exit Sub

ERRH:
5020    Select Case ERR.Number
        Case Else
5030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5040    End Select
5050    Resume EXITP

End Sub

Public Function RecCnt() As Long

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

5110    lngRetVal = 0&

5120    With Me
5130      Set rst = .RecordsetClone
5140      With rst
5150        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
5160        Else
5170          .MoveLast
5180          lngRetVal = .RecordCount
5190        End If
5200        .Close
5210      End With
5220    End With

EXITP:
5230    Set rst = Nothing
5240    RecCnt = lngRetVal
5250    Exit Function

ERRH:
5260    lngRetVal = 0&
5270    Select Case ERR.Number
        Case Else
5280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5290    End Select
5300    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

5410    Select Case intWhere
        Case acCmdRecordsGoToFirst
5420      DoCmd.RunCommand acCmdRecordsGoToFirst
5430    Case acCmdRecordsGoToLast
5440      DoCmd.RunCommand acCmdRecordsGoToLast
5450    Case acCmdRecordsGoToPrevious
5460      DoCmd.RunCommand acCmdRecordsGoToPrevious
5470    Case acCmdRecordsGoToNext
5480      DoCmd.RunCommand acCmdRecordsGoToNext
5490    Case Else
5500      If IsMissing(varID) = False Then
5510        If IsEmpty(varID) = False Then
5520          If IsNull(varID) = False Then
5530            If Trim(varID) <> vbNullString Then
5540              With Me
5550                Set rst = .RecordsetClone
5560                With rst
5570                  If .BOF = True And .EOF = True Then
                        ' ** Skip it!
5580                  Else
5590                    .MoveFirst
5600                    .FindFirst "[accountno] = '" & CStr(varID) & "'"
5610                    If .NoMatch = False Then
5620                      Me.Bookmark = .Bookmark
5630                    End If
5640                  End If
5650                  .Close
5660                End With
5670              End With
5680            End If
5690          End If
5700        End If
5710      End If
5720    End Select

EXITP:
5730    Set rst = Nothing
5740    Exit Sub

ERRH:
5750    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
5760    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5770    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5780    Case Else
5790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5800    End Select
5810    Resume EXITP

End Sub
