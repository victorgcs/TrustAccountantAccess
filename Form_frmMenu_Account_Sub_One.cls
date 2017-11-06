VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Account_Sub_One"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Account_Sub_One"

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

Private lngForm_Width As Long
Private lngDateClosed_Left As Long, lngDateClosedLbl_left As Long, lngShortName_Width As Long, lngLegalname_Width As Long

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

190         lngForm_Width = .Width
200         lngDateClosed_Left = .dateclosed.Left
210         lngDateClosedLbl_left = .dateclosed_lbl.Left
220         lngShortName_Width = .shortname.Width
230         lngLegalname_Width = .legalname.Width

240       Else
250         Cancel = -1
260       End If
270     End With

280     If Cancel = -1 Then
290       Beep
300       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
310     End If

EXITP:
320     Exit Sub

ERRH:
330     Select Case ERR.Number
        Case Else
340       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
350     End Select
360     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .accountno.ForeColor = CLR_DISABLED_FG
430       .accountno.BackColor = CLR_DISABLED_BG
440       .dateclosed.ForeColor = CLR_DISABLED_FG
450       .dateclosed.BackColor = CLR_DISABLED_BG
460       .shortname.ForeColor = CLR_DISABLED_FG
470       .shortname.BackColor = CLR_DISABLED_BG
480       .legalname.ForeColor = CLR_DISABLED_FG
490       .legalname.BackColor = CLR_DISABLED_BG
500     End With

EXITP:
510     Exit Sub

ERRH:
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

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
740     End With

EXITP:
750     Exit Sub

ERRH:
760     Select Case ERR.Number
        Case Else
770       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
780     End Select
790     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim Cancel As Integer
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

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
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyReturn
870         With Me
880           intRetVal = 0
890           If IsNull(.accountno) = False Then
900             Form_DblClick Cancel  ' ** Procedure: Below.
910           End If
920         End With
930       Case vbKeyEscape
940         With Me
950           intRetVal = 0
960           .Parent.cmdClose_Click  ' ** Form Procedure: frmMenu_Account.
970         End With
980       End Select
990     End If

        ' ** Alt keys.
1000    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1010      Select Case intRetVal
          Case vbKeyC
1020        With Me
1030          intRetVal = 0
1040          .Parent.cmdClose_Click  ' ** Form Procedure: frmMenu_Account.
1050        End With
1060      Case vbKeyN
1070        With Me
1080          intRetVal = 0
1090          .shortname.SetFocus
1100        End With
1110      Case vbKeyX
1120        With Me
1130          intRetVal = 0
1140          .Parent.cmdClose_Click  ' ** Form Procedure: frmMenu_Account.
1150        End With
1160      End Select
1170    End If

        ' ** Ctrl keys.
1180    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1190      Select Case intRetVal
          Case vbKeyN
1200        With Me
1210          intRetVal = 0
1220          .Parent.frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
1230        End With
1240      Case vbKeyP
1250        With Me
1260          intRetVal = 0
1270          .Parent.frmMenu_Account_Sub_Nav.Form.MoveRec acCmdRecordsGoToPrevious  ' ** Form Procedure: frmMenu_Account_Sub_Nav.
1280        End With
1290      End Select
1300    End If

        ' ** Ctrl-Shift keys.
1310    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1320      Select Case intRetVal
          Case vbKeyC
1330        With Me
              ' ** In case this gets stuck on.
1340          .Parent.FormCloseVar False  ' ** blnFormOpening = False
1350          DoBeeps 2, 300  ' ** Module Function: modWindowFunctions.
1360        End With
1370      Case vbKeyF
1380        With Me
1390          intRetVal = 0
1400          DoCmd.SelectObject acForm, .Parent.Name, False
1410          .Parent.FocusHolder.SetFocus
1420        End With
1430      Case vbKeyT
1440        With Me
1450          intRetVal = 0  ' ** User doesn't matter.
1460          If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
1470            FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
1480          Else
1490            If CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Module Function: modFileUtilities.
1500              With .Parent
1510                Select Case .chkLoadTimer_vis
                    Case True
1520                  .chkLoadTimer.Visible = False
1530                  .chkLoadTimer_lbl1.Visible = False
1540                  .cmdDelete.Caption = "&Delete Account"
1550                  .cmdDelete.StatusBarText = "Delete the current account - Alt+D"
1560                  .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Alt+D"
1570                Case False
1580                  .chkLoadTimer.Visible = True
1590                  .chkLoadTimer_lbl1.Visible = True
1600                  .cmdDelete.Caption = "Delete Account"
1610                  .cmdDelete.StatusBarText = "Delete the current account - Ctrl+D"
1620                  .cmdDelete.ControlTipText = "Delete Account" & vbCrLf & "- Ctrl+D"
1630                End Select
1640              End With
1650            End If
1660          End If
1670          .Parent.chkLoadTimer_vis = .Parent.chkLoadTimer.Visible
1680        End With
1690      Case vbKeyX
1700        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1710          With Me
1720            intRetVal = 0
1730            .Parent.JustClose_Set True  ' ** Form Procedure: frmMenu_Account.
1740            SetOption_Dev  ' ** Module Function: modUtilities.
1750            DoCmd.Close acForm, strCallingForm
1760          End With
1770        End If
1780      End Select
1790    End If

EXITP:
1800    KeyCode = intRetVal
1810    Exit Sub

ERRH:
1820    intRetVal = 0
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub Form_DblClick(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_DblClick"

1910    With Me
1920      DoCmd.SelectObject acForm, .Parent.Name, False
1930      .Parent.cmdAccountProfile.SetFocus
1940      .Parent.cmdAccountProfile_Click  ' ** Form Procedure: frmMenu_Account.
1950    End With

EXITP:
1960    Exit Sub

ERRH:
1970    Select Case ERR.Number
        Case Else
1980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1990    End Select
2000    Resume EXITP

End Sub

Private Sub Detail_DblClick(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_DblClick"

2110    Form_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub accountno_DblClick(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_DblClick"

2210    Form_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2320    intShiftDown = (Shift And acShiftMask) > 0
2330    intAltDown = (Shift And acAltMask) > 0
2340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2370        With Me
2380          intRetVal = 0
2390          .shortname.SetFocus
2400        End With
2410      End Select
2420    End If

        ' ** Shift keys.
2430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2450        With Me
2460          intRetVal = 0
2470          DoCmd.SelectObject acForm, .Parent.Name, False
2480          .Parent.cmdAccountNew.SetFocus
2490        End With
2500      End Select
2510    End If

EXITP:
2520    KeyCode = intRetVal
2530    Exit Sub

ERRH:
2540    intRetVal = 0
2550    Select Case ERR.Number
        Case Else
2560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2570    End Select
2580    Resume EXITP

End Sub

Private Sub dateclosed_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "dateclosed_DblClick"

2610    Form_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub shortname_DblClick(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_DblClick"

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

Private Sub legalname_DblClick(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_DblClick"

2810    Form_DblClick Cancel  ' ** Procedure: Above.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub legalname_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_KeyDown"

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
2980          intRetVal = 0
2990          DoCmd.SelectObject acForm, .Parent.Name, False
3000          .Parent.cmdSearch.SetFocus
3010        End With
3020      End Select
3030    End If

        ' ** Shift keys.
3040    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3050      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3060        With Me
3070          intRetVal = 0
3080          .shortname.SetFocus
3090        End With
3100      End Select
3110    End If

EXITP:
3120    KeyCode = intRetVal
3130    Exit Sub

ERRH:
3140    intRetVal = 0
3150    Select Case ERR.Number
        Case Else
3160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3170    End Select
3180    Resume EXITP

End Sub

Private Function UpdateHiddenCnt() As Boolean

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateHiddenCnt"

        Dim blnRetVal As Boolean

3210    blnRetVal = True

3220    With Me
3230      If gblnAdmin = True Then
3240        If .hidden_trans > 0& Then
3250          .Parent.cmdHideTransactions_lbl.Caption = CStr(.hidden_trans)
3260        Else
3270          .Parent.cmdHideTransactions_lbl.Caption = vbNullString
3280        End If
3290      End If
3300    End With

EXITP:
3310    UpdateHiddenCnt = blnRetVal
3320    Exit Function

ERRH:
3330    blnRetVal = False
3340    Select Case ERR.Number
        Case Else
3350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3360    End Select
3370    Resume EXITP

End Function

Public Sub SubResize_One(lngWidth_Diff As Long)
' ** It won't ever get narrower.

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "SubResize_One"

3410    With Me
3420      If lngWidth_Diff > 0& Then
3430        .Width = (lngForm_Width + lngWidth_Diff)
3440        .dateclosed.Left = (lngDateClosed_Left + lngWidth_Diff)
3450        .dateclosed_lbl.Left = (lngDateClosedLbl_left + lngWidth_Diff)
3460        .shortname.Width = (lngShortName_Width + lngWidth_Diff)
3470        .legalname.Width = (lngLegalname_Width + lngWidth_Diff)
3480        .Width = (lngForm_Width + lngWidth_Diff)
3490      Else
3500        If .Width <> lngForm_Width Or .dateclosed.Left <> lngDateClosed_Left Or .dateclosed_lbl.Left <> lngDateClosedLbl_left Or _
                .shortname.Width <> lngShortName_Width Or .legalname.Width <> lngLegalname_Width Then
3510          .Width = lngForm_Width
3520          .dateclosed.Left = lngDateClosed_Left
3530          .dateclosed_lbl.Left = lngDateClosedLbl_left
3540          .shortname.Width = lngShortName_Width
3550          .legalname.Width = lngLegalname_Width
3560          .Width = lngForm_Width
3570        End If
3580      End If
3590    End With

EXITP:
3600    Exit Sub

ERRH:
3610    Select Case ERR.Number
        Case Else
3620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3630    End Select
3640    Resume EXITP

End Sub

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

3710    Select Case intWhere
        Case acCmdRecordsGoToFirst
3720      DoCmd.RunCommand acCmdRecordsGoToFirst
3730    Case acCmdRecordsGoToLast
3740      DoCmd.RunCommand acCmdRecordsGoToLast
3750    Case acCmdRecordsGoToPrevious
3760      DoCmd.RunCommand acCmdRecordsGoToPrevious
3770    Case acCmdRecordsGoToNext
3780      DoCmd.RunCommand acCmdRecordsGoToNext
3790    Case Else
3800      If IsMissing(varID) = False Then
3810        If IsEmpty(varID) = False Then
3820          If IsNull(varID) = False Then
3830            If Trim(varID) <> vbNullString Then
3840              With Me
3850                Set rst = .RecordsetClone
3860                With rst
3870                  If .BOF = True And .EOF = True Then
                        ' ** Skip it!
3880                  Else
3890                    .MoveFirst
3900                    .FindFirst "[accountno] = '" & CStr(varID) & "'"
3910                    If .NoMatch = False Then
3920                      Me.Bookmark = .Bookmark
3930                    End If
3940                  End If
3950                  .Close
3960                End With
3970              End With
3980            End If
3990          End If
4000        End If
4010      End If
4020    End Select

EXITP:
4030    Set rst = Nothing
4040    Exit Sub

ERRH:
4050    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
4060    Case 2424  ' ** The expression you entered has a field, control, or property name that Microsoft Access can't find.
          ' ** Ignore for now.
4070    Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4080    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
4090    Case Else
4100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4110    End Select
4120    Resume EXITP

End Sub
