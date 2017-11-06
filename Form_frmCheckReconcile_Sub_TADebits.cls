VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckReconcile_Sub_TADebits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckReconcile_Sub_TADebits"

'VGC 10/03/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts on frmCheckReconcile}
' **   Number:           B {opgAccountSource_optNumber on frmCheckReconcile}
' **   Close:            C {cmdClose on frmCheckReconcile}
' **   Out Checks:       D {frmCheckReconcile_Sub_OutChecks on frmCheckReconcile}
' **   BS Debits:        E {frmCheckReconcile_Sub_BSDebits on frmCheckReconcile}
' **   Total Cash:       H {opgCheckingType_optTotalCash on frmCheckReconcile}
' **   TA Credits:       I {frmCheckReconcile_Sub_TACredits on frmCheckReconcile}
' **   Uncheck All:      K {cmdCheckNone on frmCheckReconcile}
' **   Check All:        L {cmdCheckAll on frmCheckReconcile}
' **   Name:             M {opgAccountSource_optName on frmCheckReconcile}
' **   BS Balance:       N {cracct_bsbalance_display on frmCheckReconcile}
' **   TA Debits:        O {frmCheckReconcile_Sub_TADebits on frmCheckReconcile}
' **   Print Report:     P {cmdPrintReport on frmCheckReconcile}
' **   BS Credits:       R {frmCheckReconcile_Sub_BSCredits on frmCheckReconcile}
' **   Asset:            S {cmbAssets on frmCheckReconcile}
' **   Balance Through:  T {DateEnd on frmCheckReconcile}
' **   Update:           U {cmdUpdate on frmCheckReconcile}
' **   Individual:       V {opgCheckingType_optIndividual on frmCheckReconcile}
' **   Preview Report:   W {cmdPreviewReport on frmCheckReconcile}
' **   Exit:             X {cmdClose on frmCheckReconcile}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc Tots:      F5 {RecalcTots_CR}
' **   Date Picker:      F7 {cmdCalendar on frmCheckReconcile}

' ** Shortcut Ctrl keys responsive from this form:
' **   2nd Asset:        A {cmdAddAsset on frmCheckReconcile}
' **   Delete:           D {cmdDelete on frmCheckReconcile}
' **   Next Asset:       N {cmdAssetNext on frmCheckReconcile}
' **   Previous Asset:   P {cmdAssetPrevious on frmCheckReconcile}
' **   Clear:            R {cmdClear on frmCheckReconcile}
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Combo box column constants: cmbAccounts.
Private Const CBX_ACT_ACTNO   As Integer = 0
Private Const CBX_ACT_DESC    As Integer = 1
'Private Const CBX_ACT_STMDATE As Integer = 2
'Private Const CBX_ACT_SHORT   As Integer = 3
'Private Const CBX_ACT_LEGAL   As Integer = 4
'Private Const CBX_ACT_BALDATE As Integer = 5
'Private Const CBX_ACT_HASREL  As Integer = 6
'Private Const CBX_ACT_DEFAST  As Integer = 7

Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmCheckReconcile" Then

160         gstrFormQuerySpec = strCallingForm

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

190       Else
200         Cancel = -1
210       End If
220     End With

230     If Cancel = -1 Then
240       Beep
250       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
260     End If

EXITP:
270     Exit Sub

ERRH:
280     DoCmd.Hourglass False
290     Select Case ERR.Number
        Case Else
300       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
310     End Select
320     Resume EXITP

End Sub

Private Sub Form_Current()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

410     With Me
420       If IsNull(.Parent.cmbAccounts) = False Then
430         If IsNull(.Parent.cmbAssets) = False Then
440           If .Parent.cmbAssets > 0 Then
450             .Parent.TATotalDebits.Requery
460             lngRecsCur = RecCnt  ' ** Procedure: Below.
470             .Parent.TADebitsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
480             UpdateTATotal_CR .Parent  ' ** Module Procedure: modCheckReconcile.
490           End If
500         End If
510       End If
520     End With

EXITP:
530     Exit Sub

ERRH:
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strAccountNo As String, lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality.
        ' **   Recalc Tots:      F5 {RecalcTots_CR}
        ' **   Date Picker:      F7 {cmdCalendar on frmCheckReconcile}

        ' ** Shortcut Alt to other forms and functionality.
        ' **   Out Checks:       D {frmCheckReconcile_Sub_OutChecks on frmCheckReconcile}
        ' **   BS Debits:        E {frmCheckReconcile_Sub_BSDebits on frmCheckReconcile}
        ' **   TA Credits:       I {frmCheckReconcile_Sub_TACredits on frmCheckReconcile}
        ' **   BS Credits:       R {frmCheckReconcile_Sub_BSCredits on frmCheckReconcile}

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   2nd Asset:        A {cmdAddAsset on frmCheckReconcile}
        ' **   Delete:           D {cmdDelete on frmCheckReconcile}
        ' **   Next Asset:       N {cmdAssetNext on frmCheckReconcile}
        ' **   Previous Asset:   P {cmdAssetPrevious on frmCheckReconcile}
        ' **   Clear:            R {cmdClear on frmCheckReconcile}
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyF5
670         With Me
680           DoCmd.Hourglass True
690           DoEvents
700           strAccountNo = .Parent.cmbAccounts.Column(CBX_ACT_ACTNO)
710           arr_varAsset_ThisAcct = .Parent.AssetsThisArray_Get  ' ** Form Function: frmCheckReconcile.
720           lngAssets_ThisAcct = UBound(arr_varAsset_ThisAcct, 2) + 1&
730           RecalcTots_CR strAccountNo, lngAssets_ThisAcct, arr_varAsset_ThisAcct, .Parent  ' ** Module Procedure: modCheckReconcile.
740           DoCmd.Hourglass False
750         End With
760       Case vbKeyF7
770         With Me
780           intRetVal = 0
790           If .Parent.cmdCalendar.Enabled = True Then
800             DoCmd.SelectObject acForm, .Parent.Name, False
810             .Parent.cmdCalendar.SetFocus
820             .Parent.cmdCalendar_Click  ' ** Form Procedure: frmCheckReconcile.
830           Else
840             Beep
850           End If
860         End With
870       Case vbKeyUp
880         With Me
890           intRetVal = 0
900           If .CurrentRecord > 1 Then
910             MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
920           End If
930         End With
940       Case vbKeyDown
950         With Me
960           intRetVal = 0
970           lngRecsCur = RecCnt  ' ** Function: Below.
980           If .CurrentRecord < lngRecsCur Then
990             MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1000          End If
1010        End With
1020      End Select
1030    End If

        ' ** Alt keys.
1040    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1050      Select Case intRetVal
          Case vbKeyD
1060        With Me
1070          intRetVal = 0
1080          If .Parent.frmCheckReconcile_Sub_OutChecks.Enabled = True Then
1090            DoCmd.SelectObject acForm, .Parent.Name, False
1100            .Parent.frmCheckReconcile_Sub_OutChecks.SetFocus
1110          Else
1120            Beep
1130          End If
1140        End With
1150      Case vbKeyE
1160        With Me
1170          intRetVal = 0
1180          If .Parent.frmCheckReconcile_Sub_BSDebits.Enabled = True Then
1190            DoCmd.SelectObject acForm, .Parent.Name, False
1200            lngRecsCur = .Parent.frmCheckReconcile_Sub_BSDebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSDebits.
1210            .Parent.frmCheckReconcile_Sub_BSDebits.SetFocus
1220            If lngRecsCur > 0& Then
1230              .Parent.frmCheckReconcile_Sub_BSDebits.Form.FocusHolder.SetFocus
1240            Else
1250              .Parent.frmCheckReconcile_Sub_BSDebits.Form.critem_description.SetFocus
1260            End If
1270          Else
1280            Beep
1290          End If
1300        End With
1310      Case vbKeyI
1320        With Me
1330          intRetVal = 0
1340          If .Parent.frmCheckReconcile_Sub_TACredits.Enabled = True Then
1350            DoCmd.SelectObject acForm, .Parent.Name, False
1360            lngRecsCur = .Parent.frmCheckReconcile_Sub_BSCredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSCredits.
1370            .Parent.frmCheckReconcile_Sub_TACredits.SetFocus
1380            If lngRecsCur > 0& Then
1390              .Parent.frmCheckReconcile_Sub_TACredits.Form.FocusHolder.SetFocus
1400            Else
1410              .Parent.frmCheckReconcile_Sub_TACredits.Form.critem_description.SetFocus
1420            End If
1430          Else
1440            Beep
1450          End If
1460        End With
1470      Case vbKeyR
1480        With Me
1490          intRetVal = 0
1500          If .Parent.frmCheckReconcile_Sub_BSCredits.Enabled = True Then
1510            DoCmd.SelectObject acForm, .Parent.Name, False
1520            lngRecsCur = .Parent.frmCheckReconcile_Sub_BSCredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSCredits.
1530            .Parent.frmCheckReconcile_Sub_BSCredits.SetFocus
1540            If lngRecsCur > 0& Then
1550              .Parent.frmCheckReconcile_Sub_BSCredits.Form.FocusHolder.SetFocus
1560            Else
1570              .Parent.frmCheckReconcile_Sub_BSCredits.Form.critem_description.SetFocus
1580            End If
1590          Else
1600            Beep
1610          End If
1620        End With
1630      Case vbKeyX
1640        With Me
1650          intRetVal = 0
1660          DoCmd.SelectObject acForm, .Parent.Name, False
1670          .Parent.cmdClose.SetFocus
1680          .Parent.cmdClose_Click  ' ** Form Procedure: frmCheckReconcile.
1690        End With
1700      End Select
1710    End If

        ' ** Ctrl keys.
1720    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1730      Select Case intRetVal
          Case vbKeyA
1740        With Me
1750          intRetVal = 0
1760          Select Case .Parent.cmdAddAsset.Enabled
              Case True
1770            DoCmd.SelectObject acForm, .Parent.Name, False
1780            .Parent.cmdAddAsset.SetFocus
1790            .Parent.cmdAddAsset_Click  ' ** Form Procedure: frmCheckReconcile.
1800          Case False
1810            Beep
1820          End Select
1830        End With
1840      Case vbKeyD
1850        With Me
1860          intRetVal = 0
1870          Select Case .Parent.cmdDelete.Enabled
              Case True
1880            DoCmd.SelectObject acForm, .Parent.Name, False
1890            .Parent.cmdDelete.SetFocus
1900            .Parent.cmdDelete_Click  ' ** Form Procedure: frmCheckReconcile.
1910          Case False
1920            Beep
1930          End Select
1940        End With
1950      Case vbKeyN
1960        With Me
1970          intRetVal = 0
1980          Select Case .Parent.cmdAssetNext.Enabled
              Case True
1990            DoCmd.SelectObject acForm, .Parent.Name, False
2000            .Parent.cmdAssetNext.SetFocus
2010            .Parent.cmdAssetNext_Click  ' ** Form Procedure: frmCheckReconcile.
2020          Case False
2030            Beep
2040          End Select
2050        End With
2060      Case vbKeyP
2070        With Me
2080          intRetVal = 0
2090          Select Case .Parent.cmdAssetPrevious.Enabled
              Case True
2100            DoCmd.SelectObject acForm, .Parent.Name, False
2110            .Parent.cmdAssetPrevious.SetFocus
2120            .Parent.cmdAssetPrevious_Click  ' ** Form Procedure: frmCheckReconcile.
2130          Case False
2140            Beep
2150          End Select
2160        End With
2170      Case vbKeyR
2180        With Me
2190          intRetVal = 0
2200          Select Case .Parent.cmdClear.Enabled
              Case True
2210            DoCmd.SelectObject acForm, .Parent.Name, False
2220            .Parent.cmdClear.SetFocus
2230            .Parent.cmdClear_Click  ' ** Form Procedure: frmCheckReconcile.
2240          Case False
2250            Beep
2260          End Select
2270        End With
2280      Case vbKeyS
2290        intRetVal = 0
2300        cmdSave_Click  ' ** Procedure: Below.
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
2400      End Select
2410    End If

EXITP:
2420    KeyCode = intRetVal
2430    Exit Sub

ERRH:
2440    intRetVal = 0
2450    Select Case ERR.Number
        Case Else
2460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2470    End Select
2480    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

        Dim strAccountNo As String, lngAssetNo As Long

2510    With Me
2520      strAccountNo = vbNullString: lngAssetNo = 0&
2530      If IsNull(.Parent.cmbAccounts) = False Then
2540        strAccountNo = .Parent.cmbAccounts
2550      End If
2560      If IsNull(.Parent.cmbAssets) = False Then
2570        lngAssetNo = .Parent.cmbAssets
2580      End If
2590      If strAccountNo = vbNullString Or (strAccountNo <> "CRTC01" And lngAssetNo = 0&) Then
2600        Cancel = -1
2610      Else
2620        If IsNull(.accountno) = True Then
2630          .accountno = strAccountNo
2640        End If
2650        If IsNull(.assetno) = True Then
2660          .assetno = lngAssetNo
2670        Else
2680          If .assetno = 0 Then
2690            .assetno = lngAssetNo
2700          End If
2710        End If
2720        .critem_datemodified = Now()
2730      End If
2740      If IsNull(.cracct_id) = True Then
2750        .cracct_id = .Parent.cracct_id
2760      End If
2770    End With

EXITP:
2780    Exit Sub

ERRH:
2790    Select Case ERR.Number
        Case Else
2800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2810    End Select
2820    Resume EXITP

End Sub

Private Sub Form_AfterUpdate()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_AfterUpdate"

2910    With Me
2920      .Parent.TATotalDebits.Requery
2930      lngRecsCur = RecCnt  ' ** Procedure: Below.
2940      .Parent.TADebitsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
2950      UpdateTATotal_CR .Parent  ' ** Module Procedure: modCheckReconcile.
2960      .Parent.Recalc
2970    End With

EXITP:
2980    Exit Sub

ERRH:
2990    Select Case ERR.Number
        Case Else
3000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3010    End Select
3020    Resume EXITP

End Sub

Private Sub Form_AfterDelConfirm(STATUS As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_AfterDelConfirm"

3110    With Me
3120      .Parent.TATotalDebits.Requery
3130      lngRecsCur = RecCnt  ' ** Procedure: Below.
3140      .Parent.TADebitsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
3150      UpdateTATotal_CR .Parent  ' ** Module Procedure: modCheckReconcile.
3160    End With

EXITP:
3170    Exit Sub

ERRH:
3180    Select Case ERR.Number
        Case Else
3190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3200    End Select
3210    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3310    If DataErr <> 0 Then
3320      Select Case DataErr
          Case Else
3330        Response = acDataErrDisplay
            'zErrorHandler THIS_NAME, THIS_PROC, DataErr, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3340      End Select
3350    End If

3360    gstrFormQuerySpec = THIS_NAME

EXITP:
3370    Exit Sub

ERRH:
3380    Select Case ERR.Number
        Case Else
3390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3400    End Select
3410    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3510    DoCmd.RunCommand acCmdSaveRecord
3520    DoEvents

EXITP:
3530    Exit Sub

ERRH:
3540    Select Case ERR.Number
        Case Else
3550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3560    End Select
3570    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
3670        With Me
3680          intRetVal = 0
3690          .critem_description.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3750        With Me
3760          intRetVal = 0
3770          If .CurrentRecord > 1 Then
3780            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3790            .critem_amount.SetFocus
3800          Else
3810            DoCmd.SelectObject acForm, .Parent.Name, False
3820            lngRecsCur = .Parent.frmCheckReconcile_Sub_TACredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TACredits.
3830            .Parent.frmCheckReconcile_Sub_TACredits.SetFocus
3840            If lngRecsCur > 0& Then
3850              .Parent.frmCheckReconcile_Sub_TACredits.Form.FocusHolder.SetFocus
3860            Else
3870              .Parent.frmCheckReconcile_Sub_TACredits.Form.critem_description.SetFocus
3880            End If
3890          End If
3900        End With
3910      End Select
3920    End If

EXITP:
3930    KeyCode = intRetVal
3940    Exit Sub

ERRH:
3950    intRetVal = 0
3960    Select Case ERR.Number
        Case Else
3970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3980    End Select
3990    Resume EXITP

End Sub

Private Sub critem_description_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "critem_description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4020    intShiftDown = (Shift And acShiftMask) > 0
4030    intAltDown = (Shift And acAltMask) > 0
4040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4070        With Me
4080          intRetVal = 0
4090          .critem_amount.SetFocus
4100        End With
4110      End Select
4120    End If

        ' ** Shift keys.
4130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4150        With Me
4160          intRetVal = 0
4170          If .CurrentRecord > 1 Then
4180            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4190            .critem_amount.SetFocus
4200          Else
4210            DoCmd.SelectObject acForm, .Parent.Name, False
4220            lngRecsCur = .Parent.frmCheckReconcile_Sub_TACredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TACredits.
4230            .Parent.frmCheckReconcile_Sub_TACredits.SetFocus
4240            If lngRecsCur > 0& Then
4250              .Parent.frmCheckReconcile_Sub_TACredits.Form.FocusHolder.SetFocus
4260            Else
4270              .Parent.frmCheckReconcile_Sub_TACredits.Form.critem_description.SetFocus
4280            End If
4290          End If
4300        End With
4310      End Select
4320    End If

EXITP:
4330    KeyCode = intRetVal
4340    Exit Sub

ERRH:
4350    intRetVal = 0
4360    Select Case ERR.Number
        Case Else
4370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4380    End Select
4390    Resume EXITP

End Sub

Private Sub critem_amount_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "critem_amount_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4420    intShiftDown = (Shift And acShiftMask) > 0
4430    intAltDown = (Shift And acAltMask) > 0
4440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4470        With Me
4480          intRetVal = 0
4490          lngRecsCur = RecCnt  ' ** Function: Below.
4500          If .CurrentRecord < lngRecsCur Then
4510            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4520            .critem_description.SetFocus
4530          Else
4540            Select Case IsNull(.critem_id)
                Case True
4550              DoCmd.SelectObject acForm, .Parent.Name, False
4560              .Parent.frmCheckReconcile_Sub_OutChecks.SetFocus
4570            Case False
4580              MoveRec acCmdRecordsGoToNew  ' ** Procedure: Below.
4590              .critem_description.SetFocus
4600            End Select
4610          End If
4620        End With
4630      End Select
4640    End If

        ' ** Shift keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4670        With Me
4680          intRetVal = 0
4690          .critem_description.SetFocus
4700        End With
4710      End Select
4720    End If

EXITP:
4730    KeyCode = intRetVal
4740    Exit Sub

ERRH:
4750    intRetVal = 0
4760    Select Case ERR.Number
        Case Else
4770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4780    End Select
4790    Resume EXITP

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
5190    Case acCmdRecordsGoToNew
5200      cmdSave_Click  ' ** Procedure: Above.
5210      DoEvents
5220      DoCmd.RunCommand acCmdRecordsGoToNew
5230    Case Else
5240      If IsMissing(varID) = False Then
5250        With Me
5260          Set rst = .RecordsetClone
5270          With rst
5280            .FindFirst "[critem_id] = " & CStr(varID)
5290            If .NoMatch = False Then
5300              Me.Bookmark = .Bookmark
5310            End If
5320            .Close
5330          End With
5340        End With
5350      End If
5360    End Select

EXITP:
5370    Set rst = Nothing
5380    Exit Sub

ERRH:
5390    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5400    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5410    Case Else
5420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5430    End Select
5440    Resume EXITP

End Sub
