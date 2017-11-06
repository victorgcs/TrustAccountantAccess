VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckReconcile_Sub_BSCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckReconcile_Sub_BSCredits"

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

        Dim lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant

410     With Me
420       If IsNull(.Parent.cmbAccounts) = False Then
430         If IsNull(.Parent.cmbAssets) = False Then
440           If .Parent.cmbAssets > 0 Then
450             .Parent.BSTotalCredits.Requery
                ' ** cmbAccounts_AfterUpdate() loads array, so it must have been run before getting here.
460             arr_varAsset_ThisAcct = .Parent.AssetsThisArray_Get  ' ** Form Function: frmCheckReconcile.
470   On Error Resume Next
480             lngAssets_ThisAcct = UBound(arr_varAsset_ThisAcct, 2) + 1&
490             If ERR.Number = 0 Then
500   On Error GoTo ERRH
510               lngRecsCur = RecCnt  ' ** Procedure: Below.
520               .Parent.BSCreditsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
530               UpdateBSTotal_CR .Parent, lngAssets_ThisAcct, arr_varAsset_ThisAcct, THIS_NAME & "." & THIS_PROC  ' ** Module Procedure: modCheckReconcile.
540             Else
550   On Error GoTo ERRH
560             End If
570           End If
580         End If
590       End If
600     End With

EXITP:
610     Exit Sub

ERRH:
620     Select Case ERR.Number
        Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strAccountNo As String, lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality.
        ' **   Recalc Tots:      F5 {RecalcTots_CR}
        ' **   Date Picker:      F7 {cmdCalendar on frmCheckReconcile}

        ' ** Shortcut Alt to other forms and functionality.
        ' **   Out Checks:       D {frmCheckReconcile_Sub_OutChecks on frmCheckReconcile}
        ' **   BS Debits:        E {frmCheckReconcile_Sub_BSDebits on frmCheckReconcile}
        ' **   TA Credits:       I {frmCheckReconcile_Sub_TACredits on frmCheckReconcile}
        ' **   TA Debits:        O {frmCheckReconcile_Sub_TADebits on frmCheckReconcile}

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   2nd Asset:        A {cmdAddAsset on frmCheckReconcile}
        ' **   Delete:           D {cmdDelete on frmCheckReconcile}
        ' **   Next Asset:       N {cmdAssetNext on frmCheckReconcile}
        ' **   Previous Asset:   P {cmdAssetPrevious on frmCheckReconcile}
        ' **   Clear:            R {cmdClear on frmCheckReconcile}
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyF5
770         With Me
780           DoCmd.Hourglass True
790           DoEvents
800           strAccountNo = .Parent.cmbAccounts.Column(CBX_ACT_ACTNO)
810           arr_varAsset_ThisAcct = .Parent.AssetsThisArray_Get  ' ** Form Function: frmCheckReconcile.
820           lngAssets_ThisAcct = UBound(arr_varAsset_ThisAcct, 2) + 1&
830           RecalcTots_CR strAccountNo, lngAssets_ThisAcct, arr_varAsset_ThisAcct, Me  ' ** Module Procedure: modCheckReconcile.
840           DoCmd.Hourglass False
850         End With
860       Case vbKeyF7
870         With Me
880           intRetVal = 0
890           If .Parent.cmdCalendar.Enabled = True Then
900             DoCmd.SelectObject acForm, .Parent.Name, False
910             .Parent.cmdCalendar.SetFocus
920             .Parent.cmdCalendar_Click  ' ** Form Procedure: frmCheckReconcile.
930           Else
940             Beep
950           End If
960         End With
970       Case vbKeyUp
980         With Me
990           intRetVal = 0
1000          If .CurrentRecord > 1 Then
1010            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1020          End If
1030        End With
1040      Case vbKeyDown
1050        With Me
1060          intRetVal = 0
1070          lngRecsCur = RecCnt  ' ** Function: Below.
1080          If .CurrentRecord < lngRecsCur Then
1090            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1100          End If
1110        End With
1120      End Select
1130    End If

        ' ** Alt keys.
1140    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1150      Select Case intRetVal
          Case vbKeyD
1160        With Me
1170          intRetVal = 0
1180          If .Parent.frmCheckReconcile_Sub_OutChecks.Enabled = True Then
1190            DoCmd.SelectObject acForm, .Parent.Name, False
1200            .Parent.frmCheckReconcile_Sub_OutChecks.SetFocus
1210          Else
1220            Beep
1230          End If
1240        End With
1250      Case vbKeyE
1260        With Me
1270          intRetVal = 0
1280          If .Parent.frmCheckReconcile_Sub_BSDebits.Enabled = True Then
1290            DoCmd.SelectObject acForm, .Parent.Name, False
1300            lngRecsCur = .Parent.frmCheckReconcile_Sub_BSDebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSDebits.
1310            .Parent.frmCheckReconcile_Sub_BSDebits.SetFocus
1320            If lngRecsCur > 0& Then
1330              .Parent.frmCheckReconcile_Sub_BSDebits.Form.FocusHolder.SetFocus
1340            Else
1350              .Parent.frmCheckReconcile_Sub_BSDebits.Form.critem_description.SetFocus
1360            End If
1370          Else
1380            Beep
1390          End If
1400        End With
1410      Case vbKeyI
1420        With Me
1430          intRetVal = 0
1440          If .Parent.frmCheckReconcile_Sub_TACredits.Enabled = True Then
1450            DoCmd.SelectObject acForm, .Parent.Name, False
1460            lngRecsCur = .Parent.frmCheckReconcile_Sub_TACredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TACredits.
1470            .Parent.frmCheckReconcile_Sub_TACredits.SetFocus
1480            If lngRecsCur > 0& Then
1490              .Parent.frmCheckReconcile_Sub_TACredits.Form.FocusHolder.SetFocus
1500            Else
1510              .Parent.frmCheckReconcile_Sub_TACredits.critem_description.SetFocus
1520            End If
1530          Else
1540            Beep
1550          End If
1560        End With
1570      Case vbKeyO
1580        With Me
1590          intRetVal = 0
1600          If .Parent.frmCheckReconcile_Sub_TADebits.Enabled = True Then
1610            DoCmd.SelectObject acForm, .Parent.Name, False
1620            lngRecsCur = .Parent.frmCheckReconcile_Sub_TACredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TACredits.
1630            .Parent.frmCheckReconcile_Sub_TADebits.SetFocus
1640            If lngRecsCur > 0& Then
1650              .Parent.frmCheckReconcile_Sub_TACredits.Form.FocusHolder.SetFocus
1660            Else
1670              .Parent.frmCheckReconcile_Sub_TACredits.Form.critem_description.SetFocus
1680            End If
1690          Else
1700            Beep
1710          End If
1720        End With
1730      Case vbKeyX
1740        With Me
1750          intRetVal = 0
1760          DoCmd.SelectObject acForm, .Parent.Name, False
1770          .Parent.cmdClose.SetFocus
1780          .Parent.cmdClose_Click  ' ** Form Procedure: frmCheckReconcile.
1790        End With
1800      End Select
1810    End If

        ' ** Ctrl keys.
1820    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1830      Select Case intRetVal
          Case vbKeyA
1840        With Me
1850          intRetVal = 0
1860          Select Case .Parent.cmdAddAsset.Enabled
              Case True
1870            DoCmd.SelectObject acForm, .Parent.Name, False
1880            .Parent.cmdAddAsset.SetFocus
1890            .Parent.cmdAddAsset_Click  ' ** Form Procedure: frmCheckReconcile.
1900          Case False
1910            Beep
1920          End Select
1930        End With
1940      Case vbKeyD
1950        With Me
1960          intRetVal = 0
1970          Select Case .Parent.cmdDelete.Enabled
              Case True
1980            DoCmd.SelectObject acForm, .Parent.Name, False
1990            .Parent.cmdDelete.SetFocus
2000            .Parent.cmdDelete_Click  ' ** Form Procedure: frmCheckReconcile.
2010          Case False
2020            Beep
2030          End Select
2040        End With
2050      Case vbKeyN
2060        With Me
2070          intRetVal = 0
2080          Select Case .Parent.cmdAssetNext.Enabled
              Case True
2090            DoCmd.SelectObject acForm, .Parent.Name, False
2100            .Parent.cmdAssetNext.SetFocus
2110            .Parent.cmdAssetNext_Click  ' ** Form Procedure: frmCheckReconcile.
2120          Case False
2130            Beep
2140          End Select
2150        End With
2160      Case vbKeyP
2170        With Me
2180          intRetVal = 0
2190          Select Case .Parent.cmdAssetPrevious.Enabled
              Case True
2200            DoCmd.SelectObject acForm, .Parent.Name, False
2210            .Parent.cmdAssetPrevious.SetFocus
2220            .Parent.cmdAssetPrevious_Click  ' ** Form Procedure: frmCheckReconcile.
2230          Case False
2240            Beep
2250          End Select
2260        End With
2270      Case vbKeyR
2280        With Me
2290          intRetVal = 0
2300          Select Case .Parent.cmdClear.Enabled
              Case True
2310            DoCmd.SelectObject acForm, .Parent.Name, False
2320            .Parent.cmdClear.SetFocus
2330            .Parent.cmdClear_Click  ' ** Form Procedure: frmCheckReconcile.
2340          Case False
2350            Beep
2360          End Select
2370        End With
2380      Case vbKeyS
2390        intRetVal = 0
2400        cmdSave_Click  ' ** Procedure: Below.
2410      End Select
2420    End If

        ' ** Ctrl-Shift keys.
2430    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2440      Select Case intRetVal
          Case vbKeyF
2450        With Me
2460          intRetVal = 0
2470          DoCmd.SelectObject acForm, .Parent.Name, False
2480          .Parent.FocusHolder.SetFocus
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

Private Sub Form_BeforeUpdate(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

        Dim strAccountNo As String, lngAssetNo As Long

2610    With Me
2620      strAccountNo = vbNullString: lngAssetNo = 0&
2630      If IsNull(.Parent.cmbAccounts) = False Then
2640        strAccountNo = .Parent.cmbAccounts
2650      End If
2660      If IsNull(.Parent.cmbAssets) = False Then
2670        lngAssetNo = .Parent.cmbAssets
2680      End If
2690      If strAccountNo = vbNullString Or (strAccountNo <> "CRTC01" And lngAssetNo = 0&) Then
2700        Cancel = -1
2710      Else
2720        If IsNull(.accountno) = True Then
2730          .accountno = strAccountNo
2740        End If
2750        If IsNull(.assetno) = True Then
2760          .assetno = lngAssetNo
2770        Else
2780          If .assetno = 0 Then
2790            .assetno = lngAssetNo
2800          End If
2810        End If
2820        .critem_datemodified = Now()
2830      End If
2840      If IsNull(.cracct_id) = True Then
2850        .cracct_id = .Parent.cracct_id
2860      End If
2870    End With

EXITP:
2880    Exit Sub

ERRH:
2890    Select Case ERR.Number
        Case Else
2900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2910    End Select
2920    Resume EXITP

End Sub

Private Sub Form_AfterUpdate()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_AfterUpdate"

        Dim lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant

3010    With Me
3020      .Parent.BSTotalCredits.Requery
3030      arr_varAsset_ThisAcct = .Parent.AssetsThisArray_Get  ' ** Form Function: frmCheckReconcile.
3040      lngAssets_ThisAcct = UBound(arr_varAsset_ThisAcct, 2) + 1&
3050      lngRecsCur = RecCnt  ' ** Procedure: Below.
3060      .Parent.BSCreditsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
3070      UpdateBSTotal_CR .Parent, lngAssets_ThisAcct, arr_varAsset_ThisAcct, THIS_NAME & "." & THIS_PROC  ' ** Module Procedure: modCheckReconcile.
3080      .Parent.Recalc
3090    End With

EXITP:
3100    Exit Sub

ERRH:
3110    Select Case ERR.Number
        Case Else
3120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3130    End Select
3140    Resume EXITP

End Sub

Private Sub Form_AfterDelConfirm(STATUS As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_AfterDelConfirm"

        Dim lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant

3210    With Me
3220      .Parent.BSTotalCredits.Requery
3230      arr_varAsset_ThisAcct = .Parent.AssetsThisArray_Get  ' ** Form Function: frmCheckReconcile.
3240      lngAssets_ThisAcct = UBound(arr_varAsset_ThisAcct, 2) + 1&
3250      lngRecsCur = RecCnt  ' ** Procedure: Below.
3260      .Parent.BSCreditsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
3270      UpdateBSTotal_CR .Parent, lngAssets_ThisAcct, arr_varAsset_ThisAcct, THIS_NAME & "." & THIS_PROC  ' ** Module Procedure: modCheckReconcile.
3280    End With

EXITP:
3290    Exit Sub

ERRH:
3300    Select Case ERR.Number
        Case Else
3310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3320    End Select
3330    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3410    If DataErr <> 0 Then
3420      Select Case DataErr
          Case Else
3430        Response = acDataErrDisplay
            'zErrorHandler THIS_NAME, THIS_PROC, DataErr, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3440      End Select
3450    End If

3460    gstrFormQuerySpec = THIS_NAME

EXITP:
3470    Exit Sub

ERRH:
3480    Select Case ERR.Number
        Case Else
3490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3500    End Select
3510    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3610    DoCmd.RunCommand acCmdSaveRecord
3620    DoEvents

EXITP:
3630    Exit Sub

ERRH:
3640    Select Case ERR.Number
        Case Else
3650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3660    End Select
3670    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3720    intShiftDown = (Shift And acShiftMask) > 0
3730    intAltDown = (Shift And acAltMask) > 0
3740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3770        With Me
3780          intRetVal = 0
3790          .critem_description.SetFocus
3800        End With
3810      End Select
3820    End If

        ' ** Shift keys.
3830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3850        With Me
3860          intRetVal = 0
3870          If .CurrentRecord > 1 Then
3880            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3890            .critem_amount.SetFocus
3900          Else
3910            DoCmd.SelectObject acForm, .Parent.Name, False
3920            .Parent.cracct_bsbalance_display.SetFocus
3930          End If
3940        End With
3950      End Select
3960    End If

EXITP:
3970    KeyCode = intRetVal
3980    Exit Sub

ERRH:
3990    intRetVal = 0
4000    Select Case ERR.Number
        Case Else
4010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4020    End Select
4030    Resume EXITP

End Sub

Private Sub critem_description_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "critem_description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4120    intShiftDown = (Shift And acShiftMask) > 0
4130    intAltDown = (Shift And acAltMask) > 0
4140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4170        With Me
4180          intRetVal = 0
4190          .critem_amount.SetFocus
4200        End With
4210      End Select
4220    End If

        ' ** Shift keys.
4230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4250        With Me
4260          intRetVal = 0
4270          If .CurrentRecord > 1 Then
4280            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4290            .critem_amount.SetFocus
4300          Else
4310            DoCmd.SelectObject acForm, .Parent.Name, False
4320            .Parent.cracct_bsbalance_display.SetFocus
4330          End If
4340        End With
4350      End Select
4360    End If

EXITP:
4370    KeyCode = intRetVal
4380    Exit Sub

ERRH:
4390    intRetVal = 0
4400    Select Case ERR.Number
        Case Else
4410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4420    End Select
4430    Resume EXITP

End Sub

Private Sub critem_amount_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "critem_amount_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4570        With Me
4580          intRetVal = 0
4590          lngRecsCur = RecCnt  ' ** Function: Below.
4600          If .CurrentRecord < lngRecsCur Then
4610            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4620            .critem_description.SetFocus
4630          Else
4640            Select Case IsNull(.critem_id)
                Case True
4650              DoCmd.SelectObject acForm, .Parent.Name, False
4660              lngRecsCur = .Parent.frmCheckReconcile_Sub_BSDebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSDebits.
4670              .Parent.frmCheckReconcile_Sub_BSDebits.SetFocus
4680              If lngRecsCur > 0& Then
4690                .Parent.frmCheckReconcile_Sub_BSDebits.Form.FocusHolder.SetFocus
4700              Else
4710                .Parent.frmCheckReconcile_Sub_BSDebits.Form.critem_description.SetFocus
4720              End If
4730            Case False
4740              MoveRec acCmdRecordsGoToNew  ' ** Procedure: Below.
4750              .critem_description.SetFocus
4760            End Select
4770          End If
4780        End With
4790      End Select
4800    End If

        ' ** Shift keys.
4810    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4820      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4830        With Me
4840          intRetVal = 0
4850          .critem_description.SetFocus
4860        End With
4870      End Select
4880    End If

EXITP:
4890    KeyCode = intRetVal
4900    Exit Sub

ERRH:
4910    intRetVal = 0
4920    Select Case ERR.Number
        Case Else
4930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4940    End Select
4950    Resume EXITP

End Sub

Public Function RecCnt() As Long

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

5010    lngRetVal = 0&

5020    With Me
5030      Set rst = .RecordsetClone
5040      With rst
5050        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
5060        Else
5070          .MoveLast
5080          lngRetVal = .RecordCount
5090        End If
5100        .Close
5110      End With
5120    End With

EXITP:
5130    Set rst = Nothing
5140    RecCnt = lngRetVal
5150    Exit Function

ERRH:
5160    lngRetVal = 0&
5170    Select Case ERR.Number
        Case Else
5180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5190    End Select
5200    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

5310    Select Case intWhere
        Case acCmdRecordsGoToFirst
5320      DoCmd.RunCommand acCmdRecordsGoToFirst
5330    Case acCmdRecordsGoToLast
5340      DoCmd.RunCommand acCmdRecordsGoToLast
5350    Case acCmdRecordsGoToPrevious
5360      DoCmd.RunCommand acCmdRecordsGoToPrevious
5370    Case acCmdRecordsGoToNext
5380      DoCmd.RunCommand acCmdRecordsGoToNext
5390    Case acCmdRecordsGoToNew
5400      cmdSave_Click  ' ** Procedure: Above.
5410      DoEvents
5420      DoCmd.RunCommand acCmdRecordsGoToNew
5430    Case Else
5440      If IsMissing(varID) = False Then
5450        With Me
5460          Set rst = .RecordsetClone
5470          With rst
5480            .FindFirst "[critem_id] = " & CStr(varID)
5490            If .NoMatch = False Then
5500              Me.Bookmark = .Bookmark
5510            End If
5520            .Close
5530          End With
5540        End With
5550      End If
5560    End Select

EXITP:
5570    Set rst = Nothing
5580    Exit Sub

ERRH:
5590    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5600    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5610    Case Else
5620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5630    End Select
5640    Resume EXITP

End Sub
