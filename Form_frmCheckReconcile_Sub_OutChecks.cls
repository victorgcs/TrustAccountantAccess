VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckReconcile_Sub_OutChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckReconcile_Sub_OutChecks"

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

Private Const ASSIGN_MSG As String = "Associated with CUSIP:"

Private Const strSortOrig As String = "[transdate] DESC, [croutchk_id]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine1 As String = "Sort_line1"
Private Const strSortLine2 As String = "Sort_line2"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngRecsCur As Long, lngTpp As Long, strCusip As String
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

170         .transdate_left = .transdate.Left
180         lngSortLbl_Width = .Sort_lbl.Width

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         CLR_DISABLED_FG = CLR_DKGRY
210         CLR_DISABLED_BG = CLR_LTTEAL

220         .Filter = vbNullString
230         .FilterOn = False

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
330     DoCmd.Hourglass False
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .transdate.ForeColor = CLR_DISABLED_FG
430       .transdate.BackColor = CLR_DISABLED_BG
440       .croutchk_payee.ForeColor = CLR_DISABLED_FG
450       .croutchk_payee.BackColor = CLR_DISABLED_BG
460       .croutchk_amount.ForeColor = CLR_DISABLED_FG
470       .croutchk_amount.BackColor = CLR_DISABLED_BG
480     End With

490     SortNow THIS_PROC  ' ** Procedure: Below.

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

        Dim strAccountNo As String, strQryName As String
        Dim lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant
        Dim varTmp00 As Variant, varTmp01 As Variant, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

610     With Me
620       If IsNull(.Parent.cmbAccounts) = False Then
630         If Trim(.Parent.cmbAccounts) <> vbNullString Then
640           strAccountNo = .Parent.cmbAccounts
650           gstrAccountNo = strAccountNo

660           lngTmp02 = .Parent.AssetNumGet(3, 0)  ' ** Form Procedure: frmCheckReconcile.
670           Select Case lngTmp02
              Case 1&
                ' ** For accounts with only 1 asset.
680             Select Case strAccountNo
                Case "CRTC01"
                  ' ** tblCheckReconcile_Check, grouped and summed, forCheckPaid = False,
                  ' ** croutchk_assign = 0, accountno = 'CRTC01'.
690               strQryName = "qryCheckReconcile_OutChecks_09_04"
700             Case Else
                  ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                  ' ** croutchk_assign = 0, by specified GlobalVarGet("gstrAccountNo").
710               strQryName = "qryCheckReconcile_OutChecks_09_01"
720             End Select
730           Case Else
740             lngTmp03 = .Parent.AssetNumGet(1, 1)  ' ** Form Procedure: frmCheckReconcile.
750             lngTmp04 = .Parent.AssetNumGet(1, 2)  ' ** Form Procedure: frmCheckReconcile.
760             Select Case .Parent.cmbAssets
                Case lngTmp03
                  ' ** For accounts with 2 assets, and this is the first.
770               Select Case strAccountNo
                  Case "CRTC01"
                    ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                    ' ** croutchk_assign = 0,1, accountno = 'CRTC01'.
780                 strQryName = "qryCheckReconcile_OutChecks_09_05"
790               Case Else
                    ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                    ' ** croutchk_assign = 0,1, by specified GlobalVarGet("gstrAccountNo").
800                 strQryName = "qryCheckReconcile_OutChecks_09_02"
810               End Select
820             Case lngTmp04
                  ' ** For accounts with 2 assets, and this is the second.
830               Select Case strAccountNo
                  Case "CRTC01"
                    ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                    ' ** croutchk_assign = 0,2, accountno = 'CRTC01'.
840                 strQryName = "qryCheckReconcile_OutChecks_09_06"
850               Case Else
                    ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                    ' ** croutchk_assign = 0,2, by specified GlobalVarGet("gstrAccountNo").
860                 strQryName = "qryCheckReconcile_OutChecks_09_03"
870               End Select
880             End Select
890           End Select
900           If strQryName <> vbNullString Then
910             varTmp00 = DLookup("CalcTotal", strQryName)
920           End If

930   On Error Resume Next
940           varTmp01 = .Parent.TotalOutstandingChecks
950           If ERR.Number <> 0 Then
960   On Error GoTo ERRH
970             varTmp01 = CLng(0)
980           Else
990   On Error GoTo ERRH
1000          End If
1010          varTmp00 = ZeroIfNull(varTmp00)  ' ** Module Function: modStringFuncs.
1020          varTmp01 = ZeroIfNull(varTmp01)  ' ** Module Function: modStringFuncs.
1030          If varTmp00 <> varTmp01 Then
1040            TotOutChks .Parent, strAccountNo  ' ** Module Function: modCheckReconcile.
                '.Parent.TotalOutstandingChecks.Requery
                '.Parent.TotalOutstandingChecks2.Requery
1050            arr_varAsset_ThisAcct = .Parent.AssetsThisArray_Get  ' ** Form Function: frmCheckReconcile.
1060  On Error Resume Next
1070            lngAssets_ThisAcct = UBound(arr_varAsset_ThisAcct, 2) + 1&
1080            If ERR.Number = 0 Then
1090  On Error GoTo ERRH
1100              lngRecsCur = RecCnt  ' ** Function: Below.
1110              .Parent.OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
1120              UpdateBSTotal_CR .Parent, lngAssets_ThisAcct, arr_varAsset_ThisAcct, THIS_NAME & "." & THIS_PROC  ' ** Module Procedure: modCheckReconcile.
1130            Else
1140  On Error GoTo ERRH
1150            End If
1160          End If

              'NO PROVISION FOR MULTI-ASSETS!
              'SHOULDN'T THIS BE THE OTHER WAY AROUND? SET THE COMBO TO THIS RECORD?  DON'T TOUCH COMBO!
1170          If .croutchk_assign > 0 Then
1180            Select Case .croutchk_assign
                Case 1
1190              strCusip = .Parent.AssetNumGet(2, 1)
1200            Case 2
1210              strCusip = .Parent.AssetNumGet(2, 2)
1220            End Select
1230            DoEvents
1240          End If

1250        End If
1260      End If
1270      DoCmd.Hourglass False
1280    End With

EXITP:
1290    Exit Sub

ERRH:
1300    DoCmd.Hourglass False
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strAccountNo As String, lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant
        Dim intRetVal As Integer

1410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1420    intShiftDown = (Shift And acShiftMask) > 0
1430    intAltDown = (Shift And acAltMask) > 0
1440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-Keys to other forms and functionality.
        ' **   Recalc Tots:      F5 {RecalcTots_CR}
        ' **   Date Picker:      F7 {cmdCalendar on frmCheckReconcile}

        ' ** Shortcut Alt to other forms and functionality.
        ' **   BS Debits:        E {frmCheckReconcile_Sub_BSDebits on frmCheckReconcile}
        ' **   TA Credits:       I {frmCheckReconcile_Sub_TACredits on frmCheckReconcile}
        ' **   TA Debits:        O {frmCheckReconcile_Sub_TADebits on frmCheckReconcile}
        ' **   BS Credits:       R {frmCheckReconcile_Sub_BSCredits on frmCheckReconcile}

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   2nd Asset:        A {cmdAddAsset on frmCheckReconcile}
        ' **   Delete:           D {cmdDelete on frmCheckReconcile}
        ' **   Next Asset:       N {cmdAssetNext on frmCheckReconcile}
        ' **   Previous Asset:   P {cmdAssetPrevious on frmCheckReconcile}
        ' **   Clear:            R {cmdClear on frmCheckReconcile}
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyF5
1470        With Me
1480          DoCmd.Hourglass True
1490          DoEvents
1500          strAccountNo = .Parent.cmbAccounts.Column(CBX_ACT_ACTNO)
1510          arr_varAsset_ThisAcct = .Parent.AssetsThisArray_Get  ' ** Form Function: frmCheckReconcile.
1520          lngAssets_ThisAcct = UBound(arr_varAsset_ThisAcct, 2) + 1&
1530          RecalcTots_CR strAccountNo, lngAssets_ThisAcct, arr_varAsset_ThisAcct, .Parent  ' ** Module Procedure: modCheckReconcile.
1540          DoCmd.Hourglass False
1550        End With
1560      Case vbKeyF7
1570        With Me
1580          intRetVal = 0
1590          If .Parent.cmdCalendar.Enabled = True Then
1600            DoCmd.SelectObject acForm, .Parent.Name, False
1610            .Parent.cmdCalendar.SetFocus
1620            .Parent.cmdCalendar_Click  ' ** Form Procedure: frmCheckReconcile.
1630          Else
1640            Beep
1650          End If
1660        End With
1670      Case vbKeyUp
1680        With Me
1690          intRetVal = 0
1700          If .CurrentRecord > 1 Then
1710            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1720          End If
1730        End With
1740      Case vbKeyDown
1750        With Me
1760          intRetVal = 0
1770          lngRecsCur = RecCnt  ' ** Function: Below.
1780          If .CurrentRecord < lngRecsCur Then
1790            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1800          End If
1810        End With
1820      End Select
1830    End If

        ' ** Alt keys.
1840    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1850      Select Case intRetVal
          Case vbKeyE
1860        With Me
1870          intRetVal = 0
1880          If .Parent.frmCheckReconcile_Sub_BSDebits.Enabled = True Then
1890            DoCmd.SelectObject acForm, .Parent.Name, False
1900            .Parent.frmCheckReconcile_Sub_BSDebits.SetFocus
1910          Else
1920            Beep
1930          End If
1940        End With
1950      Case vbKeyI
1960        With Me
1970          intRetVal = 0
1980          If .Parent.frmCheckReconcile_Sub_TACredits.Enabled = True Then
1990            DoCmd.SelectObject acForm, .Parent.Name, False
2000            .Parent.frmCheckReconcile_Sub_TACredits.SetFocus
2010          Else
2020            Beep
2030          End If
2040        End With
2050      Case vbKeyO
2060        With Me
2070          intRetVal = 0
2080          If .Parent.frmCheckReconcile_Sub_TADebits.Enabled = True Then
2090            DoCmd.SelectObject acForm, .Parent.Name, False
2100            .Parent.frmCheckReconcile_Sub_TADebits.SetFocus
2110          Else
2120            Beep
2130          End If
2140        End With
2150      Case vbKeyR
2160        With Me
2170          intRetVal = 0
2180          If .Parent.frmCheckReconcile_Sub_BSCredits.Enabled = True Then
2190            DoCmd.SelectObject acForm, .Parent.Name, False
2200            .Parent.frmCheckReconcile_Sub_BSCredits.SetFocus
2210          Else
2220            Beep
2230          End If
2240        End With
2250      Case vbKeyX
2260        With Me
2270          intRetVal = 0
2280          DoCmd.SelectObject acForm, .Parent.Name, False
2290          .Parent.cmdClose.SetFocus
2300          .Parent.cmdClose_Click  ' ** Form Procedure: frmCheckReconcile.
2310        End With
2320      End Select
2330    End If

        ' ** Ctrl keys.
2340    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2350      Select Case intRetVal
          Case vbKeyA
2360        With Me
2370          intRetVal = 0
2380          Select Case .Parent.cmdAddAsset.Enabled
              Case True
2390            DoCmd.SelectObject acForm, .Parent.Name, False
2400            .Parent.cmdAddAsset.SetFocus
2410            .Parent.cmdAddAsset_Click  ' ** Form Procedure: frmCheckReconcile.
2420          Case False
2430            Beep
2440          End Select
2450        End With
2460      Case vbKeyD
2470        With Me
2480          intRetVal = 0
2490          Select Case .Parent.cmdDelete.Enabled
              Case True
2500            DoCmd.SelectObject acForm, .Parent.Name, False
2510            .Parent.cmdDelete.SetFocus
2520            .Parent.cmdDelete_Click  ' ** Form Procedure: frmCheckReconcile.
2530          Case False
2540            Beep
2550          End Select
2560        End With
2570      Case vbKeyN
2580        With Me
2590          intRetVal = 0
2600          Select Case .Parent.cmdAssetNext.Enabled
              Case True
2610            DoCmd.SelectObject acForm, .Parent.Name, False
2620            .Parent.cmdAssetNext.SetFocus
2630            .Parent.cmdAssetNext_Click  ' ** Form Procedure: frmCheckReconcile.
2640          Case False
2650            Beep
2660          End Select
2670        End With
2680      Case vbKeyP
2690        With Me
2700          intRetVal = 0
2710          Select Case .Parent.cmdAssetPrevious.Enabled
              Case True
2720            DoCmd.SelectObject acForm, .Parent.Name, False
2730            .Parent.cmdAssetPrevious.SetFocus
2740            .Parent.cmdAssetPrevious_Click  ' ** Form Procedure: frmCheckReconcile.
2750          Case False
2760            Beep
2770          End Select
2780        End With
2790      Case vbKeyR
2800        With Me
2810          intRetVal = 0
2820          Select Case .Parent.cmdClear.Enabled
              Case True
2830            DoCmd.SelectObject acForm, .Parent.Name, False
2840            .Parent.cmdClear.SetFocus
2850            .Parent.cmdClear_Click  ' ** Form Procedure: frmCheckReconcile.
2860          Case False
2870            Beep
2880          End Select
2890        End With
2900      Case vbKeyS
2910        intRetVal = 0
2920        cmdSave_Click THIS_PROC  ' ** Procedure: Below.
2930      End Select
2940    End If

        ' ** Ctrl-Shift keys.
2950    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2960      Select Case intRetVal
          Case vbKeyF
2970        With Me
2980          intRetVal = 0
2990          DoCmd.SelectObject acForm, .Parent.Name, False
3000          .Parent.FocusHolder.SetFocus
3010        End With
3020      End Select
3030    End If

EXITP:
3040    KeyCode = intRetVal
3050    Exit Sub

ERRH:
3060    intRetVal = 0
3070    Select Case ERR.Number
        Case Else
3080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3090    End Select
3100    Resume EXITP

End Sub

Private Sub Form_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_AfterUpdate"

        Dim lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant

3210    With Me
3220      TotOutChks .Parent, gstrAccountNo  ' ** Module Function: modCheckReconcile.
          '.Parent.TotalOutstandingChecks.Requery
          '.Parent.TotalOutstandingChecks2.Requery
3230      arr_varAsset_ThisAcct = .Parent.AssetsThisArray_Get  ' ** Form Function: frmCheckReconcile.
3240      lngAssets_ThisAcct = UBound(arr_varAsset_ThisAcct, 2) + 1&
3250      lngRecsCur = RecCnt  ' ** Procedure: Below.
3260      .Parent.OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
3270      UpdateBSTotal_CR .Parent, lngAssets_ThisAcct, arr_varAsset_ThisAcct, THIS_NAME & "." & THIS_PROC  ' ** Module Procedure: modCheckReconcile.
3280      .Parent.Recalc
3290    End With

EXITP:
3300    Exit Sub

ERRH:
3310    Select Case ERR.Number
        Case Else
3320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3330    End Select
3340    Resume EXITP

End Sub

Public Sub cmdSave_Click(Optional varProc As Variant)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

        Dim lngJNo As Long

3410    With Me
3420      lngJNo = .journalno
3430  On Error Resume Next
3440      DoCmd.RunCommand acCmdSaveRecord
3450      If ERR.Number <> 0 Then
3460        Select Case ERR.Number
            Case 2113  ' ** The value you entered isn't valid for this field.
3470  On Error GoTo ERRH
3480          Beep
3490          MsgBox "Check Number must be a numeric value.", vbInformation + vbOKOnly, "Invalid Entry"
3500        End Select
3510      Else
3520  On Error GoTo ERRH
3530      End If
3540      MoveRec 0, lngJNo  ' ** Procedure: Below.
3550    End With

EXITP:
3560    Exit Sub

ERRH:
3570    Select Case ERR.Number
        Case Else
3580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3590    End Select
3600    Resume EXITP

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
3790          If .opgAssign.Visible = True And .opgAssign.Enabled = True Then
3800            .opgAssign.SetFocus
3810          Else
3820            .transdate.SetFocus
3830          End If
3840        End With
3850      End Select
3860    End If

        ' ** Shift keys.
3870    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3880      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3890        With Me
3900          intRetVal = 0
3910          If .CurrentRecord > 1 Then
3920            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3930            .CheckPaid.SetFocus
3940          Else
3950            DoCmd.SelectObject acForm, .Parent.Name, False
3960            lngRecsCur = .Parent.frmCheckReconcile_Sub_TADebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TADebits.
3970            .Parent.frmCheckReconcile_Sub_TADebits.SetFocus
3980            If lngRecsCur > 0& Then
3990              .Parent.frmCheckReconcile_Sub_TADebits.Form.FocusHolder.SetFocus
4000            Else
4010              .Parent.frmCheckReconcile_Sub_TADebits.Form.critem_description.SetFocus
4020            End If
4030          End If
4040        End With
4050      End Select
4060    End If

EXITP:
4070    KeyCode = intRetVal
4080    Exit Sub

ERRH:
4090    intRetVal = 0
4100    Select Case ERR.Number
        Case Else
4110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4120    End Select
4130    Resume EXITP

End Sub

Private Sub opgAssign_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssign_lbl_DblClick"

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

Private Sub opgAssign_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssign_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String

        'THERE'S CONFUSION HERE!
        'MULTI ASSETS SHOW UP IN THE PARENT BY
'2 DIFFERENT RECORDS, ONE FOR EACH ASSET!
        'YET WE'RE ASSIGNING CHECKS FOR BOTH DOWN HERE!
        'IF SO, WE CAN'T JUST TAKE THE ASSET COMBO BOX,
        'WE HAVE TO KNOW WHICH IS 1 AND WHICH IS 2,
        'REGARDLESS OF THE COMBO BOX!
        'BOTH ASSETS WILL HAVE TO SHOW THE SAME CHECK LIST!
4310    With Me
4320      If .opgAssign > 0 Then

4330        Select Case .opgAssign  'THIS IS TAKING THE ASSET SHOWING, BUT I DON'T THINK THAT CHANGES WITH MULTI-ASSETS!
            Case .opgAssign_optOne.OptionValue
4340          .croutchk_asset1 = .Parent.AssetNumGet(1, 1)  ' ** Form Function: frmCheckReconcile.
4350          .croutchk_asset2 = 0&
4360          strCusip = .Parent.AssetNumGet(2, 1)  ' ** Form Function: frmCheckReconcile.
4370        Case .opgAssign_optTwo.OptionValue
4380          .croutchk_asset1 = 0&
4390          .croutchk_asset2 = .Parent.AssetNumGet(1, 2)  ' ** Form Function: frmCheckReconcile.
4400          strCusip = .Parent.AssetNumGet(2, 2)  ' ** Form Function: frmCheckReconcile.
4410        End Select

4420        If IsNull(.description) = True Then
4430          .description = ASSIGN_MSG & " " & strCusip
4440        Else
4450          intPos01 = InStr(.description, ASSIGN_MSG)
4460          If intPos01 > 0 Then
4470            If intPos01 = 1 Then
4480              intPos01 = InStr(.description, ";")
4490              If intPos01 > 0 Then
4500                strTmp01 = Mid(.description, (intPos01 + 1))
4510                strTmp01 = Rem_Semi(strTmp01)  ' ** Module Function: modStringFuncs.
4520                strTmp01 = strTmp01 & ";" & ASSIGN_MSG & " " & strCusip
4530                .description = strTmp01
4540              Else
4550                .description = ASSIGN_MSG & " " & strCusip
4560              End If
4570            Else
4580              .description = Left(.description, (intPos01 - 1)) & ";" & ASSIGN_MSG & " " & strCusip
4590            End If
4600          Else
4610            .description = .description & ";" & ASSIGN_MSG & " " & strCusip
4620          End If
4630        End If
4640        .Parent.UpdateCheckArray .journalno, .croutchk_id, .CheckPaid, .CheckNum, .description  ' ** Form Procedure: frmCheckReconcile.
4650        cmdSave_Click THIS_PROC  ' ** Procedure: Above.
4660        If .Parent.cmdUpdate.Enabled = False Then .Parent.cmdUpdate.Enabled = True
4670        DoEvents
4680        If .transdate.Enabled = True Then
4690          .transdate.SetFocus
4700        Else
4710          .CheckNum.SetFocus
4720        End If

4730      End If
4740    End With

EXITP:
4750    Exit Sub

ERRH:
4760    Select Case ERR.Number
        Case Else
4770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4780    End Select
4790    Resume EXITP

End Sub

Private Sub opgAssign_optOne_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssign_optOne_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4820    intShiftDown = (Shift And acShiftMask) > 0
4830    intAltDown = (Shift And acAltMask) > 0
4840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4870        With Me
4880          intRetVal = 0
4890          .transdate.SetFocus
4900        End With
4910      End Select
4920    End If

        ' ** Shift keys.
4930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4950        With Me
4960          intRetVal = 0
4970          If .CurrentRecord > 1 Then
4980            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4990            .CheckPaid.SetFocus
5000          Else
5010            DoCmd.SelectObject acForm, .Parent.Name, False
5020            lngRecsCur = .Parent.frmCheckReconcile_Sub_TADebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TADebits.
5030            If lngRecsCur = 0 Then
5040              .Parent.frmCheckReconcile_Sub_TADebits.Form.critem_description.SetFocus
5050            Else
5060              .Parent.frmCheckReconcile_Sub_TADebits.Form.FocusHolder.SetFocus
5070            End If
5080          End If
5090        End With
5100      End Select
5110    End If

EXITP:
5120    KeyCode = intRetVal
5130    Exit Sub

ERRH:
5140    intRetVal = 0
5150    Select Case ERR.Number
        Case Else
5160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5170    End Select
5180    Resume EXITP

End Sub

Private Sub opgAssign_optTwo_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAssign_optTwo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5270        With Me
5280          intRetVal = 0
5290          .transdate.SetFocus
5300        End With
5310      End Select
5320    End If

        ' ** Shift keys.
5330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5350        With Me
5360          intRetVal = 0
5370          If .CurrentRecord > 1 Then
5380            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5390            .CheckPaid.SetFocus
5400          Else
5410            DoCmd.SelectObject acForm, .Parent.Name, False
5420            lngRecsCur = .Parent.frmCheckReconcile_Sub_TADebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TADebits.
5430            If lngRecsCur = 0 Then
5440              .Parent.frmCheckReconcile_Sub_TADebits.Form.critem_description.SetFocus
5450            Else
5460              .Parent.frmCheckReconcile_Sub_TADebits.Form.FocusHolder.SetFocus
5470            End If
5480          End If
5490        End With
5500      End Select
5510    End If

EXITP:
5520    KeyCode = intRetVal
5530    Exit Sub

ERRH:
5540    intRetVal = 0
5550    Select Case ERR.Number
        Case Else
5560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5570    End Select
5580    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

5610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub transdate_GotFocus()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_GotFocus"

5710    With Me
5720      .transdate.SelLength = 0
5730      .transdate.SelStart = 99
5740    End With

EXITP:
5750    Exit Sub

ERRH:
5760    Select Case ERR.Number
        Case Else
5770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5780    End Select
5790    Resume EXITP

End Sub

Private Sub CheckNum_lbl_DblClick(Cancel As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_lbl_DblClick"

5810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5820    Exit Sub

ERRH:
5830    Select Case ERR.Number
        Case Else
5840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5850    End Select
5860    Resume EXITP

End Sub

Private Sub CheckNum_AfterUpdate()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckNum_AfterUpdate"

5910    With Me
5920      .Parent.UpdateCheckArray .journalno, .croutchk_id, .CheckPaid, .CheckNum, .description  ' ** Form Procedure: frmCheckReconcile.
5930      cmdSave_Click THIS_PROC  ' ** Procedure: Above.
5940      If .Parent.cmdUpdate.Enabled = False Then .Parent.cmdUpdate.Enabled = True
5950      .croutchk_payee.SetFocus
5960    End With

EXITP:
5970    Exit Sub

ERRH:
5980    Select Case ERR.Number
        Case Else
5990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6000    End Select
6010    Resume EXITP

End Sub

Private Sub croutchk_payee_lbl_DblClick(Cancel As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "croutchk_payee_lbl_DblClick"

6110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6120    Exit Sub

ERRH:
6130    Select Case ERR.Number
        Case Else
6140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6150    End Select
6160    Resume EXITP

End Sub

Private Sub croutchk_payee_GotFocus()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "croutchk_payee_GotFocus"

6210    With Me
6220      .croutchk_payee.SelLength = 0
6230      .croutchk_payee.SelStart = 99
6240    End With

EXITP:
6250    Exit Sub

ERRH:
6260    Select Case ERR.Number
        Case Else
6270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6280    End Select
6290    Resume EXITP

End Sub

Private Sub croutchk_amount_lbl_DblClick(Cancel As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "croutchk_amount_lbl_DblClick"

6310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub croutchk_amount_GotFocus()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "croutchk_amount_GotFocus"

6410    With Me
6420      .croutchk_amount.SelLength = 0
6430      .croutchk_amount.SelStart = 0
6440    End With

EXITP:
6450    Exit Sub

ERRH:
6460    Select Case ERR.Number
        Case Else
6470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6480    End Select
6490    Resume EXITP

End Sub

Private Sub CheckPaid_lbl_DblClick(Cancel As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckPaid_lbl_DblClick"

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

Private Sub CheckPaid_AfterUpdate()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckPaid_AfterUpdate"

6610    With Me
6620      .Parent.UpdateCheckArray .journalno, .croutchk_id, .CheckPaid, .CheckNum, .description  ' ** Form Procedure: frmCheckReconcile.
6630      If .Parent.cmdUpdate.Enabled = False Then .Parent.cmdUpdate.Enabled = True
6640      cmdSave_Click THIS_PROC  ' ** Procedure: Above.
6650      .FocusHolder2.SetFocus
6660    End With

EXITP:
6670    Exit Sub

ERRH:
6680    Select Case ERR.Number
        Case Else
6690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6700    End Select
6710    Resume EXITP

End Sub

Private Sub CheckPaid_KeyDown(KeyCode As Integer, Shift As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckPaid_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6820    intShiftDown = (Shift And acShiftMask) > 0
6830    intAltDown = (Shift And acAltMask) > 0
6840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6870        With Me
6880          intRetVal = 0
6890          lngRecsCur = RecCnt  ' ** Function: Below.
6900          If .CurrentRecord < lngRecsCur Then
6910            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
6920            If .opgAssign.Visible = True And .opgAssign.Enabled = True Then
6930  On Error Resume Next
6940              .opgAssign.SetFocus
6950  On Error GoTo ERRH
6960            Else
6970  On Error Resume Next
6980              .transdate.SetFocus
6990  On Error GoTo ERRH
7000            End If
7010          Else
7020            DoCmd.SelectObject acForm, .Parent.Name, False
7030  On Error Resume Next
7040            .Parent.chkShowPaid.SetFocus
7050  On Error GoTo ERRH
7060          End If
7070        End With
7080      End Select
7090    End If

        ' ** Shift keys.
7100    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7110      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7120        With Me
7130          intRetVal = 0
7140          .croutchk_amount.SetFocus
7150        End With
7160      End Select
7170    End If

EXITP:
7180    KeyCode = intRetVal
7190    Exit Sub

ERRH:
7200    intRetVal = 0
7210    Select Case ERR.Number
        Case Else
7220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7230    End Select
7240    Resume EXITP

End Sub

Private Sub FocusHolder2_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7320    intShiftDown = (Shift And acShiftMask) > 0
7330    intAltDown = (Shift And acAltMask) > 0
7340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7370        With Me
7380          intRetVal = 0
7390          lngRecsCur = RecCnt  ' ** Function: Below.
7400          If .CurrentRecord < lngRecsCur Then
7410            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
7420            If .opgAssign.Visible = True And .opgAssign.Enabled = True Then
7430              .opgAssign.SetFocus
7440            Else
7450              .transdate.SetFocus
7460            End If
7470          Else
7480            DoCmd.SelectObject acForm, .Parent.Name, False
7490            .Parent.chkShowPaid.SetFocus
7500          End If
7510        End With
7520      End Select
7530    End If

        ' ** Shift keys.
7540    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7550      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7560        With Me
7570          intRetVal = 0
7580          .CheckPaid.SetFocus
7590        End With
7600      End Select
7610    End If

EXITP:
7620    KeyCode = intRetVal
7630    Exit Sub

ERRH:
7640    intRetVal = 0
7650    Select Case ERR.Number
        Case Else
7660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7670    End Select
7680    Resume EXITP

End Sub

Public Function RecCnt() As Long

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

7710    lngRetVal = 0&

7720    With Me
7730      Set rst = .RecordsetClone
7740      If rst.BOF = True And rst.EOF = True Then
            ' ** Shouldn't really be here if that's the case.
7750        rst.Close
7760        Set rst = Nothing
7770        Set rst = Me.RecordsetClone
7780  On Error Resume Next
7790        rst.MoveFirst
7800        If ERR.Number = 0 Then
7810  On Error GoTo ERRH
7820          rst.MoveLast
7830          lngRetVal = rst.RecordCount
7840        Else
7850  On Error GoTo ERRH
7860        End If
7870      Else
7880        rst.MoveLast
7890        lngRetVal = rst.RecordCount
7900      End If
7910      rst.Close

7920    End With

EXITP:
7930    Set rst = Nothing
7940    RecCnt = lngRetVal
7950    Exit Function

ERRH:
7960    lngRetVal = 0&
7970    Select Case ERR.Number
        Case Else
7980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7990    End Select
8000    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

8110    Select Case intWhere
        Case acCmdRecordsGoToFirst
8120      DoCmd.RunCommand acCmdRecordsGoToFirst
8130    Case acCmdRecordsGoToLast
8140      DoCmd.RunCommand acCmdRecordsGoToLast
8150    Case acCmdRecordsGoToPrevious
8160      DoCmd.RunCommand acCmdRecordsGoToPrevious
8170    Case acCmdRecordsGoToNext
8180      DoCmd.RunCommand acCmdRecordsGoToNext
8190    Case Else
8200      If IsMissing(varID) = False Then
8210        With Me
8220          Set rst = .RecordsetClone
8230          With rst
8240            .FindFirst "[journalno] = " & CStr(varID)
8250            If .NoMatch = False Then
8260              Me.Bookmark = .Bookmark
8270            End If
8280            .Close
8290          End With
8300        End With
8310      End If
8320    End Select

EXITP:
8330    Set rst = Nothing
8340    Exit Sub

ERRH:
8350    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8360    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8370    Case Else
8380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8390    End Select
8400    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [transdate], [croutchk_id]"

8510    With Me
8520      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
8530        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
8540      End If
8550      .Controls(strSortLbl).Visible = False
8560      .Controls(strSortLine1).Visible = False
8570      .Controls(strSortLine2).Visible = False
8580      .Controls(strSortLine1).Width = lngTpp  ' ** So they don't push off the right side of the form.
8590      .Controls(strSortLine2).Width = lngTpp
8600      If strProc = "Form_Load" Then
8610        strCalled = "transdate"
8620        strSortNow = strSortOrig
8630        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
8640        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
8650        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
8660        .Controls(strSortLbl).Left = lngSortLbl_Left
8670        .Controls(strSortLine1).Left = lngSortLine_Left
8680        .Controls(strSortLine2).Left = lngSortLine_Left
8690        .Controls(strSortLine1).Width = lngSortLine_Width
8700        .Controls(strSortLine2).Width = lngSortLine_Width
8710        .Controls(strSortLbl).Caption = strArwDn
8720        .Controls(strSortLbl).ForeColor = CLR_DKRED
8730      Else
8740        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
8750        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
8760        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
8770        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
8780        .Controls(strSortLbl).Left = lngSortLbl_Left
8790        .Controls(strSortLine1).Left = lngSortLine_Left
8800        .Controls(strSortLine2).Left = lngSortLine_Left
8810        .Controls(strSortLine1).Width = lngSortLine_Width
8820        .Controls(strSortLine2).Width = lngSortLine_Width
8830        If strCalled = "transdate" Then
8840          strSortAsc = strSortOrig
8850          If strSortNow = strSortAsc Then
8860            strSortNow = "[transdate], [croutchk_id] DESC"
8870            .Controls(strSortLbl).Caption = strArwUp
8880            .Controls(strSortLbl).ForeColor = CLR_DKBLU
8890          Else
8900            strSortNow = strSortAsc
8910            .Controls(strSortLbl).Caption = strArwDn
8920            .Controls(strSortLbl).ForeColor = CLR_DKRED
8930          End If
8940        Else
8950          .Controls(strSortLbl).Caption = strArwUp
8960          .Controls(strSortLbl).ForeColor = CLR_DKBLU
8970          Select Case strCalled
              Case "CheckNum"
8980            strSortAsc = "[CheckNum]" & strStdAsc
8990          Case "croutchk_payee"
9000            strSortAsc = "[croutchk_payee]" & strStdAsc
9010          Case "croutchk_amount"
9020            strSortAsc = "[croutchk_amount]" & strStdAsc
9030          Case "CheckPaid"
9040            strSortAsc = "[CheckPaid]" & strStdAsc
9050            lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
9060            lngSortLine_Left = .CheckPaid_lbl_line.Left
9070            lngSortLine_Width = (.CheckPaid_lbl_line.Width + lngTpp)
9080            .Controls(strSortLbl).Left = lngSortLbl_Left
9090            .Controls(strSortLine1).Left = lngSortLine_Left
9100            .Controls(strSortLine2).Left = lngSortLine_Left
9110            .Controls(strSortLine1).Width = lngSortLine_Width
9120            .Controls(strSortLine2).Width = lngSortLine_Width
9130          Case "opgAssign"
9140            strSortAsc = "[croutchk_assign]" & strStdAsc
9150            lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
9160            .Controls(strSortLbl).Left = lngSortLbl_Left
9170          End Select
9180          If strSortNow = strSortAsc Then
9190            intCnt = (CharCnt(strSortAsc, ",") + 1)  ' ** Module Function: modStringFuncs.
9200            Select Case intCnt
                Case 1
9210              strSortNow = strSortAsc & " DESC"
9220            Case 2
9230              intPos01 = InStr(strSortAsc, ",")
9240              strSortNow = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
9250            Case 3
9260              intPos01 = InStr(strSortAsc, ",")
9270              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9280              strTmp02 = Mid(strSortAsc, intPos01)
9290              intPos01 = InStr(2, strTmp02, ",")
9300              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9310              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
9320              strSortNow = strTmp01
9330            End Select
9340            .Controls(strSortLbl).Caption = strArwDn
9350            .Controls(strSortLbl).ForeColor = CLR_DKRED
9360          Else
9370            strSortNow = strSortAsc
9380          End If
9390        End If
9400      End If
9410      .Controls(strSortLbl).Visible = True
9420      .Controls(strSortLine1).Visible = True
9430      .Controls(strSortLine2).Visible = True
9440      .OrderBy = strSortNow
9450      .OrderByOn = True
9460    End With

EXITP:
9470    Exit Sub

ERRH:
9480    Select Case ERR.Number
        Case Else
9490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9500    End Select
9510    Resume EXITP

End Sub

Public Function SortNow_Get() As String

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

9610    strRetVal = strSortNow

EXITP:
9620    SortNow_Get = strRetVal
9630    Exit Function

ERRH:
9640    strRetVal = strSortOrig
9650    Select Case ERR.Number
        Case Else
9660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9670    End Select
9680    Resume EXITP

End Function
