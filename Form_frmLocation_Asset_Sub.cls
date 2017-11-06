VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLocation_Asset_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLocation_Asset_Sub"

'VGC 07/24/017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts on frmLocation_Asset}
' **   Number:           B {opgAccountSource_optNumber on frmLocation_Asset}
' **   Close:            C {cmdClose on frmLocation_Asset}
' **   Description:      D {asset_description}
' **   Select Asset:     E {cmbAssets on frmLocation_Asset}
' **   Current Loc:      I {cmbLocations_Current on frmLocation_Asset}
' **   All Group:        L {opgGroup_optAll on frmLocation_Asset}
' **   Acct Name:        M {opgAccountSource_optName on frmLocation_Asset}
' **   Select None:      N {cmdSelectNone on frmLocation_Asset}
' **   Account Group:    O {opgGroup_optAccount on frmLocation_Asset}
' **   Reset:            R {cmdResetFilter on frmLocation_Asset}
' **   Select All:       S {cmdSelectAll on frmLocation_Asset}
' **   Asset Group:      T {opgGroup_optAsset on frmLocation_Asset}
' **   Update:           U {cmdUpdate on frmLocation_Asset}
' **   Move:             V {locasset_mark}
' **   New Loc:          W {cmbLocations_New on frmLocation_Asset}
' **   Exit:             X {cmdClose on frmLocation_Asset}
' **   Accnt/Asset Grp:  / {opgGroup_optAccountAsset on frmLocation_Asset}

' ** Shortcut Alt-Shift keys responsive from this form:
' **   Asset Cusip:      I {opgAssetSource_optCusip on frmLocation_Asset}
' **   Asset Name:       N {opgAssetSource_optName on frmLocation_Asset}
' **   Type:             Y {opgAssetSource_optType on frmLocation_Asset}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[alphasort], [assettype], [asset_description], [assetdate]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Width As Long

Private blnRecalcTots As Boolean
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
150       If strCallingForm = "frmLocation_Asset" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         blnRecalcTots = False
200         lngSortLbl_Width = .Sort_lbl.Width

210       Else
220         Cancel = -1
230       End If
240     End With

250     If Cancel = -1 Then
260       Beep
270       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
280     End If

EXITP:
290     Exit Sub

ERRH:
300     Select Case ERR.Number
        Case Else
310       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
320     End Select
330     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me

420       .accountno.ForeColor = CLR_DISABLED_FG
430       .accountno.BackColor = CLR_DISABLED_BG
440       .cusip.ForeColor = CLR_DISABLED_FG
450       .cusip.BackColor = CLR_DISABLED_BG
460       .asset_description.ForeColor = CLR_DISABLED_FG
470       .asset_description.BackColor = CLR_DISABLED_BG
480       .assettype.ForeColor = CLR_DISABLED_FG
490       .assettype.BackColor = CLR_DISABLED_BG
500       .assetdate.ForeColor = CLR_DISABLED_FG
510       .assetdate.BackColor = CLR_DISABLED_BG
520       .transdate.ForeColor = CLR_DISABLED_FG
530       .transdate.BackColor = CLR_DISABLED_BG
540       .shareface.ForeColor = CLR_DISABLED_FG
550       .shareface.BackColor = CLR_DISABLED_BG

560       SortNow THIS_PROC  ' ** Procedure: Below.

570     End With

EXITP:
580     Exit Sub

ERRH:
590     Select Case ERR.Number
        Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strControlName As String
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

750   On Error Resume Next
760     strControlName = Screen.ActiveControl.Name
770   On Error GoTo ERRH

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Number:           B {opgAccountSource_optNumber on frmLocation_Asset}
        ' **   Description:      D {asset_description}
        ' **   Acct Name:        M {opgAccountSource_optName on frmLocation_Asset}
        ' **   Move:             V {locasset_mark}
        ' **   Type:             Y {opgAssetSource_optType on frmLocation_Asset}

        ' ** Shortcut Alt-Shift keys to other forms and functionality:
        ' **   Asset Cusip:      I {opgAssetSource_optCusip on frmLocation_Asset}
        ' **   Asset Name:       N {opgAssetSource_optName on frmLocation_Asset}
        ' **   Type:             Y {opgAssetSource_optType on frmLocation_Asset}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
780     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
790       Select Case intRetVal
          Case vbKeyEscape
800         With Me
810           intRetVal = 0
820           .Parent.cmdClose_Click  ' ** Form Procedure: frmLocation_Asset.
830         End With
840       Case vbKeyUp
850         With Me
860           intRetVal = 0
870           If .CurrentRecord > 1 Then
880             MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
890             Select Case strControlName
                Case "FocusHolder"
900               .FocusHolder.SetFocus
910             Case "locasset_mark"
920               .locasset_mark.SetFocus
930             Case Else
940               .FocusHolder.SetFocus
950             End Select
960           End If
970         End With
980       Case vbKeyDown
990         With Me
1000          intRetVal = 0
1010          lngRecsCur = RecCnt  ' ** Function: Below.
1020          If .CurrentRecord < lngRecsCur Then
1030            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1040            Select Case strControlName
                Case "FocusHolder"
1050              .FocusHolder.SetFocus
1060            Case "locasset_mark"
1070              .locasset_mark.SetFocus
1080            Case Else
1090              .FocusHolder.SetFocus
1100            End Select
1110          End If
1120        End With
1130      End Select
1140    End If

        ' ** Alt keys.
1150    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyB  ' ** Because the '_lbl' is empty.
1170        With Me
1180          intRetVal = 0
1190          If .Parent.opgAccountSource.Enabled = True Then
1200            DoCmd.SelectObject acForm, .Parent.Name, False
1210            .Parent.opgAccountSource.SetFocus
1220            .Parent.opgAccountSource = .Parent.opgAccountSource_optNumber.OptionValue
1230            .Parent.opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1240          Else
1250            Beep
1260          End If
1270        End With
1280      Case vbKeyD
1290        With Me
1300          intRetVal = 0
1310          .FocusHolder.SetFocus
1320        End With
1330      Case vbKeyM  ' ** Because the '_lbl' is empty.
1340        With Me
1350          intRetVal = 0
1360          If .Parent.opgAccountSource.Enabled = True Then
1370            DoCmd.SelectObject acForm, .Parent.Name, False
1380            .Parent.opgAccountSource.SetFocus
1390            .Parent.opgAccountSource = .Parent.opgAccountSource_optName.OptionValue
1400            .Parent.opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1410          Else
1420            Beep
1430          End If
1440        End With
1450      Case vbKeyV
1460        With Me
1470          intRetVal = 0
1480          .locasset_mark.SetFocus
1490        End With
1500      Case vbKeyX
1510        With Me
1520          intRetVal = 0
1530          .Parent.cmdClose_Click  ' ** Form Procedure: frmLocation_Asset.
1540        End With
1550      Case vbKeyY  ' ** Because the '_lbl' is empty.
1560        With Me
1570          intRetVal = 0
1580          If .Parent.opgAssetSource.Enabled = True Then
1590            DoCmd.SelectObject acForm, .Parent.Name, False
1600            .Parent.opgAssetSource.SetFocus
1610            .Parent.opgAssetSource = .Parent.opgAssetSource_optType.OptionValue
1620            .Parent.opgAssetSource_AfterUpdate  ' ** Procedure: Below.
1630          Else
1640            Beep
1650          End If
1660        End With
1670      End Select
1680    End If

        ' ** Alt-Shift keys.
1690    If (Not intCtrlDown) And intAltDown And intShiftDown Then
1700      Select Case intRetVal
          Case vbKeyI
1710        With Me
1720          intRetVal = 0
1730          If .Parent.opgAssetSource.Enabled = True Then
1740            DoCmd.SelectObject acForm, .Parent.Name, False
1750            .Parent.opgAssetSource.SetFocus
1760            .Parent.opgAssetSource = .Parent.opgAssetSource_optCusip.OptionValue
1770            .Parent.opgAssetSource_AfterUpdate  ' ** Procedure: Below.
1780          Else
1790            Beep
1800          End If
1810        End With
1820      Case vbKeyN
1830        With Me
1840          intRetVal = 0
1850          If .Parent.opgAssetSource.Enabled = True Then
1860            DoCmd.SelectObject acForm, .Parent.Name, False
1870            .Parent.opgAssetSource.SetFocus
1880            .Parent.opgAssetSource = .Parent.opgAssetSource_optName.OptionValue
1890            .Parent.opgAssetSource_AfterUpdate  ' ** Form Procedure: frmLocation_Asset.
1900          Else
1910            Beep
1920          End If
1930        End With
1940      Case vbKeyY
1950        With Me
1960          intRetVal = 0
1970          If .Parent.opgAssetSource.Enabled = True Then
1980            DoCmd.SelectObject acForm, .Parent.Name, False
1990            .Parent.opgAssetSource.SetFocus
2000            .Parent.opgAssetSource = .Parent.opgAssetSource_optType.OptionValue
2010            .Parent.opgAssetSource_AfterUpdate  ' ** Form Procedure: frmLocation_Asset.
2020          Else
2030            Beep
2040          End If
2050        End With
2060      End Select
2070    End If

        ' ** Ctrl keys.
2080    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2090      Select Case intRetVal
          Case vbKeyS
2100        intRetVal = 0
2110        cmdSave_Click  ' ** Procedure: Below.
2120      Case vbKeyTab, vbKeyReturn
2130        With Me
2140          intRetVal = 0
2150          DoCmd.SelectObject acForm, .Parent.Name, False
2160          If .Parent.cmdSelectAll.Enabled = True Then
2170            .Parent.cmdSelectAll.SetFocus
2180          ElseIf .Parent.cmdUpdate.Enabled = True Then
2190            .Parent.cmdUpdate.SetFocus
2200          Else
2210            .Parent.cmdClose.SetFocus
2220          End If
2230        End With
2240      Case vbKeyUp, vbKeyPageUp, vbKeyHome
2250        intRetVal = 0
2260        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2270      Case vbKeyDown, vbKeyPageDown, vbKeyEnd
2280        intRetVal = 0
2290        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2300      End Select
2310    End If

        ' ** Ctrl-Shift keys.
2320    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2330      Select Case intRetVal
          Case vbKeyTab
2340        With Me
2350          intRetVal = 0
2360          DoCmd.SelectObject acForm, .Parent.Name, False
2370          If .Parent.cmdResetFilter.Enabled = True Then
2380            .Parent.cmdResetFilter.SetFocus
2390          ElseIf .Parent.cmbAssets.Enabled = True Then
2400            .Parent.cmbAssets.SetFocus
2410          ElseIf .Parent.cmbAccounts.Enabled = True Then
2420            .Parent.cmbAccounts.SetFocus
2430          Else
2440            .Parent.opgGroup.SetFocus
2450          End If
2460        End With
2470      End Select
2480    End If

EXITP:
2490    KeyCode = intRetVal
2500    Exit Sub

ERRH:
2510    intRetVal = 0
2520    Select Case ERR.Number
        Case Else
2530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2540    End Select
2550    Resume EXITP

End Sub

Private Sub Form_Timer()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngTmp01 As Long

2610    With Me
2620      .TimerInterval = 0&
2630      If blnRecalcTots = True Then
2640        blnRecalcTots = False
2650  On Error Resume Next
2660        lngTmp01 = Abs(Nz(.locasset_mark_sum, 0))
2670        If ERR.Number <> 0 Then
2680  On Error GoTo ERRH
2690          lngTmp01 = 0&
2700        Else
2710  On Error GoTo ERRH
2720        End If
2730        .Parent.locasset_mark_sum = lngTmp01
2740        If lngTmp01 > 0& Then
2750          .Parent.locasset_mark_sum.ForeColor = CLR_GRN
2760        Else
2770          .Parent.locasset_mark_sum.ForeColor = CLR_BLU
2780        End If
2790        lngRecsCur = RecCnt  ' ** Function: Below
2800        If lngRecsCur > 0& Then
2810          ColHeaders True  ' ** Procedure: Below
2820        Else
2830          ColHeaders False  ' ** Procedure: Below
2840        End If
2850      End If
2860    End With

EXITP:
2870    Exit Sub

ERRH:
2880    Select Case ERR.Number
        Case Else
2890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2900    End Select
2910    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

3010    With Me
3020      .locasset_datemodified = Now()
3030    End With

EXITP:
3040    Exit Sub

ERRH:
3050    Select Case ERR.Number
        Case Else
3060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3070    End Select
3080    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3110    DoCmd.RunCommand acCmdSaveRecord

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

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
3290          lngRecsCur = RecCnt  ' ** Function: Below.
3300          If .CurrentRecord < lngRecsCur Then
3310            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3320          Else
3330            DoCmd.SelectObject acForm, .Parent.Name, False
3340            If .Parent.cmdSelectAll.Enabled = True Then
3350              .Parent.cmdSelectAll.SetFocus
3360            ElseIf .Parent.cmdUpdate.Enabled = True Then
3370              .Parent.cmdUpdate.SetFocus
3380            Else
3390              .Parent.cmdClose.SetFocus
3400            End If
3410          End If
3420        End With
3430      End Select
3440    End If

        ' ** Shift keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3470        With Me
3480          intRetVal = 0
3490          If .CurrentRecord > 1 Then
3500            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3510            .FocusHolder.SetFocus
3520          Else
3530            DoCmd.SelectObject acForm, .Parent.Name, False
3540            If .Parent.cmdResetFilter.Enabled = True Then
3550              .Parent.cmdResetFilter.SetFocus
3560            ElseIf .Parent.cmbAssets.Enabled = True Then
3570              .Parent.cmbAssets.SetFocus
3580            ElseIf .Parent.cmbAccounts.Enabled = True Then
3590              .Parent.cmbAccounts.SetFocus
3600            Else
3610              .Parent.opgGroup.SetFocus
3620            End If
3630          End If
3640        End With
3650      End Select
3660    End If

EXITP:
3670    KeyCode = intRetVal
3680    Exit Sub

ERRH:
3690    intRetVal = 0
3700    Select Case ERR.Number
        Case Else
3710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3720    End Select
3730    Resume EXITP

End Sub

Private Sub locasset_mark_lbl_DblClick(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "locasset_mark_lbl_DblClick"

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

Private Sub locasset_mark_AfterUpdate()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "locasset_mark_AfterUpdate"

3910    With Me
3920      Select Case .locasset_mark
          Case True
3930        .Location_ID2 = .Parent.cmbLocations_New.Column(0)
3940        .Location_Name2 = .Parent.cmbLocations_New.Column(1)
3950      Case False
3960        .Location_ID2 = Null
3970        .Location_Name2 = Null
3980      End Select
3990      cmdSave_Click  ' ** Procedure: Above.
4000      blnRecalcTots = True
4010      .TimerInterval = 500&
4020    End With

EXITP:
4030    Exit Sub

ERRH:
4040    Select Case ERR.Number
        Case Else
4050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4060    End Select
4070    Resume EXITP

End Sub

Private Sub locasset_mark_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "locasset_mark_KeyDown"

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
4190          lngRecsCur = RecCnt  ' ** Function: Below.
4200          If .CurrentRecord < lngRecsCur Then
4210            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4220            .locasset_mark.SetFocus
4230          Else
4240            DoCmd.SelectObject acForm, .Parent.Name, False
4250            If .Parent.cmdSelectAll.Enabled = True Then
4260              .Parent.cmdSelectAll.SetFocus
4270            ElseIf .Parent.cmdUpdate.Enabled = True Then
4280              .Parent.cmdUpdate.SetFocus
4290            Else
4300              .Parent.cmdClose.SetFocus
4310            End If
4320          End If
4330        End With
4340      End Select
4350    End If

        ' ** Shift keys.
4360    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4370      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4380        With Me
4390          intRetVal = 0
4400          If .CurrentRecord > 1 Then
4410            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4420            .locasset_mark.SetFocus
4430          Else
4440            DoCmd.SelectObject acForm, .Parent.Name, False
4450            If .Parent.cmdResetFilter.Enabled = True Then
4460              .Parent.cmdResetFilter.SetFocus
4470            ElseIf .Parent.cmbAssets.Enabled = True Then
4480              .Parent.cmbAssets.SetFocus
4490            ElseIf .Parent.cmbAccounts.Enabled = True Then
4500              .Parent.cmbAccounts.SetFocus
4510            Else
4520              .Parent.opgGroup.SetFocus
4530            End If
4540          End If
4550        End With
4560      End Select
4570    End If

EXITP:
4580    KeyCode = intRetVal
4590    Exit Sub

ERRH:
4600    intRetVal = 0
4610    Select Case ERR.Number
        Case Else
4620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4630    End Select
4640    Resume EXITP

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

Private Sub cusip_lbl_DblClick(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_lbl_DblClick"

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

Private Sub asset_description_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_lbl_DblClick"

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

Private Sub assettype_lbl_DblClick(Cancel As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_lbl_DblClick"

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

Private Sub assetdate_lbl_DblClick(Cancel As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_lbl_DblClick"

5110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub transdate_lbl_DblClick(Cancel As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "transdate_lbl_DblClick"

5210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5220    Exit Sub

ERRH:
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

5310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Public Function RecCnt() As Long

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

5410    lngRetVal = 0&

5420    With Me
5430      Set rst = .RecordsetClone
5440      With rst
5450        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
5460        Else
5470          .MoveLast
5480          lngRetVal = .RecordCount
5490        End If
5500        .Close
5510      End With
5520    End With

EXITP:
5530    Set rst = Nothing
5540    RecCnt = lngRetVal
5550    Exit Function

ERRH:
5560    lngRetVal = 0&
5570    Select Case ERR.Number
        Case Else
5580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5590    End Select
5600    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

5710    Select Case intWhere
        Case acCmdRecordsGoToFirst
5720      DoCmd.RunCommand acCmdRecordsGoToFirst
5730    Case acCmdRecordsGoToLast
5740      DoCmd.RunCommand acCmdRecordsGoToLast
5750    Case acCmdRecordsGoToPrevious
5760      DoCmd.RunCommand acCmdRecordsGoToPrevious
5770    Case acCmdRecordsGoToNext
5780      DoCmd.RunCommand acCmdRecordsGoToNext
5790    Case Else
5800      If IsMissing(varID) = False Then
5810        With Me
5820          Set rst = .RecordsetClone
5830          With rst
5840            .FindFirst "[accountno] = '" & CStr(varID) & "'"
5850            If .NoMatch = False Then
5860              Me.Bookmark = .Bookmark
5870            End If
5880            .Close
5890          End With
5900        End With
5910      End If
5920    End Select

EXITP:
5930    Set rst = Nothing
5940    Exit Sub

ERRH:
5950    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5960    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
5970    Case Else
5980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5990    End Select
6000    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim ctl As Access.Control
        Dim strCalled As String, strSortAsc As String
        Dim intFlds As Integer
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [alphasort], [assettype], [asset_description], [assetdate]"

6110    With Me
6120      lngRecsCur = RecCnt  ' ** Function: Above.
6130      If lngRecsCur > 0& Then
6140        .Controls(strSortLbl).Visible = False
6150        .Controls(strSortLine).Visible = False
6160        .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the the right side of the form.
6170        If strProc = "Form_Load" Then
6180          strCalled = "accountno"
6190          strSortNow = strSortOrig
6200          lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6210          lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
6220          lngSortLine_Top = lngSortLbl_Top ' - lngTpp
6230          lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6240          .Controls(strSortLbl).Top = lngSortLbl_Top
6250          .Controls(strSortLbl).Left = lngSortLbl_Left
6260          .Controls(strSortLbl).Caption = strArwUp
6270          .Controls(strSortLbl).ForeColor = CLR_DKBLU
6280          .Controls(strSortLine).Top = lngSortLine_Top
6290          .Controls(strSortLine).Left = .Controls(strCalled & "_lbl").Left
6300          .Controls(strSortLine).Width = lngSortLine_Width
6310        Else
6320          strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
6330          lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
6340          lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
6350          lngSortLine_Top = lngSortLbl_Top '- lngTpp
6360          lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6370          .Controls(strSortLbl).Top = lngSortLbl_Top
6380          .Controls(strSortLbl).Left = lngSortLbl_Left
6390          .Controls(strSortLine).Top = lngSortLine_Top
6400          .Controls(strSortLine).Left = .Controls(strCalled & "_lbl").Left
6410          .Controls(strSortLine).Width = lngSortLine_Width
6420          If strCalled = "accountno" Then
6430            strSortAsc = strSortOrig
6440            If strSortNow = strSortAsc Then
6450              strSortNow = "[alphasort] DESC, [assettype] DESC, [asset_description] DESC, [assetdate] DESC"
6460              .Controls(strSortLbl).Caption = strArwDn
6470              .Controls(strSortLbl).ForeColor = CLR_DKRED
6480            Else
6490              strSortNow = strSortAsc
6500              .Controls(strSortLbl).Caption = strArwUp
6510              .Controls(strSortLbl).ForeColor = CLR_DKBLU
6520            End If
6530          Else
6540            .Controls(strSortLbl).Caption = strArwUp
6550            .Controls(strSortLbl).ForeColor = CLR_DKBLU
6560            Select Case strCalled
                Case "locasset_mark"
6570              strSortAsc = "[locasset_mark]" & strStdAsc
6580              lngSortLbl_Left = (lngSortLbl_Left + (3& * lngTpp))
6590              .Controls(strSortLbl).Left = lngSortLbl_Left
6600            Case "cusip"
6610              strSortAsc = "[cusip]" & strStdAsc
6620            Case "asset_description"
6630              strSortAsc = "[asset_description], [cusip]" & strStdAsc
6640            Case "assettype"
6650              strSortAsc = "[assettype]" & strStdAsc
6660              lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
6670              .Controls(strSortLbl).Left = lngSortLbl_Left
6680            Case "assetdate"
6690              strSortAsc = "assetdate" & strStdAsc
6700            Case "transdate"
6710              strSortAsc = "transdate" & strStdAsc
6720            Case "shareface"
6730              strSortAsc = "shareface" & strStdAsc
6740            End Select
6750            If strSortNow = strSortAsc Then
6760              intPos01 = InStr(strSortAsc, ",")
6770              If intPos01 = 0 Then
                    ' ** [accounttype_description] DESC.
6780                strSortNow = strSortNow & " DESC"
6790              Else
                    ' ** [accounttypegroup_name] DESC, [accounttype] DESC.
6800                intFlds = CharCnt(strSortAsc, ",")
6810                intFlds = intFlds + 1
6820                Select Case intFlds  ' ** 6 Max.
                    Case 2
6830                  strSortNow = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
6840                Case 3
6850                  strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC,"
6860                  strTmp02 = Mid(strSortAsc, (intPos01 + 1))
6870                  intPos02 = InStr(strTmp02, ",")
6880                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC"
6890                  strTmp01 = strTmp01 & Mid(strTmp02, intPos02) & " DESC"
6900                  strSortNow = strTmp01
6910                Case 4
6920                  strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC,"
6930                  strTmp02 = Mid(strSortAsc, (intPos01 + 1))
6940                  intPos02 = InStr(strTmp02, ",")
6950                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC,"
6960                  strTmp02 = Mid(strTmp02, (intPos02 + 1))
6970                  intPos02 = InStr(strTmp02, ",")
6980                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC"
6990                  strTmp01 = strTmp01 & Mid(strTmp02, intPos02) & " DESC"
7000                  strSortNow = strTmp01
7010                Case 5
7020                  strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC,"
7030                  strTmp02 = Mid(strSortAsc, (intPos01 + 1))
7040                  intPos02 = InStr(strTmp02, ",")
7050                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC,"
7060                  strTmp02 = Mid(strTmp02, (intPos02 + 1))
7070                  intPos02 = InStr(strTmp02, ",")
7080                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC,"
7090                  strTmp02 = Mid(strTmp02, (intPos02 + 1))
7100                  intPos02 = InStr(strTmp02, ",")
7110                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC"
7120                  strTmp01 = strTmp01 & Mid(strTmp02, intPos02) & " DESC"
7130                  strSortNow = strTmp01
7140                Case 6
7150                  strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC,"
7160                  strTmp02 = Mid(strSortAsc, (intPos01 + 1))
7170                  intPos02 = InStr(strTmp02, ",")
7180                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC,"
7190                  strTmp02 = Mid(strTmp02, (intPos02 + 1))
7200                  intPos02 = InStr(strTmp02, ",")
7210                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC,"
7220                  strTmp02 = Mid(strTmp02, (intPos02 + 1))
7230                  intPos02 = InStr(strTmp02, ",")
7240                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC,"
7250                  strTmp02 = Mid(strTmp02, (intPos02 + 1))
7260                  intPos02 = InStr(strTmp02, ",")
7270                  strTmp01 = strTmp01 & Left(strTmp02, (intPos02 - 1)) & " DESC"
7280                  strTmp01 = strTmp01 & Mid(strTmp02, intPos02) & " DESC"
7290                  strSortNow = strTmp01
7300                End Select
7310              End If
7320              .Controls(strSortLbl).Caption = strArwDn
7330              .Controls(strSortLbl).ForeColor = CLR_DKRED
7340            Else
7350              strSortNow = strSortAsc
7360            End If
7370          End If
7380        End If
7390        .Controls(strSortLbl).Visible = True
7400        .Controls(strSortLine).Visible = True
7410        .OrderBy = strSortNow
7420        .OrderByOn = True
7430      End If
7440    End With

EXITP:
7450    Set ctl = Nothing
7460    Exit Sub

ERRH:
7470    Select Case ERR.Number
        Case Else
7480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7490    End Select
7500    Resume EXITP

End Sub

Public Sub RecalcTots_Set(blnRecalc As Boolean)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots_Set"

7610    With Me
7620      blnRecalcTots = blnRecalc
7630      .TimerInterval = 500&
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

Public Sub ColHeaders(blnEnable As Boolean)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "ColHeaders"

        Dim frm As Access.Form

7710    With Me
7720      Set frm = .Parent
7730      Select Case blnEnable
          Case True
7740        If .locasset_mark_lbl_dim_hi.Visible = True Then
7750          .locasset_mark_lbl.ForeColor = CLR_DKGRY2
7760          .locasset_mark_lbl_dim_hi.Visible = False
7770          .locasset_mark_lbl_line.BorderColor = CLR_DKGRY
7780          .locasset_mark_lbl_line_dim_hi.Visible = False
7790          .accountno_lbl.ForeColor = CLR_DKGRY2
7800          .accountno_lbl_dim_hi.Visible = False
7810          .accountno_lbl2.ForeColor = CLR_DKGRY2
7820          .accountno_lbl2_dim_hi.Visible = False
7830          .accountno_lbl_line.BorderColor = CLR_DKGRY
7840          .accountno_lbl_line_dim_hi.Visible = False
7850          .cusip_lbl.ForeColor = CLR_DKGRY2
7860          .cusip_lbl_dim_hi.Visible = False
7870          .cusip_lbl_line.BorderColor = CLR_DKGRY
7880          .cusip_lbl_line_dim_hi.Visible = False
7890          .asset_description_lbl.ForeColor = CLR_DKGRY2
7900          .asset_description_lbl_dim_hi.Visible = False
7910          .asset_description_lbl_line.BorderColor = CLR_DKGRY
7920          .asset_description_lbl_line_dim_hi.Visible = False
7930          .assettype_lbl.ForeColor = CLR_DKGRY2
7940          .assettype_lbl_dim_hi.Visible = False
7950          .assettype_lbl2.ForeColor = CLR_DKGRY2
7960          .assettype_lbl2_dim_hi.Visible = False
7970          .assettype_lbl_line.BorderColor = CLR_DKGRY
7980          .assettype_lbl_line_dim_hi.Visible = False
7990          .assetdate_lbl.ForeColor = CLR_DKGRY2
8000          .assetdate_lbl_dim_hi.Visible = False
8010          .assetdate_lbl2.ForeColor = CLR_DKGRY2
8020          .assetdate_lbl2_dim_hi.Visible = False
8030          .assetdate_lbl_line.BorderColor = CLR_DKGRY
8040          .assetdate_lbl_line_dim_hi.Visible = False
8050          .transdate_lbl.ForeColor = CLR_DKGRY2
8060          .transdate_lbl_dim_hi.Visible = False
8070          .transdate_lbl2.ForeColor = CLR_DKGRY2
8080          .transdate_lbl2_dim_hi.Visible = False
8090          .transdate_lbl_line.BorderColor = CLR_DKGRY
8100          .transdate_lbl_line_dim_hi.Visible = False
8110          .shareface_lbl.ForeColor = CLR_DKGRY2
8120          .shareface_lbl_dim_hi.Visible = False
8130          .shareface_lbl2.ForeColor = CLR_DKGRY2
8140          .shareface_lbl2_dim_hi.Visible = False
8150          .shareface_lbl_line.BorderColor = CLR_DKGRY
8160          .shareface_lbl_line_dim_hi.Visible = False
8170          With frm
8180            .cmdSelectAll.Enabled = True
8190            .cmdSelectAll_raised_img.Visible = True
8200            .cmdSelectAll_raised_semifocus_dots_img.Visible = False
8210            .cmdSelectAll_raised_focus_img.Visible = False
8220            .cmdSelectAll_raised_focus_dots_img.Visible = False
8230            .cmdSelectAll_sunken_focus_dots_img.Visible = False
8240            .cmdSelectAll_raised_img_dis.Visible = False
8250            .cmdSelectNone.Enabled = True
8260            .cmdSelectNone_raised_img.Visible = True
8270            .cmdSelectNone_raised_semifocus_dots_img.Visible = False
8280            .cmdSelectNone_raised_focus_img.Visible = False
8290            .cmdSelectNone_raised_focus_dots_img.Visible = False
8300            .cmdSelectNone_sunken_focus_dots_img.Visible = False
8310            .cmdSelectNone_raised_img_dis.Visible = False
8320          End With
8330        End If
8340      Case False
8350        If .locasset_mark_lbl_dim_hi.Visible = False Then
8360          .locasset_mark_lbl.ForeColor = WIN_CLR_DISF
8370          .locasset_mark_lbl_dim_hi.Visible = True
8380          .locasset_mark_lbl_line.BorderColor = WIN_CLR_DISR
8390          .locasset_mark_lbl_line_dim_hi.Visible = True
8400          .accountno_lbl.ForeColor = WIN_CLR_DISF
8410          .accountno_lbl_dim_hi.Visible = True
8420          .accountno_lbl2.ForeColor = WIN_CLR_DISF
8430          .accountno_lbl2_dim_hi.Visible = True
8440          .accountno_lbl_line.BorderColor = WIN_CLR_DISR
8450          .accountno_lbl_line_dim_hi.Visible = True
8460          .cusip_lbl.ForeColor = WIN_CLR_DISF
8470          .cusip_lbl_dim_hi.Visible = True
8480          .cusip_lbl_line.BorderColor = WIN_CLR_DISR
8490          .cusip_lbl_line_dim_hi.Visible = True
8500          .asset_description_lbl.ForeColor = WIN_CLR_DISF
8510          .asset_description_lbl_dim_hi.Visible = True
8520          .asset_description_lbl_line.BorderColor = WIN_CLR_DISR
8530          .asset_description_lbl_line_dim_hi.Visible = True
8540          .assettype_lbl.ForeColor = WIN_CLR_DISF
8550          .assettype_lbl_dim_hi.Visible = True
8560          .assettype_lbl2.ForeColor = WIN_CLR_DISF
8570          .assettype_lbl2_dim_hi.Visible = True
8580          .assettype_lbl_line.BorderColor = WIN_CLR_DISR
8590          .assettype_lbl_line_dim_hi.Visible = True
8600          .assetdate_lbl.ForeColor = WIN_CLR_DISF
8610          .assetdate_lbl_dim_hi.Visible = True
8620          .assetdate_lbl2.ForeColor = WIN_CLR_DISF
8630          .assetdate_lbl2_dim_hi.Visible = True
8640          .assetdate_lbl_line.BorderColor = WIN_CLR_DISR
8650          .assetdate_lbl_line_dim_hi.Visible = True
8660          .transdate_lbl.ForeColor = WIN_CLR_DISF
8670          .transdate_lbl_dim_hi.Visible = True
8680          .transdate_lbl2.ForeColor = WIN_CLR_DISF
8690          .transdate_lbl2_dim_hi.Visible = True
8700          .transdate_lbl_line.BorderColor = WIN_CLR_DISR
8710          .transdate_lbl_line_dim_hi.Visible = True
8720          .shareface_lbl.ForeColor = WIN_CLR_DISF
8730          .shareface_lbl_dim_hi.Visible = True
8740          .shareface_lbl2.ForeColor = WIN_CLR_DISF
8750          .shareface_lbl2_dim_hi.Visible = True
8760          .shareface_lbl_line.BorderColor = WIN_CLR_DISR
8770          .shareface_lbl_line_dim_hi.Visible = True
8780          With frm
8790            .cmdSelectAll.Enabled = False
8800            .cmdSelectAll_raised_img_dis.Visible = True
8810            .cmdSelectAll_raised_img.Visible = False
8820            .cmdSelectAll_raised_semifocus_dots_img.Visible = False
8830            .cmdSelectAll_raised_focus_img.Visible = False
8840            .cmdSelectAll_raised_focus_dots_img.Visible = False
8850            .cmdSelectAll_sunken_focus_dots_img.Visible = False
8860            .cmdSelectNone.Enabled = False
8870            .cmdSelectNone_raised_img_dis.Visible = True
8880            .cmdSelectNone_raised_img.Visible = False
8890            .cmdSelectNone_raised_semifocus_dots_img.Visible = False
8900            .cmdSelectNone_raised_focus_img.Visible = False
8910            .cmdSelectNone_raised_focus_dots_img.Visible = False
8920            .cmdSelectNone_sunken_focus_dots_img.Visible = False
8930          End With
8940        End If
8950      End Select
8960    End With

EXITP:
8970    Set frm = Nothing
8980    Exit Sub

ERRH:
8990    Select Case ERR.Number
        Case Else
9000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9010    End Select
9020    Resume EXITP

End Sub
