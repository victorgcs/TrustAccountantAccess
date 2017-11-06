VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssetType_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssetType_Sub"

'VGC 04/19/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmAssetType}
' **   Show Assets:      H {chkShowAssets on frmAssetType}
' **   Print Report:     P {cmdPrintReport on frmAssetType}
' **   Exit:             X {cmdClose on frmAssetType}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[assettype]"
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
150       If strCallingForm = "frmAssetType" Then

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170         lngSortLbl_Width = .Sort_lbl.Width

180         CLR_DISABLED_FG = CLR_DKGRY
190         CLR_DISABLED_BG = CLR_LTTEAL

200       Else
210         Cancel = -1
220       End If
230     End With

240     If Cancel = -1 Then
250       Beep
260       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
270     End If

EXITP:
280     Exit Sub

ERRH:
290     Select Case ERR.Number
        Case Else
300       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
310     End Select
320     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .assettype.ForeColor = CLR_DISABLED_FG
430       .assettype.BackColor = CLR_DISABLED_BG
440       .assettype_description.ForeColor = CLR_DISABLED_FG
450       .assettype_description.BackColor = CLR_DISABLED_BG
460       .assettypegroup_name.ForeColor = CLR_DISABLED_FG
470       .assettypegroup_name.BackColor = CLR_DISABLED_BG
480       .assettypegroup_sequence.ForeColor = CLR_DISABLED_FG
490       .assettypegroup_sequence.BackColor = CLR_DISABLED_BG
500       .taxcode_description.ForeColor = CLR_DISABLED_FG
510       .taxcode_description.BackColor = CLR_DISABLED_BG
520       .cusip.ForeColor = CLR_DISABLED_FG
530       .cusip.BackColor = CLR_DISABLED_BG
540       .asset_description.ForeColor = CLR_DISABLED_FG
550       .asset_description.BackColor = CLR_DISABLED_BG
560     End With

570     SortNow THIS_PROC  ' ** Procedure: Below.

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
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyUp
770         intRetVal = 0
780         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
790       Case vbKeyDown
800         intRetVal = 0
810         MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
820       Case vbKeyEscape
830         With Me
840           intRetVal = 0
850           .Parent.cmdClose_Click  ' ** Form Procedure: frmAssetType.
860         End With
870       End Select
880     End If

        ' ** Alt keys.
890     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
900       Select Case intRetVal
          Case vbKeyX
910         With Me
920           intRetVal = 0
930           .Parent.cmdClose_Click  ' ** Form Procedure: frmAssetType.
940         End With
950       End Select
960     End If

        ' ** Ctrl keys.
970     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
980       Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
990         With Me
1000          intRetVal = 0
1010          DoCmd.SelectObject acForm, .Parent.Name, False
1020          .Parent.chkShowAssets.SetFocus
1030        End With
1040      End Select
1050    End If

        ' ** Ctrl-Shift keys.
1060    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1070      Select Case intRetVal
          Case vbKeyF
1080        With Me
1090          intRetVal = 0
1100          DoCmd.SelectObject acForm, .Parent.Name, False
1110          .Parent.FocusHolder.SetFocus
1120        End With
1130      Case vbKeyTab, vbKeyReturn
1140        With Me
1150          intRetVal = 0
1160          DoCmd.SelectObject acForm, .Parent.Name, False
1170          .Parent.cmdClose.SetFocus
1180        End With
1190      End Select
1200    End If

EXITP:
1210    KeyCode = intRetVal
1220    Exit Sub

ERRH:
1230    intRetVal = 0
1240    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1250    Case Else
1260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1270    End Select
1280    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1320    intShiftDown = (Shift And acShiftMask) > 0
1330    intAltDown = (Shift And acAltMask) > 0
1340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1370        With Me
1380          intRetVal = 0
1390          lngRecsCur = RecCnt  ' ** Function: Below.
1400          If .CurrentRecord < lngRecsCur Then
1410            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1420            .FocusHolder.SetFocus
1430          Else
1440            DoCmd.SelectObject acForm, .Parent.Name, False
1450            .Parent.chkShowAssets.SetFocus
1460          End If
1470        End With
1480      End Select
1490    End If

        ' ** Shift keys.
1500    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1510      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1520        With Me
1530          intRetVal = 0
1540          If .CurrentRecord > 1 Then
1550            MoveRec acCmdRecordsGoToPrevious
1560            .FocusHolder.SetFocus
1570          Else
1580            DoCmd.SelectObject acForm, .Parent.Name, False
1590            .Parent.cmdClose.SetFocus
1600          End If
1610        End With
1620      End Select
1630    End If

EXITP:
1640    KeyCode = intRetVal
1650    Exit Sub

ERRH:
1660    intRetVal = 0
1670    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1680    Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Sub

Private Sub assettype_lbl_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_lbl_DblClick"

1810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub assettype_GotFocus()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_GotFocus"

1910    With Me
1920      .assettype.SelLength = 0
1930      .assettype.SelStart = 9
1940    End With

EXITP:
1950    Exit Sub

ERRH:
1960    Select Case ERR.Number
        Case Else
1970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1980    End Select
1990    Resume EXITP

End Sub

Private Sub assettype_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2020    intShiftDown = (Shift And acShiftMask) > 0
2030    intAltDown = (Shift And acAltMask) > 0
2040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2070        With Me
2080          intRetVal = 0
2090          .assettype_description.SetFocus
2100        End With
2110      End Select
2120    End If

        ' ** Shift keys.
2130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2150        With Me
2160          intRetVal = 0
2170          If .CurrentRecord > 1 Then
2180            MoveRec acCmdRecordsGoToPrevious
2190            .FocusHolder.SetFocus
2200          Else
2210            DoCmd.SelectObject acForm, .Parent.Name, False
2220            .Parent.cmdClose.SetFocus
2230          End If
2240        End With
2250      End Select
2260    End If

EXITP:
2270    KeyCode = intRetVal
2280    Exit Sub

ERRH:
2290    intRetVal = 0
2300    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2310    Case Else
2320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2330    End Select
2340    Resume EXITP

End Sub

Private Sub assettype_description_lbl_DblClick(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_description_lbl_DblClick"

2410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub assettype_description_GotFocus()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_description_GotFocus"

2510    With Me
2520      .assettype_description.SelLength = 0
2530      .assettype_description.SelStart = 99
2540    End With

EXITP:
2550    Exit Sub

ERRH:
2560    Select Case ERR.Number
        Case Else
2570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2580    End Select
2590    Resume EXITP

End Sub

Private Sub assettype_description_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2620    intShiftDown = (Shift And acShiftMask) > 0
2630    intAltDown = (Shift And acAltMask) > 0
2640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2670        With Me
2680          intRetVal = 0
2690          If .assettypegroup_name.Visible = True And .assettypegroup_name.Enabled = True Then
2700            .assettypegroup_name.SetFocus
2710          ElseIf .cusip.Visible = True And .cusip.Enabled = True Then
2720            .cusip.SetFocus
2730          End If
2740        End With
2750      End Select
2760    End If

        ' ** Shift keys.
2770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2780      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2790        With Me
2800          intRetVal = 0
2810          .assettype.SetFocus
2820        End With
2830      End Select
2840    End If

EXITP:
2850    KeyCode = intRetVal
2860    Exit Sub

ERRH:
2870    intRetVal = 0
2880    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2890    Case Else
2900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2910    End Select
2920    Resume EXITP

End Sub

Private Sub assettypegroup_name_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_name_lbl_DblClick"

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

Private Sub assettypegroup_name_GotFocus()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_name_GotFocus"

3110    With Me
3120      .assettypegroup_name.SelLength = 0
3130      .assettypegroup_name.SelStart = 99
3140    End With

EXITP:
3150    Exit Sub

ERRH:
3160    Select Case ERR.Number
        Case Else
3170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3180    End Select
3190    Resume EXITP

End Sub

Private Sub assettypegroup_name_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_name_KeyDown"

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
3290          .assettypegroup_sequence.SetFocus
3300        End With
3310      End Select
3320    End If

        ' ** Shift keys.
3330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3350        With Me
3360          intRetVal = 0
3370          .assettype_description.SetFocus
3380        End With
3390      End Select
3400    End If

EXITP:
3410    KeyCode = intRetVal
3420    Exit Sub

ERRH:
3430    intRetVal = 0
3440    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3450    Case Else
3460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3470    End Select
3480    Resume EXITP

End Sub

Private Sub assettypegroup_sequence_lbl_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_sequence_lbl_DblClick"

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

Private Sub assettypegroup_sequence_GotFocus()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_sequence_GotFocus"

3610    With Me
3620      .assettypegroup_sequence.SelLength = 0
3630      .assettypegroup_sequence.SelStart = 9
3640    End With

EXITP:
3650    Exit Sub

ERRH:
3660    Select Case ERR.Number
        Case Else
3670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3680    End Select
3690    Resume EXITP

End Sub

Private Sub assettypegroup_sequence_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "assettypegroup_sequence_KeyDown"

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
3790          .Interest.SetFocus
3800        End With
3810      End Select
3820    End If

        ' ** Shift keys.
3830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3850        With Me
3860          intRetVal = 0
3870          .assettypegroup_name.SetFocus
3880        End With
3890      End Select
3900    End If

EXITP:
3910    KeyCode = intRetVal
3920    Exit Sub

ERRH:
3930    intRetVal = 0
3940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3950    Case Else
3960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3970    End Select
3980    Resume EXITP

End Sub

Private Sub Interest_lbl_DblClick(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Interest_lbl_DblClick"

4010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub Interest_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Interest_KeyDown"

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
4190          .Dividend.SetFocus
4200        End With
4210      End Select
4220    End If

        ' ** Shift keys.
4230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4250        With Me
4260          intRetVal = 0
4270          .assettypegroup_sequence.SetFocus
4280        End With
4290      End Select
4300    End If

EXITP:
4310    KeyCode = intRetVal
4320    Exit Sub

ERRH:
4330    intRetVal = 0
4340    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4350    Case Else
4360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4370    End Select
4380    Resume EXITP

End Sub

Private Sub Dividend_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Dividend_lbl_DblClick"

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

Private Sub Dividend_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Dividend_KeyDown"

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
4590          If .taxcode_description.Enabled = True Then
4600            .taxcode_description.SetFocus
4610          Else
4620            lngRecsCur = RecCnt  ' ** Function: Below.
4630            If .CurrentRecord < lngRecsCur Then
4640              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4650              .FocusHolder.SetFocus
4660            Else
4670              DoCmd.SelectObject acForm, .Parent.Name, False
4680              .Parent.chkShowAssets.SetFocus
4690            End If
4700          End If
4710        End With
4720      End Select
4730    End If

        ' ** Shift keys.
4740    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4750      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4760        With Me
4770          intRetVal = 0
4780          .Interest.SetFocus
4790        End With
4800      End Select
4810    End If

EXITP:
4820    KeyCode = intRetVal
4830    Exit Sub

ERRH:
4840    intRetVal = 0
4850    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4860    Case Else
4870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4880    End Select
4890    Resume EXITP

End Sub

Private Sub taxcode_description_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_lbl_DblClick"

4910    With Me
4920      If .taxcode_description.Enabled = True Then
4930        SortNow THIS_PROC  ' ** Procedure: Below.
4940      End If
4950    End With

EXITP:
4960    Exit Sub

ERRH:
4970    Select Case ERR.Number
        Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Sub

Private Sub taxcode_description_GotFocus()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_GotFocus"

5110    With Me
5120      .taxcode_description.SelLength = 0
5130      .taxcode_description.SelStart = 99
5140    End With

EXITP:
5150    Exit Sub

ERRH:
5160    Select Case ERR.Number
        Case Else
5170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5180    End Select
5190    Resume EXITP

End Sub

Private Sub taxcode_description_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "taxcode_description_KeyDown"

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
5290          lngRecsCur = RecCnt  ' ** Function: below.
5300          If .CurrentRecord < lngRecsCur Then
5310            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5320            .FocusHolder.SetFocus
5330          Else
5340            DoCmd.SelectObject acForm, .Parent.Name
5350            .Parent.chkShowAssets.SetFocus
5360          End If
5370        End With
5380      End Select
5390    End If

        ' ** Shift keys.
5400    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5410      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5420        With Me
5430          intRetVal = 0
5440          .Dividend.SetFocus
5450        End With
5460      End Select
5470    End If

EXITP:
5480    KeyCode = intRetVal
5490    Exit Sub

ERRH:
5500    intRetVal = 0
5510    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5520    Case Else
5530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5540    End Select
5550    Resume EXITP

End Sub

Private Sub cusip_lbl_DblClick(Cancel As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_lbl_DblClick"

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

Private Sub cusip_GotFocus()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_GotFocus"

5710    With Me
5720      .cusip.SelLength = 0
5730      .cusip.SelStart = 10
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

Private Sub cusip_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5820    intShiftDown = (Shift And acShiftMask) > 0
5830    intAltDown = (Shift And acAltMask) > 0
5840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5870        With Me
5880          intRetVal = 0
5890          .asset_description.SetFocus
5900        End With
5910      End Select
5920    End If

        ' ** Shift keys.
5930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5950        With Me
5960          intRetVal = 0
5970          .assettype_description.SetFocus
5980        End With
5990      End Select
6000    End If

EXITP:
6010    KeyCode = intRetVal
6020    Exit Sub

ERRH:
6030    intRetVal = 0
6040    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6050    Case Else
6060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6070    End Select
6080    Resume EXITP

End Sub

Private Sub asset_description_lbl_DblClick(Cancel As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_lbl_DblClick"

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

Private Sub asset_description_GotFocus()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_GotFocus"

6210    With Me
6220      .asset_description.SelLength = 0
6230      .asset_description.SelStart = 0
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

Private Sub asset_description_KeyDown(KeyCode As Integer, Shift As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "asset_description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6320    intShiftDown = (Shift And acShiftMask) > 0
6330    intAltDown = (Shift And acAltMask) > 0
6340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6370        With Me
6380          intRetVal = 0
6390          lngRecsCur = RecCnt  ' ** Function: below.
6400          If .CurrentRecord < lngRecsCur Then
6410            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
6420            .FocusHolder.SetFocus
6430          Else
6440            DoCmd.SelectObject acForm, .Parent.Name
6450            .Parent.chkShowAssets.SetFocus
6460          End If
6470        End With
6480      End Select
6490    End If

        ' ** Shift keys.
6500    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6510      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6520        With Me
6530          intRetVal = 0
6540          .cusip.SetFocus
6550        End With
6560      End Select
6570    End If

EXITP:
6580    KeyCode = intRetVal
6590    Exit Sub

ERRH:
6600    intRetVal = 0
6610    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6620    Case Else
6630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6640    End Select
6650    Resume EXITP

End Sub

Public Sub ShowAsset(blnShow As Boolean)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAsset"

        Dim blnSortHere As Boolean, blnResort As Boolean, blnIsDesc As Boolean
        Dim strFilter As String

6710    With Me
6720      blnSortHere = False: blnResort = False
6730      .FocusHolder.SetFocus
6740      DoEvents
6750      Select Case blnShow
          Case True
6760        If .Sort_line.Left = .assettypegroup_name_lbl.Left Then blnSortHere = True
6770        .assettypegroup_name.Visible = False
6780        .assettypegroup_name_lbl.Visible = False
6790        .assettypegroup_name_lbl_line.Visible = False
6800        If blnSortHere = True Then
6810          .Sort_line.Visible = False
6820          .Sort_lbl.Visible = False
6830          blnResort = True
6840        End If
6850        blnSortHere = False
6860        If .Sort_line.Left = .assettypegroup_sequence_lbl.Left Then blnSortHere = True
6870        .assettypegroup_sequence.Visible = False
6880        .assettypegroup_sequence_lbl.Visible = False
6890        .assettypegroup_sequence_lbl2.Visible = False
6900        .assettypegroup_sequence_lbl_line.Visible = False
6910        If blnSortHere = True Then
6920          .Sort_line.Visible = False
6930          .Sort_lbl.Visible = False
6940          blnResort = True
6950        End If
6960        blnSortHere = False
6970        If .Sort_line.Left = .Interest_lbl.Left Then blnSortHere = True
6980        .Interest.Visible = False
6990        .Interest_lbl.Visible = False
7000        .Interest_lbl_line.Visible = False
7010        If blnSortHere = True Then
7020          .Sort_line.Visible = False
7030          .Sort_lbl.Visible = False
7040          blnResort = True
7050        End If
7060        blnSortHere = False
7070        If .Sort_line.Left = .Dividend_lbl.Left Then blnSortHere = True
7080        .Dividend.Visible = False
7090        .Dividend_lbl.Visible = False
7100        .Dividend_lbl_line.Visible = False
7110        If blnSortHere = True Then
7120          .Sort_line.Visible = False
7130          .Sort_lbl.Visible = False
7140          blnResort = True
7150        End If
7160        blnSortHere = False
7170        If .Sort_line.Left = .TaxCode_Description_lbl.Left Then blnSortHere = True
7180        .taxcode_description.Visible = False
7190        .TaxCode_Description_lbl.Visible = False
7200        .taxcode_description_lbl2.Visible = False
7210        .Parent.taxcode_description_lbl3.Visible = False
7220        .Parent.taxcode_description_lbl4.Visible = False
7230        .TaxCode_Description_lbl_line.Visible = False
7240        If gblnIncomeTaxCoding = False Then
7250          .taxcode_description_lbl_dim_hi.Visible = False
7260          .taxcode_description_lbl2_dim_hi.Visible = False
7270          .Parent.taxcode_description_lbl3_dim_hi.Visible = False
7280          .Parent.taxcode_description_lbl4_dim_hi.Visible = False
7290          .taxcode_description_lbl_line_dim_hi.Visible = False
7300          .taxcode_description_lbl3.Visible = False
7310        End If
7320        If blnSortHere = True Then
7330          .Sort_line.Visible = False
7340          .Sort_lbl.Visible = False
7350          blnResort = True
7360        End If
7370        blnSortHere = False
7380        .cusip.Visible = True
7390        .cusip_lbl.Visible = True
7400        .cusip_lbl_line.Visible = True
7410        .asset_description.Visible = True
7420        .asset_description_lbl.Visible = True
7430        .asset_description_lbl_line.Visible = True
            ' ** A check box could be added for this.
7440        strFilter = "[IsHidden] = False"
7450        .Filter = strFilter
7460        .FilterOn = True
7470      Case False
7480        If .Sort_line.Left = .cusip_lbl.Left Then blnSortHere = True
7490        .cusip.Visible = False
7500        .cusip_lbl.Visible = False
7510        .cusip_lbl_line.Visible = False
7520        If blnSortHere = True Then
7530          .Sort_line.Visible = False
7540          .Sort_lbl.Visible = False
7550          blnResort = True
7560        End If
7570        blnSortHere = False
7580        If .Sort_line.Left = .asset_description_lbl.Left Then blnSortHere = True
7590        .asset_description.Visible = False
7600        .asset_description_lbl.Visible = False
7610        .asset_description_lbl_line.Visible = False
7620        If blnSortHere = True Then
7630          .Sort_line.Visible = False
7640          .Sort_lbl.Visible = False
7650          blnResort = True
7660        End If
7670        blnSortHere = False
7680        .assettypegroup_name.Visible = True
7690        .assettypegroup_name_lbl.Visible = True
7700        .assettypegroup_name_lbl_line.Visible = True
7710        .assettypegroup_sequence.Visible = True
7720        .assettypegroup_sequence_lbl.Visible = True
7730        .assettypegroup_sequence_lbl2.Visible = True
7740        .assettypegroup_sequence_lbl_line.Visible = True
7750        .Interest.Visible = True
7760        .Interest_lbl.Visible = True
7770        .Interest_lbl_line.Visible = True
7780        .Dividend.Visible = True
7790        .Dividend_lbl.Visible = True
7800        .Dividend_lbl_line.Visible = True
7810        .taxcode_description.Visible = True
7820        .TaxCode_Description_lbl.Visible = True
7830        .taxcode_description_lbl2.Visible = True
7840        .Parent.taxcode_description_lbl3.Visible = True
7850        .Parent.taxcode_description_lbl4.Visible = True
7860        .TaxCode_Description_lbl_line.Visible = True
7870        If gblnIncomeTaxCoding = False Then
7880          .taxcode_description_lbl_dim_hi.Visible = True
7890          .taxcode_description_lbl2_dim_hi.Visible = True
7900          .Parent.taxcode_description_lbl3_dim_hi.Visible = True
7910          .Parent.taxcode_description_lbl4_dim_hi.Visible = True
7920          .taxcode_description_lbl_line_dim_hi.Visible = True
7930          .taxcode_description_lbl3.Visible = True
7940        End If
            ' ** A check box could be added for this.
7950        strFilter = vbNullString
7960        .Filter = strFilter
7970        .FilterOn = False
7980      End Select
7990      DoEvents
8000      If InStr(strSortNow, " DESC") > 0 Then blnIsDesc = True Else blnIsDesc = False
8010      If blnResort = True Then
8020        SortNow "Form_Load"  ' ** Procedure: Below.
8030      ElseIf Left(strSortNow, 11) = "[assettype]" Then
8040        SortNow "assettype_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
8050      ElseIf Left(strSortNow, 23) = "[assettype_description]" Then
8060        SortNow "assettype_description_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
8070      ElseIf Left(strSortNow, 21) = "[assettypegroup_name]" Then
8080        SortNow "assettypegroup_name_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
8090      ElseIf Left(strSortNow, 25) = "[assettypegroup_sequence]" Then
8100        SortNow "assettypegroup_sequence_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
8110      ElseIf Left(strSortNow, 10) = "[Interest]" Then
8120        SortNow "Interest_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
8130      ElseIf Left(strSortNow, 10) = "[Dividend]" Then
8140        SortNow "Dividend_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
8150      ElseIf Left(strSortNow, 21) = "[taxcode_description]" Then
8160        SortNow "taxcode_description_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
8170      ElseIf Left(strSortNow, 7) = "[cusip]" Then
8180        SortNow "cusip_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
8190      ElseIf Left(strSortNow, 19) = "[asset_description]" Then
8200        SortNow "asset_description_lbl_DblClick", True, blnIsDesc  ' ** Procedure: Below.
8210      End If
8220    End With

EXITP:
8230    Exit Sub

ERRH:
8240    Select Case ERR.Number
        Case Else
8250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8260    End Select
8270    Resume EXITP

End Sub

Public Function RecCnt() As Long

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

8310    lngRetVal = 0&

8320    With Me
8330      Set rst = .RecordsetClone
8340      With rst
8350        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
8360        Else
8370          .MoveLast
8380          lngRetVal = .RecordCount
8390        End If
8400        .Close
8410      End With
8420    End With

EXITP:
8430    Set rst = Nothing
8440    RecCnt = lngRetVal
8450    Exit Function

ERRH:
8460    lngRetVal = 0&
8470    Select Case ERR.Number
        Case Else
8480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8490    End Select
8500    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

8610    Select Case intWhere
        Case acCmdRecordsGoToFirst
8620      DoCmd.RunCommand acCmdRecordsGoToFirst
8630    Case acCmdRecordsGoToLast
8640      DoCmd.RunCommand acCmdRecordsGoToLast
8650    Case acCmdRecordsGoToPrevious
8660      DoCmd.RunCommand acCmdRecordsGoToPrevious
8670    Case acCmdRecordsGoToNext
8680      DoCmd.RunCommand acCmdRecordsGoToNext
8690    Case Else
8700      If IsMissing(varID) = False Then
8710        With Me
8720          Set rst = .RecordsetClone
8730          With rst
8740            .FindFirst "[assettype] = '" & varID & "'"
8750            If .NoMatch = False Then
8760              Me.Bookmark = .Bookmark
8770            End If
8780            .Close
8790          End With
8800        End With
8810      End If
8820    End Select

EXITP:
8830    Set rst = Nothing
8840    Exit Sub

ERRH:
8850    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8860    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8870    Case Else
8880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8890    End Select
8900    Resume EXITP

End Sub

Public Sub SortNow(strProc As String, Optional varShow As Variant, Optional varIsDesc As Variant)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim blnDontReverse As Boolean, blnIsDesc As Boolean
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [assettype], [asset_description]"

9010    With Me
9020      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
9030        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
9040      End If
9050      Select Case IsMissing(varShow)
          Case True
9060        blnDontReverse = False
9070      Case False
9080        blnDontReverse = CBool(varShow)
9090      End Select
9100      Select Case IsMissing(varIsDesc)
          Case True
9110        blnIsDesc = False
9120      Case False
9130        blnIsDesc = CBool(varIsDesc)
9140      End Select
9150      .Controls(strSortLbl).Visible = False
9160      .Controls(strSortLine).Visible = False
9170      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
9180      If strProc = "Form_Load" Then
9190        strCalled = "assettype"
9200        strSortNow = strSortOrig
9210        If .Parent.chkShowAssets = True Then
9220          strSortNow = strSortNow & ", [asset_description]"
9230        End If
9240        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9250        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + lngTpp)
9260        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9270        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
9280        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
9290        .Controls(strSortLbl).Top = lngSortLbl_Top
9300        .Controls(strSortLbl).Left = lngSortLbl_Left
9310        .Controls(strSortLine).Top = lngSortLine_Top
9320        .Controls(strSortLine).Left = lngSortLine_Left
9330        .Controls(strSortLine).Width = lngSortLine_Width
9340        .Controls(strSortLbl).Caption = strArwUp
9350        .Controls(strSortLbl).ForeColor = CLR_DKBLU
9360      Else
9370        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
9380        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9390        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
9400        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9410        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
9420        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
9430        .Controls(strSortLbl).Left = lngSortLbl_Left
9440        .Controls(strSortLbl).Top = lngSortLbl_Top
9450        .Controls(strSortLine).Top = lngSortLine_Top
9460        .Controls(strSortLine).Left = lngSortLine_Left
9470        .Controls(strSortLine).Width = lngSortLine_Width
9480        If strCalled = "assettype" Then
9490          lngSortLbl_Left = (lngSortLbl_Left + lngTpp)
9500          .Controls(strSortLbl).Left = lngSortLbl_Left
9510          strSortAsc = strSortOrig
9520          If .Parent.chkShowAssets = True Then
9530            strSortAsc = strSortAsc & ", [asset_description]"
9540          End If
9550          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
9560            strSortNow = "[assettype] DESC"
9570            If .Parent.chkShowAssets = True Then
9580              strSortNow = strSortNow & ", [asset_description] DESC"
9590            End If
9600            .Controls(strSortLbl).Caption = strArwDn
9610            .Controls(strSortLbl).ForeColor = CLR_DKRED
9620          Else
9630            strSortNow = strSortAsc
9640            .Controls(strSortLbl).Caption = strArwUp
9650            .Controls(strSortLbl).ForeColor = CLR_DKBLU
9660          End If
9670        Else
9680          .Controls(strSortLbl).Caption = strArwUp
9690          .Controls(strSortLbl).ForeColor = CLR_DKBLU
9700          Select Case strCalled
              Case "assettype_description"
9710            strSortAsc = "[assettype_description]"
9720          Case "assettypegroup_name"
9730            strSortAsc = "[assettypegroup_name]" & strStdAsc
9740          Case "assettypegroup_sequence"
9750            strSortAsc = "[assettypegroup_sequence]" & strStdAsc
9760          Case "Interest"
9770            strSortAsc = "[Interest]" & strStdAsc
9780            lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
9790            .Controls(strSortLbl).Left = lngSortLbl_Left
9800          Case "Dividend"
9810            strSortAsc = "[Dividend]" & strStdAsc
9820            lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
9830            .Controls(strSortLbl).Left = lngSortLbl_Left
9840          Case "taxcode_description"
9850            strSortAsc = "[taxcode_description]" & strStdAsc
9860          Case "cusip"
9870            strSortAsc = "[cusip]"
9880          Case "asset_description"
9890            strSortAsc = "[asset_description], [cusip]"
9900          End Select
9910          If ((strSortNow = strSortAsc And blnDontReverse = False) Or (blnIsDesc = True)) Then
9920            intCnt = CharCnt(strSortAsc, ",") + 1
9930            Select Case intCnt
                Case 1
9940              strTmp01 = strSortAsc & " DESC"
9950            Case 2
9960              intPos01 = InStr(strSortAsc, ",")
9970              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9980              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
9990            Case 3
10000             intPos01 = InStr(strSortAsc, ",")
10010             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10020             strTmp02 = Mid(strSortAsc, intPos01)
10030             intPos01 = InStr(2, strTmp02, ",")
10040             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10050             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10060           End Select
10070           strSortNow = strTmp01
10080           .Controls(strSortLbl).Caption = strArwDn
10090           .Controls(strSortLbl).ForeColor = CLR_DKRED
10100         Else
10110           strSortNow = strSortAsc
10120         End If
10130       End If
10140     End If
10150     .Controls(strSortLbl).Visible = True
10160     .Controls(strSortLine).Visible = True
10170     .OrderBy = strSortNow
10180     .OrderByOn = True
10190   End With

        ' ** Potential Sorts:
        ' **   [assettype]
        ' **   [assettype_description]
        ' **   [assettypegroup_name], [assettype]
        ' **   [assettypegroup_sequence], [assettype]
        ' **   [Interest], [assettype]
        ' **   [Dividend], [assettype]
        ' **   [taxcode_description], [assettype]

EXITP:
10200   Exit Sub

ERRH:
10210   Select Case ERR.Number
        Case Else
10220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10230   End Select
10240   Resume EXITP

End Sub

Public Function SortNow_Get() As String

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

10310   strRetVal = strSortNow

EXITP:
10320   SortNow_Get = strRetVal
10330   Exit Function

ERRH:
10340   strRetVal = strSortOrig
10350   Select Case ERR.Number
        Case Else
10360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10370   End Select
10380   Resume EXITP

End Function
