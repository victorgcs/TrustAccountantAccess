VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Misc_LTCL_Detail_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Misc_LTCL_Detail_Sub"

'VGC 07/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel on frmMap_Misc_LTCL_Detail}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Misc on frmMap_Misc_LTCL_Detail}
' **   OK:               O {cmdOK on frmMap_Misc_LTCL_Detail}
' **   Print Report:     P {cmdPrintReport on frmMap_Misc_LTCL_Detail}
' **   Short Name:       S {shortname}
' **   Tax Codes:        T {cmbTaxCodes_Misc on frmMap_Misc_LTCL_Detail}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots on frmMap_Misc_LTCL_Detail}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description on frmMap_Misc_LTCL_Detail}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[alphasort]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngTmp01 As Long
Private lngRecsCur As Long, lngTpp As Long, blnCurrID As Boolean, lngCurrID As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmMap_Misc_LTCL_Detail" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         lngSortLbl_Width = .Sort_lbl.Width
200         .form_width_line.Left = 0&  ' ** Just get it out of the way.

210         lngTmp01 = .total_pcash_usd_lbl.Left

220       Else
230         Cancel = -1
240       End If
250     End With

260     If Cancel = -1 Then
270       Beep
280       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
290     End If

EXITP:
300     Exit Sub

ERRH:
310     Select Case ERR.Number
        Case Else
320       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
330     End Select
340     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .accountno.ForeColor = CLR_DISABLED_FG
430       .accountno.BackColor = CLR_DISABLED_BG
440       .shortname.ForeColor = CLR_DISABLED_FG
450       .shortname.BackColor = CLR_DISABLED_BG
460       .total_shareface.ForeColor = CLR_DISABLED_FG
470       .total_shareface.BackColor = CLR_DISABLED_BG
480       .total_pcash.ForeColor = CLR_DISABLED_FG
490       .total_pcash.BackColor = CLR_VLTRED
500       .total_pcash_usd.ForeColor = CLR_DISABLED_FG
510       .total_pcash_usd.BackColor = CLR_VLTRED
520       SortNow THIS_PROC  ' ** Procedure: Below
530       DoEvents
540     End With

EXITP:
550     Exit Sub

ERRH:
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc Totals:    F5 {RecalcTots on frmMap_Misc_LTCL_Detail}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Short Name:       S {shortname}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Comments:         M {description on frmMap_Misc_LTCL_Detail}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyUp
670         intRetVal = 0
680         MoveRec acCmdRecordsGoToPrevious
690       Case vbKeyDown
700         intRetVal = 0
710         MoveRec acCmdRecordsGoToNext
720       Case vbKeyEscape
730         With Me
740           intRetVal = 0
750           .Parent.cmdCancel_Click  ' ** Form Procedure: frmMap_Misc_LTCL_Detail.
760         End With
770       Case vbKeyF5
780         With Me
790           intRetVal = 0
800           .Parent.RecalcTots  ' ** Form Function: frmMap_Misc_LTCL_Detail.
810         End With
820       End Select
830     End If

        ' ** Alt keys.
840     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
850       Select Case intRetVal
          Case vbKeyS
860         With Me
870           intRetVal = 0
880           .shortname.SetFocus
890         End With
900       End Select
910     End If

        ' ** Ctrl keys.
920     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
930       Select Case intRetVal
          Case vbKeyM
940         With Me
950           intRetVal = 0
960           DoCmd.SelectObject acForm, .Parent.Name, False
970           .Parent.description.SetFocus
980         End With
990       Case vbKeyTab
1000        With Me
1010          intRetVal = 0
1020          DoCmd.SelectObject acForm, .Parent.Name, False
1030          .Parent.description.SetFocus
1040        End With
1050      Case vbKeyPageUp, vbKeyHome
1060        intRetVal = 0
1070        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1080      Case vbKeyPageDown, vbKeyEnd
1090        intRetVal = 0
1100        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1110      End Select
1120    End If

        ' ** Ctrl-Shift keys.
1130    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1140      Select Case intRetVal
          Case vbKeyF
1150        With Me
1160          intRetVal = 0
1170          DoCmd.SelectObject acForm, .Parent.Name, False
1180          .Parent.FocusHolder.SetFocus
1190        End With
1200      Case vbKeyTab
1210        With Me
1220          intRetVal = 0
1230          DoCmd.SelectObject acForm, .Parent.Name, False
1240          .Parent.cmdCancel.SetFocus
1250        End With
1260      End Select
1270    End If

EXITP:
1280    KeyCode = intRetVal
1290    Exit Sub

ERRH:
1300    intRetVal = 0
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Sub

Private Sub Form_Timer()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1410    With Me
1420      .TimerInterval = 0&
1430      blnCurrID = .Parent.curr_id.Visible
1440      lngCurrID = .Parent.curr_id
1450      IncludeCurrency  ' ** Procedure: Below.
1460      .Parent.IncludeCurrency True  ' ** Form Procedure: frmMap_Misc_LTCL_Detail.
1470    End With

EXITP:
1480    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1620    intShiftDown = (Shift And acShiftMask) > 0
1630    intAltDown = (Shift And acAltMask) > 0
1640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1670        With Me
1680          intRetVal = 0
1690          lngRecsCur = RecCnt  ' ** Function: Below.
1700          If .CurrentRecord < lngRecsCur Then
1710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1720            .FocusHolder.SetFocus
1730          Else
1740            DoCmd.SelectObject acForm, .Parent.Name, False
1750            .Parent.description.SetFocus
1760          End If
1770        End With
1780      End Select
1790    End If

        ' ** Shift keys.
1800    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1810      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1820        With Me
1830          intRetVal = 0
1840          If .CurrentRecord > 1 Then
1850            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1860            .FocusHolder.SetFocus
1870          Else
1880            DoCmd.SelectObject acForm, .Parent.Name, False
1890            .Parent.cmdCancel.SetFocus
1900          End If
1910        End With
1920      End Select
1930    End If

EXITP:
1940    KeyCode = intRetVal
1950    Exit Sub

ERRH:
1960    intRetVal = 0
1970    Select Case ERR.Number
        Case Else
1980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1990    End Select
2000    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

2110    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2270        With Me
2280          intRetVal = 0
2290          lngRecsCur = RecCnt  ' ** Function: Below.
2300          If .CurrentRecord < lngRecsCur Then
2310            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2320            .FocusHolder.SetFocus
2330          Else
2340            DoCmd.SelectObject acForm, .Parent.Name, False
2350            .Parent.description.SetFocus
2360          End If
2370        End With
2380      End Select
2390    End If

        ' ** Shift keys.
2400    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2410      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2420        With Me
2430          intRetVal = 0
2440          If .CurrentRecord > 1 Then
2450            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2460            .FocusHolder.SetFocus
2470          Else
2480            DoCmd.SelectObject acForm, .Parent.Name, False
2490            .Parent.cmdCancel.SetFocus
2500          End If
2510        End With
2520      End Select
2530    End If

EXITP:
2540    KeyCode = intRetVal
2550    Exit Sub

ERRH:
2560    intRetVal = 0
2570    Select Case ERR.Number
        Case Else
2580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2590    End Select
2600    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

2710    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub shortname_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_KeyDown"

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
2890          lngRecsCur = RecCnt  ' ** Function: Below.
2900          If .CurrentRecord < lngRecsCur Then
2910            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2920            .FocusHolder.SetFocus
2930          Else
2940            DoCmd.SelectObject acForm, .Parent.Name, False
2950            .Parent.description.SetFocus
2960          End If
2970        End With
2980      End Select
2990    End If

        ' ** Shift keys.
3000    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3010      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3020        With Me
3030          intRetVal = 0
3040          If .CurrentRecord > 1 Then
3050            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3060            .FocusHolder.SetFocus
3070          Else
3080            DoCmd.SelectObject acForm, .Parent.Name, False
3090            .Parent.cmdCancel.SetFocus
3100          End If
3110        End With
3120      End Select
3130    End If

EXITP:
3140    KeyCode = intRetVal
3150    Exit Sub

ERRH:
3160    intRetVal = 0
3170    Select Case ERR.Number
        Case Else
3180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3190    End Select
3200    Resume EXITP

End Sub

Private Sub total_shareface_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_lbl_DblClick"

3310    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub total_shareface_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3420    intShiftDown = (Shift And acShiftMask) > 0
3430    intAltDown = (Shift And acAltMask) > 0
3440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3470        With Me
3480          intRetVal = 0
3490          lngRecsCur = RecCnt  ' ** Function: Below.
3500          If .CurrentRecord < lngRecsCur Then
3510            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3520            .FocusHolder.SetFocus
3530          Else
3540            DoCmd.SelectObject acForm, .Parent.Name, False
3550            .Parent.description.SetFocus
3560          End If
3570        End With
3580      End Select
3590    End If

        ' ** Shift keys.
3600    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3610      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3620        With Me
3630          intRetVal = 0
3640          If .CurrentRecord > 1 Then
3650            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3660            .FocusHolder.SetFocus
3670          Else
3680            DoCmd.SelectObject acForm, .Parent.Name, False
3690            .Parent.cmdCancel.SetFocus
3700          End If
3710        End With
3720      End Select
3730    End If

EXITP:
3740    KeyCode = intRetVal
3750    Exit Sub

ERRH:
3760    intRetVal = 0
3770    Select Case ERR.Number
        Case Else
3780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3790    End Select
3800    Resume EXITP

End Sub

Private Sub total_pcash_lbl_DblClick(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "total_pcash_lbl_DblClick"

3910    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub total_pcash_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "total_pcash_KeyDown"

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
4090          lngRecsCur = RecCnt  ' ** Function: Below.
4100          If .CurrentRecord < lngRecsCur Then
4110            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4120            .FocusHolder.SetFocus
4130          Else
4140            DoCmd.SelectObject acForm, .Parent.Name, False
4150            .Parent.description.SetFocus
4160          End If
4170        End With
4180      End Select
4190    End If

        ' ** Shift keys.
4200    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4210      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4220        With Me
4230          intRetVal = 0
4240          If .CurrentRecord > 1 Then
4250            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4260            .FocusHolder.SetFocus
4270          Else
4280            DoCmd.SelectObject acForm, .Parent.Name, False
4290            .Parent.cmdCancel.SetFocus
4300          End If
4310        End With
4320      End Select
4330    End If

EXITP:
4340    KeyCode = intRetVal
4350    Exit Sub

ERRH:
4360    intRetVal = 0
4370    Select Case ERR.Number
        Case Else
4380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4390    End Select
4400    Resume EXITP

End Sub

Private Sub total_pcash_usd_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "total_pcash_usd_lbl_DblClick"

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

Private Sub total_pcash_usd_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "total_pcash_usd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4620    intShiftDown = (Shift And acShiftMask) > 0
4630    intAltDown = (Shift And acAltMask) > 0
4640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4670        With Me
4680          intRetVal = 0
4690          lngRecsCur = RecCnt  ' ** Function: Below.
4700          If .CurrentRecord < lngRecsCur Then
4710            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4720            .FocusHolder.SetFocus
4730          Else
4740            DoCmd.SelectObject acForm, .Parent.Name, False
4750            .Parent.description.SetFocus
4760          End If
4770        End With
4780      End Select
4790    End If

        ' ** Shift keys.
4800    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4810      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4820        With Me
4830          intRetVal = 0
4840          If .CurrentRecord > 1 Then
4850            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4860            .FocusHolder.SetFocus
4870          Else
4880            DoCmd.SelectObject acForm, .Parent.Name, False
4890            .Parent.cmdCancel.SetFocus
4900          End If
4910        End With
4920      End Select
4930    End If

EXITP:
4940    KeyCode = intRetVal
4950    Exit Sub

ERRH:
4960    intRetVal = 0
4970    Select Case ERR.Number
        Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Sub

Public Sub IncludeCurrency()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim strTmp01 As String, intTmp02 As Integer

5110    With Me
5120      If blnCurrID = True And lngCurrID <> 150& Then
5130        .total_pcash_usd.BackColor = CLR_VLTRED
5140        .total_pcash_usd.Visible = True
5150        .total_pcash_usd_lbl.Visible = True
5160        .total_pcash_usd_lbl_line1.Visible = True
5170        .total_pcash_usd_lbl_line2.Visible = True
5180        strTmp01 = .Parent.total_cash_sum.Format
5190        intTmp02 = .Parent.total_cash_sum.DecimalPlaces
5200        .total_pcash.Format = strTmp01
5210        .total_pcash.DecimalPlaces = intTmp02
5220        .total_pcash.BackColor = CLR_VLTPRP
5230        If .Parent.total_cash_sum.FontName <> .total_pcash.FontName Then
5240          .total_pcash.FontName = .Parent.total_cash_sum.FontName
5250          .total_pcash.FontSize = .Parent.total_cash_sum.FontSize
5260        End If
5270      Else
5280        .total_pcash_usd.Visible = False
5290        .total_pcash_usd.Left = 0&
5300        .total_pcash_usd_lbl.Visible = False
5310        .total_pcash_usd_lbl_line1.Visible = False
5320        .total_pcash_usd_lbl_line2.Visible = False
5330        .total_pcash_usd_lbl.Left = 0&
5340        .total_pcash_usd_lbl_line1.Left = 0&
5350        .total_pcash_usd_lbl_line2.Left = 0&
5360        .total_pcash.BackColor = CLR_VLTRED
5370        .Width = .total_pcash_alt_line.Left
5380      End If
5390    End With

EXITP:
5400    Exit Sub

ERRH:
5410    Select Case ERR.Number
        Case Else
5420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5430    End Select
5440    Resume EXITP

End Sub

Public Function RecCnt() As Long

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

5510    lngRetVal = 0&

5520    With Me
5530      Set rst = .RecordsetClone
5540      With rst
5550        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
5560        Else
5570          .MoveLast
5580          lngRetVal = .RecordCount
5590        End If
5600        .Close
5610      End With
5620    End With

EXITP:
5630    Set rst = Nothing
5640    RecCnt = lngRetVal
5650    Exit Function

ERRH:
5660    lngRetVal = 0&
5670    Select Case ERR.Number
        Case Else
5680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5690    End Select
5700    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

5810    Select Case intWhere
        Case acCmdRecordsGoToFirst
5820      DoCmd.RunCommand acCmdRecordsGoToFirst
5830    Case acCmdRecordsGoToLast
5840      DoCmd.RunCommand acCmdRecordsGoToLast
5850    Case acCmdRecordsGoToPrevious
5860      DoCmd.RunCommand acCmdRecordsGoToPrevious
5870    Case acCmdRecordsGoToNext
5880      DoCmd.RunCommand acCmdRecordsGoToNext
5890    Case Else
5900      If IsMissing(varID) = False Then
5910        With Me
5920          Set rst = .RecordsetClone
5930          With rst
5940            .FindFirst "[accountno] = '" & varID & "'"
5950            If .NoMatch = False Then
5960              Me.Bookmark = .Bookmark
5970            End If
5980            .Close
5990          End With
6000        End With
6010      End If
6020    End Select

EXITP:
6030    Set rst = Nothing
6040    Exit Sub

ERRH:
6050    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
6060    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
6070    Case Else
6080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6090    End Select
6100    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [alphasort]"

6210    With Me
6220      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6230        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6240      End If
6250      .Controls(strSortLbl).Visible = False
6260      .Controls(strSortLine).Visible = False
6270      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
6280      If .total_pcash_usd.Visible = True Then
6290        If .total_pcash_usd_lbl.Left <> lngTmp01 Then
6300          .total_pcash_usd_lbl.Left = lngTmp01
6310          .total_pcash_usd_lbl_line1.Left = (.total_pcash_usd_lbl.Left + lngTpp)
6320          .total_pcash_usd_lbl_line2.Left = .total_pcash_usd_lbl_line1.Left
6330        End If
6340      End If
6350      If strProc = "Form_Load" Then
6360        strCalled = "accountno"
6370        strSortNow = strSortOrig
6380        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
6390        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
6400        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6410        .Controls(strSortLbl).Left = lngSortLbl_Left
6420        .Controls(strSortLine).Left = lngSortLine_Left
6430        .Controls(strSortLine).Width = lngSortLine_Width
6440        .Controls(strSortLbl).Caption = strArwUp
6450        .Controls(strSortLbl).ForeColor = CLR_DKBLU
6460      Else
6470        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
            'If strCalled <> "total_pcash_usd" Then
6480        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
6490        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
6500        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
            'Else
            'lngSortLbl_Left = ((.Controls(strCalled & "_lbl_box").Left + .Controls(strCalled & "_lbl_box").Width) - lngSortLbl_Width)
            'lngSortLine_Left = .Controls(strCalled & "_lbl_box").Left
            'lngSortLine_Width = (.Controls(strCalled & "_lbl_box").Width + lngTpp)
            '.Controls(strSortLbl).TextAlign = acTextAlignRight
            'End If
6510        .Controls(strSortLbl).Left = lngSortLbl_Left
6520        .Controls(strSortLine).Left = lngSortLine_Left
6530        .Controls(strSortLine).Width = lngSortLine_Width
6540        If strCalled = "accountno" Then
6550          strSortAsc = strSortOrig
6560          If strSortNow = strSortAsc Then
6570            strSortNow = "[alphasort] DESC"
6580            .Controls(strSortLbl).Caption = strArwDn
6590            .Controls(strSortLbl).ForeColor = CLR_DKRED
6600          Else
6610            strSortNow = strSortAsc
6620            .Controls(strSortLbl).Caption = strArwUp
6630            .Controls(strSortLbl).ForeColor = CLR_DKBLU
6640          End If
6650        Else
6660          .Controls(strSortLbl).Caption = strArwUp
6670          .Controls(strSortLbl).ForeColor = CLR_DKBLU
6680          Select Case strCalled
              Case "shortname"
6690            strSortAsc = "[shortname]"
6700          Case "total_shareface"
6710            strSortAsc = "[total_shareface]" & strStdAsc
6720          Case "total_pcash"
6730            strSortAsc = "[total_pcash]" & strStdAsc
6740          Case "total_pcash_usd"
6750            strSortAsc = "[total_pcash_usd]" & strStdAsc
6760          End Select
6770          If strSortNow = strSortAsc Then
6780            intCnt = CharCnt(strSortAsc, ",") + 1
6790            Select Case intCnt
                Case 1
6800              strTmp01 = strSortAsc & " DESC"
6810            Case 2
6820              intPos01 = InStr(strSortAsc, ",")
6830              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6840              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
6850            Case 3
6860              intPos01 = InStr(strSortAsc, ",")
6870              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6880              strTmp02 = Mid(strSortAsc, intPos01)
6890              intPos01 = InStr(2, strTmp02, ",")
6900              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
6910              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
6920            End Select
6930            strSortNow = strTmp01
6940            .Controls(strSortLbl).Caption = strArwDn
6950            .Controls(strSortLbl).ForeColor = CLR_DKRED
6960          Else
6970            strSortNow = strSortAsc
6980          End If
6990        End If
7000      End If
7010      .Controls(strSortLbl).Visible = True
7020      .Controls(strSortLine).Visible = True
7030      .OrderBy = strSortNow
7040      .OrderByOn = True
7050    End With

        ' ** Potential Sorts:
        ' **   [alphasort]
        ' **   [shortname]
        ' **   [total_shareface], [alphasort]
        ' **   [total_pcash], [alphasort]
        ' **   [total_pcash_usd], [alphasort]

EXITP:
7060    Exit Sub

ERRH:
7070    Select Case ERR.Number
        Case Else
7080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7090    End Select
7100    Resume EXITP

End Sub

Public Function SortNow_Get() As String

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

7210    strRetVal = strSortNow

EXITP:
7220    SortNow_Get = strRetVal
7230    Exit Function

ERRH:
7240    strRetVal = vbNullString
7250    Select Case ERR.Number
        Case Else
7260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7270    End Select
7280    Resume EXITP

End Function
