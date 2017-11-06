VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Reinvest_Rec_Detail_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Reinvest_Rec_Detail_Sub"

'VGC 07/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel on frmMap_Reinvest_Rec_Detail}
' **   Show Location:    L {chkShowLocation on frmMap_Reinvest_Rec_Detail}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Purchase on frmMap_Reinvest_Rec_Detail}
' **   OK:               O {cmdOK on frmMap_Reinvest_Rec_Detail}
' **   Print Report:     P {cmdPrintReport on frmMap_Reinvest_Rec_Detail}
' **   Tax Codes:        T {cmbTaxCodes_Purchase on frmMap_Reinvest_Rec_Detail}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description on frmMap_Reinvest_Rec_Detail}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots on frmMap_Reinvest_Rec_Detail}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[alphasort]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngPCashLbl_Left As Long
Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
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
150       If strCallingForm = "frmMap_Reinvest_Rec_Detail" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

180         lngPCashLbl_Left = .total_pcash_usd_lbl.Left

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         lngSortLbl_Width = .Sort_lbl.Width
210         .form_width_line.Left = 0&  ' ** Just get it out of the way.

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
490       .total_pcash.BackColor = CLR_DISABLED_BG
500       .total_pcash_usd.ForeColor = CLR_DISABLED_FG
510       .total_pcash_usd.BackColor = CLR_DISABLED_BG
520       .Location_ID.ForeColor = CLR_DISABLED_FG
530       .Location_ID.BackColor = CLR_DISABLED_BG
540       SortNow THIS_PROC  ' ** Procedure: Below
550       DoEvents
560     End With

EXITP:
570     Exit Sub

ERRH:
580     Select Case ERR.Number
        Case Else
590       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
600     End Select
610     Resume EXITP

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

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Comments:         M {description on frmMap_Reinvest_Rec_Detail}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc Totals:    F5 {RecalcTots on frmMap_Reinvest_Rec_Detail}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyUp
770         intRetVal = 0
780         MoveRec acCmdRecordsGoToPrevious
790       Case vbKeyDown
800         intRetVal = 0
810         MoveRec acCmdRecordsGoToNext
820       Case vbKeyEscape
830         With Me
840           intRetVal = 0
850           .Parent.cmdCancel_Click  ' ** Form Procedure: frmMap_Reinvest_Rec_Detail.
860         End With
870       Case vbKeyF5
880         With Me
890           intRetVal = 0
900           .Parent.RecalcTots  ' ** Form Procedure: frmMap_Reinvest_Rec_Detail.
910         End With
920       End Select
930     End If

        ' ** Ctrl keys.
940     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
950       Select Case intRetVal
          Case vbKeyM
960         With Me
970           intRetVal = 0
980           DoCmd.SelectObject acForm, .Parent.Name, False
990           .Parent.description.SetFocus
1000        End With
1010      Case vbKeyTab
1020        With Me
1030          intRetVal = 0
1040          DoCmd.SelectObject acForm, .Parent.Name, False
1050          .Parent.description.SetFocus
1060        End With
1070      Case vbKeyPageUp, vbKeyHome
1080        intRetVal = 0
1090        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1100      Case vbKeyPageDown, vbKeyEnd
1110        intRetVal = 0
1120        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1130      End Select
1140    End If

        ' ** Ctrl-Shift keys.
1150    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1160      Select Case intRetVal
          Case vbKeyF
1170        With Me
1180          intRetVal = 0
1190          DoCmd.SelectObject acForm, .Parent.Name, False
1200          .Parent.FocusHolder.SetFocus
1210        End With
1220      Case vbKeyTab
1230        With Me
1240          intRetVal = 0
1250          DoCmd.SelectObject acForm, .Parent.Name, False
1260          .Parent.cmdCancel.SetFocus
1270        End With
1280      End Select
1290    End If

EXITP:
1300    KeyCode = intRetVal
1310    Exit Sub

ERRH:
1320    intRetVal = 0
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Sub

Private Sub Form_Timer()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1410    With Me
1420      .TimerInterval = 0&
1430      blnCurrID = .Parent.curr_id.Visible
1440      lngCurrID = .Parent.curr_id
1450      IncludeCurrency  ' ** Procedure: Below.
1460      .Parent.IncludeCurrency True  ' ** Form Procedure: frmMap_Reinvest_Rec_Detail.
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

Private Sub Location_ID_lbl_DblClick(Cancel As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_lbl_DblClick"

5110    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub Location_ID_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_AfterUpdate"

5210    With Me
5220      Select Case IsNull(.Location_ID)
          Case True
5230        .Location_ID = 1&
5240      Case False
            'WHAT?
5250      End Select
5260      .loc_changed = True
5270      DoCmd.RunCommand acCmdSaveRecord
5280    End With

EXITP:
5290    Exit Sub

ERRH:
5300    Select Case ERR.Number
        Case Else
5310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5320    End Select
5330    Resume EXITP

End Sub

Private Sub Location_ID_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5420    intShiftDown = (Shift And acShiftMask) > 0
5430    intAltDown = (Shift And acAltMask) > 0
5440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5470        With Me
5480          intRetVal = 0
5490          lngRecsCur = RecCnt  ' ** Function: Below.
5500          If .CurrentRecord < lngRecsCur Then
5510            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5520            .FocusHolder.SetFocus
5530          Else
5540            DoCmd.SelectObject acForm, .Parent.Name, False
5550            .Parent.description.SetFocus
5560          End If
5570        End With
5580      End Select
5590    End If

        ' ** Shift keys.
5600    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5610      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5620        With Me
5630          intRetVal = 0
5640          If .CurrentRecord > 1 Then
5650            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5660            .FocusHolder.SetFocus
5670          Else
5680            DoCmd.SelectObject acForm, .Parent.Name, False
5690            .Parent.cmdCancel.SetFocus
5700          End If
5710        End With
5720      End Select
5730    End If

EXITP:
5740    KeyCode = intRetVal
5750    Exit Sub

ERRH:
5760    intRetVal = 0
5770    Select Case ERR.Number
        Case Else
5780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5790    End Select
5800    Resume EXITP

End Sub

Public Sub IncludeCurrency()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim strTmp01 As String, lngTmp02 As Long, intTmp03 As Integer

5910    With Me
5920      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5930        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5940      End If
5950      If blnCurrID = True And lngCurrID <> 150& Then
5960        lngTmp02 = ((.total_pcash_usd.Left + .total_pcash_usd.Width) + (4& * lngTpp))
5970        .Location_ID.Left = lngTmp02
5980        .Location_ID_lbl.Left = lngTmp02
5990        .Location_ID_lbl_line1.Left = (.Location_ID_lbl.Left + lngTpp)
6000        .Location_ID_lbl_line2.Left = (.Location_ID_lbl.Left + lngTpp)
6010        .total_pcash_usd.Visible = True
6020        .total_pcash_usd_lbl.Visible = True
6030        .total_pcash_usd_lbl_line1.Visible = True
6040        .total_pcash_usd_lbl_line1.Visible = True
6050        strTmp01 = .Parent.total_cash_sum.Format
6060        intTmp03 = .Parent.total_cash_sum.DecimalPlaces
6070        .total_pcash.Format = strTmp01
6080        .total_pcash.DecimalPlaces = intTmp03
6090        .total_pcash.BackColor = CLR_LTGRN
            'If .Parent.total_cash_sum.FontName <> .total_pcash.FontName Then
            '  .total_pcash.FontName = .Parent.total_cash_sum.FontName
            '  .total_pcash.FontSize = .Parent.total_cash_sum.FontSize
            'End If
6100      Else
6110        .Location_ID.Left = .total_pcash_usd.Left
6120        .Location_ID_lbl.Left = .Location_ID.Left
6130        .Location_ID_lbl_line1.Left = (.Location_ID_lbl.Left + lngTpp)
6140        .Location_ID_lbl_line2.Left = (.Location_ID_lbl.Left + lngTpp)
6150        .total_pcash_usd.Visible = False
6160        .total_pcash_usd.Left = 0&
6170        .total_pcash_usd_lbl.Visible = False
6180        .total_pcash_usd_lbl_line1.Visible = False
6190        .total_pcash_usd_lbl_line2.Visible = False
6200        .total_pcash_usd_lbl.Left = 0&
6210        .total_pcash_usd_lbl_line1.Left = 0&
6220        .total_pcash_usd_lbl_line2.Left = 0&
6230        .total_cash_sum_usd.Left = 0&
6240      End If
6250    End With

EXITP:
6260    Exit Sub

ERRH:
6270    Select Case ERR.Number
        Case Else
6280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6290    End Select
6300    Resume EXITP

End Sub

Public Sub ShowLocation(blnShow As Boolean)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowLocation"

        Dim lngTmp01 As Long

6410    With Me
6420      Select Case blnShow
          Case True
6430        lngTmp01 = ((.total_pcash.Left + .total_pcash.Width) + (4& * lngTpp))
6440        lngTmp01 = ((lngTmp01 + .Location_ID.Width) + (4& * lngTpp))
6450        If blnCurrID = True And lngCurrID <> 150& Then
6460          lngTmp01 = ((lngTmp01 + .total_pcash_usd.Width) + (4& * lngTpp))
6470        End If
6480        .Width = lngTmp01
6490        lngTmp01 = ((.total_pcash.Left + .total_pcash.Width) + (4& * lngTpp))
6500        If blnCurrID = True And lngCurrID <> 150& Then
6510          lngTmp01 = ((lngTmp01 + .total_pcash_usd.Width) + (4& * lngTpp))
6520        End If
6530        .Location_ID.Left = lngTmp01
6540        .Location_ID_lbl.Left = lngTmp01
6550        .Location_ID_lbl_line1.Left = (.Location_ID_lbl.Left + lngTpp)
6560        .Location_ID_lbl_line2.Left = (.Location_ID_lbl.Left + lngTpp)
6570        .Location_ID.Visible = True
6580        .Location_ID_lbl.Visible = True
6590        .Location_ID_lbl_line1.Visible = True
6600        .Location_ID_lbl_line2.Visible = True
6610      Case False
6620        .Location_ID.Visible = False
6630        .Location_ID.Left = 0&
6640        .Location_ID_lbl.Visible = False
6650        .Location_ID_lbl_line1.Visible = False
6660        .Location_ID_lbl_line2.Visible = False
6670        .Location_ID_lbl.Left = 0&
6680        .Location_ID_lbl_line1.Left = 0&
6690        .Location_ID_lbl_line2.Left = 0&
6700        lngTmp01 = ((.total_pcash.Left + .total_pcash.Width) + (4& * lngTpp))
6710        If blnCurrID = True And lngCurrID <> 150& Then
6720          lngTmp01 = ((lngTmp01 + .total_pcash_usd.Width) + (4& * lngTpp))
6730        End If
6740        .Width = lngTmp01
6750      End Select
6760    End With

EXITP:
6770    Exit Sub

ERRH:
6780    Select Case ERR.Number
        Case Else
6790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6800    End Select
6810    Resume EXITP

End Sub

Public Function RecCnt() As Long

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

6910    lngRetVal = 0&

6920    With Me
6930      Set rst = .RecordsetClone
6940      With rst
6950        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
6960        Else
6970          .MoveLast
6980          lngRetVal = .RecordCount
6990        End If
7000        .Close
7010      End With
7020    End With

EXITP:
7030    Set rst = Nothing
7040    RecCnt = lngRetVal
7050    Exit Function

ERRH:
7060    lngRetVal = 0&
7070    Select Case ERR.Number
        Case Else
7080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7090    End Select
7100    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

7210    Select Case intWhere
        Case acCmdRecordsGoToFirst
7220      DoCmd.RunCommand acCmdRecordsGoToFirst
7230    Case acCmdRecordsGoToLast
7240      DoCmd.RunCommand acCmdRecordsGoToLast
7250    Case acCmdRecordsGoToPrevious
7260      DoCmd.RunCommand acCmdRecordsGoToPrevious
7270    Case acCmdRecordsGoToNext
7280      DoCmd.RunCommand acCmdRecordsGoToNext
7290    Case Else
7300      If IsMissing(varID) = False Then
7310        With Me
7320          Set rst = .RecordsetClone
7330          With rst
7340            .FindFirst "[accountno] = '" & varID & "'"
7350            If .NoMatch = False Then
7360              Me.Bookmark = .Bookmark
7370            End If
7380            .Close
7390          End With
7400        End With
7410      End If
7420    End Select

EXITP:
7430    Set rst = Nothing
7440    Exit Sub

ERRH:
7450    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7460    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7470    Case Else
7480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7490    End Select
7500    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [alphasort]"

7610    With Me
7620      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7630        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7640      End If
7650      .Controls(strSortLbl).Visible = False
7660      .Controls(strSortLine).Visible = False
7670      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
7680      If blnCurrID = True And lngCurrID <> 150& Then
7690        .total_pcash_usd_lbl.Left = lngPCashLbl_Left
7700      End If
7710      If .Parent.chkShowLocation = True Then
7720        .Location_ID_lbl.Left = .Location_ID.Left
7730      End If
7740      If strProc = "Form_Load" Then
7750        strCalled = "accountno"
7760        strSortNow = strSortOrig
7770        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7780        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7790        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7800        .Controls(strSortLbl).Left = lngSortLbl_Left
7810        .Controls(strSortLine).Left = lngSortLine_Left
7820        .Controls(strSortLine).Width = lngSortLine_Width
7830        .Controls(strSortLbl).Caption = strArwUp
7840        .Controls(strSortLbl).ForeColor = CLR_DKBLU
7850      Else
7860        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
7870        If strCalled = "total_pcash_usd" Then
7880          lngSortLbl_Left = ((.Controls(strCalled).Left + .Controls(strCalled).Width) - lngTpp)
7890          lngSortLine_Left = .Controls(strCalled).Left
7900          lngSortLine_Width = (.Controls(strCalled).Width + lngTpp)
7910        Else
7920          lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7930          lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7940          lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7950        End If
7960        .Controls(strSortLbl).Left = lngSortLbl_Left
7970        .Controls(strSortLine).Left = lngSortLine_Left
7980        .Controls(strSortLine).Width = lngSortLine_Width
7990        If strCalled = "accountno" Then
8000          strSortAsc = strSortOrig
8010          If strSortNow = strSortAsc Then
8020            strSortNow = "[alphasort] DESC"
8030            .Controls(strSortLbl).Caption = strArwDn
8040            .Controls(strSortLbl).ForeColor = CLR_DKRED
8050          Else
8060            strSortNow = strSortAsc
8070            .Controls(strSortLbl).Caption = strArwUp
8080            .Controls(strSortLbl).ForeColor = CLR_DKBLU
8090          End If
8100        Else
8110          .Controls(strSortLbl).Caption = strArwUp
8120          .Controls(strSortLbl).ForeColor = CLR_DKBLU
8130          Select Case strCalled
              Case "shortname"
8140            strSortAsc = "[shortname]"
8150          Case "total_shareface"
8160            strSortAsc = "[total_shareface]" & strStdAsc
8170          Case "total_pcash"

8180            strSortAsc = "[total_pcash]" & strStdAsc
                'If blnCurrID = True And lngCurrID <> 150& Then
                '  .total_pcash_usd_lbl.Left = (lngPCashLbl_Left + (3& * lngTpp))
                'ElseIf .Parent.chkShowLocation = True Then
                '  .Location_ID_lbl.Left = (.Location_ID.Left + (4& * lngTpp))
                'End If
8190          Case "total_pcash_usd"
8200            strSortAsc = "[total_pcash_usd]" & strStdAsc
                'If .Parent.chkShowLocation = True Then
                '  .Location_ID_lbl.Left = (.Location_ID.Left + (4& * lngTpp))
                'End If
8210          Case "Location_ID"
8220            strSortAsc = "[Location_ID]" & strStdAsc
8230          End Select
8240          If strSortNow = strSortAsc Then
8250            intCnt = CharCnt(strSortAsc, ",") + 1
8260            Select Case intCnt
                Case 1
8270              strTmp01 = strSortAsc & " DESC"
8280            Case 2
8290              intPos01 = InStr(strSortAsc, ",")
8300              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8310              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
8320            Case 3
8330              intPos01 = InStr(strSortAsc, ",")
8340              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8350              strTmp02 = Mid(strSortAsc, intPos01)
8360              intPos01 = InStr(2, strTmp02, ",")
8370              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8380              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
8390            End Select
8400            strSortNow = strTmp01
8410            .Controls(strSortLbl).Caption = strArwDn
8420            .Controls(strSortLbl).ForeColor = CLR_DKRED
8430          Else
8440            strSortNow = strSortAsc
8450          End If
8460        End If
8470      End If
8480      .Controls(strSortLbl).Visible = True
8490      .Controls(strSortLine).Visible = True
8500      .OrderBy = strSortNow
8510      .OrderByOn = True
8520    End With

        ' ** Potential Sorts:
        ' **   [alphasort]
        ' **   [shortname]
        ' **   [total_shareface], [alphasort]
        ' **   [total_pcash], [alphasort]
        ' **   [total_pcash_usd], [alphasort]
        ' **   [Location_ID], [alphasort]

EXITP:
8530    Exit Sub

ERRH:
8540    Select Case ERR.Number
        Case Else
8550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8560    End Select
8570    Resume EXITP

End Sub

Public Function SortNow_Get() As String

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

8610    strRetVal = strSortNow

EXITP:
8620    SortNow_Get = strRetVal
8630    Exit Function

ERRH:
8640    strRetVal = vbNullString
8650    Select Case ERR.Number
        Case Else
8660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8670    End Select
8680    Resume EXITP

End Function
