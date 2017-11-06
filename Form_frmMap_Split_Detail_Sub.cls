VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Split_Detail_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Split_Detail_Sub"

'VGC 07/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel on frmMap_Split_Detail}
' **   Show Location:    L {chkShowLocation on frmMap_Split_Detail}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Deposit on frmMap_Split_Detail}
' **   OK:               O {cmdOK on frmMap_Split_Detail}
' **   Print Report:     P {cmdPrintReport on frmMap_Split_Detail}
' **   Tax Codes:        T {cmbTaxCodes_Deposit on frmMap_Split_Detail}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description on frmMap_Split_Detail}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots on frmMap_Split_Detail}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[alphasort]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

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
150       If strCallingForm = "frmMap_Split_Detail" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         lngSortLbl_Width = .Sort_lbl.Width
200         .form_width_line.Left = 0&  ' ** Just get it out of the way.

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
440       .shortname.ForeColor = CLR_DISABLED_FG
450       .shortname.BackColor = CLR_DISABLED_BG
460       .total_shareface.ForeColor = CLR_DISABLED_FG
470       .total_shareface.BackColor = CLR_DISABLED_BG
480       .total_icash.ForeColor = CLR_DISABLED_FG
490       .total_icash.BackColor = CLR_DISABLED_BG
500       .total_icash_rev.ForeColor = CLR_DISABLED_FG
510       .total_icash_rev.BackColor = CLR_DISABLED_BG
520       .Location_ID.ForeColor = CLR_DISABLED_FG
530       .Location_ID.BackColor = CLR_DISABLED_BG
540       SortNow THIS_PROC  ' ** Procedure: Below
550     End With

EXITP:
560     Exit Sub

ERRH:
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

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
        ' **   Comments:         M {description on frmMap_Split_Detail}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc Totals:    F5 {RecalcTots on frmMap_Split_Detail}

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
850           .Parent.cmdCancel_Click  ' ** Form Procedure: frmMap_Split_Detail.
860         End With
870       Case vbKeyF5
880         With Me
890           intRetVal = 0
900           .Parent.RecalcTots  ' ** Form Function: frmMap_Split_Detail.
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

1430    End With

EXITP:
1440    Exit Sub

ERRH:
1450    Select Case ERR.Number
        Case Else
1460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1470    End Select
1480    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1520    intShiftDown = (Shift And acShiftMask) > 0
1530    intAltDown = (Shift And acAltMask) > 0
1540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1570        With Me
1580          intRetVal = 0
1590          lngRecsCur = RecCnt  ' ** Function: Below.
1600          If .CurrentRecord < lngRecsCur Then
1610            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1620            .FocusHolder.SetFocus
1630          Else
1640            DoCmd.SelectObject acForm, .Parent.Name, False
1650            .Parent.description.SetFocus
1660          End If
1670        End With
1680      End Select
1690    End If

        ' ** Shift keys.
1700    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1710      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
1720        With Me
1730          intRetVal = 0
1740          If .CurrentRecord > 1 Then
1750            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1760            .FocusHolder.SetFocus
1770          Else
1780            DoCmd.SelectObject acForm, .Parent.Name, False
1790            .Parent.cmdCancel.SetFocus
1800          End If
1810        End With
1820      End Select
1830    End If

EXITP:
1840    KeyCode = intRetVal
1850    Exit Sub

ERRH:
1860    intRetVal = 0
1870    Select Case ERR.Number
        Case Else
1880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1890    End Select
1900    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

2010    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2170        With Me
2180          intRetVal = 0
2190          lngRecsCur = RecCnt  ' ** Function: Below.
2200          If .CurrentRecord < lngRecsCur Then
2210            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2220            .FocusHolder.SetFocus
2230          Else
2240            DoCmd.SelectObject acForm, .Parent.Name, False
2250            .Parent.description.SetFocus
2260          End If
2270        End With
2280      End Select
2290    End If

        ' ** Shift keys.
2300    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2310      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2320        With Me
2330          intRetVal = 0
2340          If .CurrentRecord > 1 Then
2350            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2360            .FocusHolder.SetFocus
2370          Else
2380            DoCmd.SelectObject acForm, .Parent.Name, False
2390            .Parent.cmdCancel.SetFocus
2400          End If
2410        End With
2420      End Select
2430    End If

EXITP:
2440    KeyCode = intRetVal
2450    Exit Sub

ERRH:
2460    intRetVal = 0
2470    Select Case ERR.Number
        Case Else
2480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2490    End Select
2500    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

2610    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub shortname_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2770        With Me
2780          intRetVal = 0
2790          lngRecsCur = RecCnt  ' ** Function: Below.
2800          If .CurrentRecord < lngRecsCur Then
2810            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
2820            .FocusHolder.SetFocus
2830          Else
2840            DoCmd.SelectObject acForm, .Parent.Name, False
2850            .Parent.description.SetFocus
2860          End If
2870        End With
2880      End Select
2890    End If

        ' ** Shift keys.
2900    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2910      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2920        With Me
2930          intRetVal = 0
2940          If .CurrentRecord > 1 Then
2950            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2960            .FocusHolder.SetFocus
2970          Else
2980            DoCmd.SelectObject acForm, .Parent.Name, False
2990            .Parent.cmdCancel.SetFocus
3000          End If
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

Private Sub total_shareface_lbl_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_lbl_DblClick"

3210    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub total_shareface_KeyDown(KeyCode As Integer, Shift As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_KeyDown"

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
3390          lngRecsCur = RecCnt  ' ** Function: Below.
3400          If .CurrentRecord < lngRecsCur Then
3410            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3420            .FocusHolder.SetFocus
3430          Else
3440            DoCmd.SelectObject acForm, .Parent.Name, False
3450            .Parent.description.SetFocus
3460          End If
3470        End With
3480      End Select
3490    End If

        ' ** Shift keys.
3500    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3510      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3520        With Me
3530          intRetVal = 0
3540          If .CurrentRecord > 1 Then
3550            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3560            .FocusHolder.SetFocus
3570          Else
3580            DoCmd.SelectObject acForm, .Parent.Name, False
3590            .Parent.cmdCancel.SetFocus
3600          End If
3610        End With
3620      End Select
3630    End If

EXITP:
3640    KeyCode = intRetVal
3650    Exit Sub

ERRH:
3660    intRetVal = 0
3670    Select Case ERR.Number
        Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Sub

Private Sub total_icash_lbl_DblClick(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "total_icash_lbl_DblClick"

3810    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub total_icash_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "total_icash_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3920    intShiftDown = (Shift And acShiftMask) > 0
3930    intAltDown = (Shift And acAltMask) > 0
3940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3970        With Me
3980          intRetVal = 0
3990          lngRecsCur = RecCnt  ' ** Function: Below.
4000          If .CurrentRecord < lngRecsCur Then
4010            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4020            .FocusHolder.SetFocus
4030          Else
4040            DoCmd.SelectObject acForm, .Parent.Name, False
4050            .Parent.description.SetFocus
4060          End If
4070        End With
4080      End Select
4090    End If

        ' ** Shift keys.
4100    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4110      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4120        With Me
4130          intRetVal = 0
4140          If .CurrentRecord > 1 Then
4150            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4160            .FocusHolder.SetFocus
4170          Else
4180            DoCmd.SelectObject acForm, .Parent.Name, False
4190            .Parent.cmdCancel.SetFocus
4200          End If
4210        End With
4220      End Select
4230    End If

EXITP:
4240    KeyCode = intRetVal
4250    Exit Sub

ERRH:
4260    intRetVal = 0
4270    Select Case ERR.Number
        Case Else
4280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4290    End Select
4300    Resume EXITP

End Sub

Private Sub total_icash_rev_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "total_icash_rev_KeyDown"

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
4520            .FocusHolder.SetFocus
4530          Else
4540            DoCmd.SelectObject acForm, .Parent.Name, False
4550            .Parent.description.SetFocus
4560          End If
4570        End With
4580      End Select
4590    End If

        ' ** Shift keys.
4600    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4610      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4620        With Me
4630          intRetVal = 0
4640          If .CurrentRecord > 1 Then
4650            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4660            .FocusHolder.SetFocus
4670          Else
4680            DoCmd.SelectObject acForm, .Parent.Name, False
4690            .Parent.cmdCancel.SetFocus
4700          End If
4710        End With
4720      End Select
4730    End If

EXITP:
4740    KeyCode = intRetVal
4750    Exit Sub

ERRH:
4760    intRetVal = 0
4770    Select Case ERR.Number
        Case Else
4780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4790    End Select
4800    Resume EXITP

End Sub

Private Sub Location_ID_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_lbl_DblClick"

4910    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub Location_ID_KeyDown(KeyCode As Integer, Shift As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5020    intShiftDown = (Shift And acShiftMask) > 0
5030    intAltDown = (Shift And acAltMask) > 0
5040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5070        With Me
5080          intRetVal = 0
5090          lngRecsCur = RecCnt  ' ** Function: Below.
5100          If .CurrentRecord < lngRecsCur Then
5110            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
5120            .FocusHolder.SetFocus
5130          Else
5140            DoCmd.SelectObject acForm, .Parent.Name, False
5150            .Parent.description.SetFocus
5160          End If
5170        End With
5180      End Select
5190    End If

        ' ** Shift keys.
5200    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5210      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5220        With Me
5230          intRetVal = 0
5240          If .CurrentRecord > 1 Then
5250            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5260            .FocusHolder.SetFocus
5270          Else
5280            DoCmd.SelectObject acForm, .Parent.Name, False
5290            .Parent.cmdCancel.SetFocus
5300          End If
5310        End With
5320      End Select
5330    End If

EXITP:
5340    KeyCode = intRetVal
5350    Exit Sub

ERRH:
5360    intRetVal = 0
5370    Select Case ERR.Number
        Case Else
5380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5390    End Select
5400    Resume EXITP

End Sub

Private Sub IncludeCurrency()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

5510    With Me
5520      If blnCurrID = True And lngCurrID <> 150& Then
            ' ** Nothing to do down here.
5530      Else
            ' ** Nope.
5540      End If
5550    End With

EXITP:
5560    Exit Sub

ERRH:
5570    Select Case ERR.Number
        Case Else
5580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5590    End Select
5600    Resume EXITP

End Sub

Public Sub ShowLocation(blnShow As Boolean)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowLocation"

        Dim lngTmp01 As Long

5710    With Me
5720      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5730        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5740      End If
5750      Select Case blnShow
          Case True
5760        lngTmp01 = (.Location_ID.Width + (4& * lngTpp))
5770        .Width = lngTmp01
5780        lngTmp01 = ((.total_icash.Left + .total_icash.Width) + (4& * lngTpp))
5790        .Location_ID.Left = lngTmp01
5800        .Location_ID_lbl.Left = lngTmp01
5810        .Location_ID_lbl_line1.Left = (.Location_ID_lbl.Left + lngTpp)
5820        .Location_ID_lbl_line1.Visible = True
5830        .Location_ID_lbl_line2.Left = .Location_ID_lbl_line1.Left
5840        .Location_ID_lbl_line2.Visible = True
5850        .Location_ID.Visible = True
5860        .Location_ID_lbl.Visible = True
5870      Case False
5880        .Location_ID.Visible = False
5890        .Location_ID.Left = 0&
5900        .Location_ID_lbl.Visible = False
5910        .Location_ID_lbl.Left = 0&
5920        .Location_ID_lbl_line1.Left = 0&
5930        .Location_ID_lbl_line1.Visible = False
5940        .Location_ID_lbl_line2.Left = 0&
5950        .Location_ID_lbl_line2.Visible = False
5960        lngTmp01 = ((.total_icash.Left + .total_icash.Width) + (4& * lngTpp))
5970        .Width = lngTmp01
5980      End Select
5990    End With

EXITP:
6000    Exit Sub

ERRH:
6010    Select Case ERR.Number
        Case Else
6020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6030    End Select
6040    Resume EXITP

End Sub

Public Function RecCnt() As Long

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

6110    lngRetVal = 0&

6120    With Me
6130      Set rst = .RecordsetClone
6140      With rst
6150        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
6160        Else
6170          .MoveLast
6180          lngRetVal = .RecordCount
6190        End If
6200        .Close
6210      End With
6220    End With

EXITP:
6230    Set rst = Nothing
6240    RecCnt = lngRetVal
6250    Exit Function

ERRH:
6260    lngRetVal = 0&
6270    Select Case ERR.Number
        Case Else
6280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6290    End Select
6300    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

6410    Select Case intWhere
        Case acCmdRecordsGoToFirst
6420      DoCmd.RunCommand acCmdRecordsGoToFirst
6430    Case acCmdRecordsGoToLast
6440      DoCmd.RunCommand acCmdRecordsGoToLast
6450    Case acCmdRecordsGoToPrevious
6460      DoCmd.RunCommand acCmdRecordsGoToPrevious
6470    Case acCmdRecordsGoToNext
6480      DoCmd.RunCommand acCmdRecordsGoToNext
6490    Case Else
6500      If IsMissing(varID) = False Then
6510        With Me
6520          Set rst = .RecordsetClone
6530          With rst
6540            .FindFirst "[accountno] = '" & varID & "'"
6550            If .NoMatch = False Then
6560              Me.Bookmark = .Bookmark
6570            End If
6580            .Close
6590          End With
6600        End With
6610      End If
6620    End Select

EXITP:
6630    Set rst = Nothing
6640    Exit Sub

ERRH:
6650    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
6660    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
6670    Case Else
6680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6690    End Select
6700    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [alphasort]"

6810    With Me
6820      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6830        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6840      End If
6850      .Controls(strSortLbl).Visible = False
6860      .Controls(strSortLine).Visible = False
6870      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
6880      If .Location_ID_lbl.Left <> .Location_ID.Left Then
6890        .Location_ID_lbl.Left = .Location_ID.Left
6900      End If
6910      If strProc = "Form_Load" Then
6920        strCalled = "accountno"
6930        strSortNow = strSortOrig
6940        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
6950        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
6960        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6970        .Controls(strSortLbl).Left = lngSortLbl_Left
6980        .Controls(strSortLine).Left = lngSortLine_Left
6990        .Controls(strSortLine).Width = lngSortLine_Width
7000        .Controls(strSortLbl).Caption = strArwUp
7010        .Controls(strSortLbl).ForeColor = CLR_DKBLU
7020      Else
7030        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
7040        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7050        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7060        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7070        .Controls(strSortLbl).Left = lngSortLbl_Left
7080        .Controls(strSortLine).Left = lngSortLine_Left
7090        .Controls(strSortLine).Width = lngSortLine_Width
7100        If strCalled = "accountno" Then
7110          strSortAsc = strSortOrig
7120          If strSortNow = strSortAsc Then
7130            strSortNow = "[alphasort] DESC"
7140            .Controls(strSortLbl).Caption = strArwDn
7150            .Controls(strSortLbl).ForeColor = CLR_DKRED
7160          Else
7170            strSortNow = strSortAsc
7180            .Controls(strSortLbl).Caption = strArwUp
7190            .Controls(strSortLbl).ForeColor = CLR_DKBLU
7200          End If
7210        Else
7220          .Controls(strSortLbl).Caption = strArwUp
7230          .Controls(strSortLbl).ForeColor = CLR_DKBLU
7240          Select Case strCalled
              Case "shortname"
7250            strSortAsc = "[shortname]"
7260          Case "total_shareface"
7270            strSortAsc = "[total_shareface]" & strStdAsc
7280          Case "total_icash", "total_icash_rev"
7290            If .total_icash_rev.Visible = True Then
7300              strSortAsc = "[total_icash_rev]" & strStdAsc
7310            Else
7320              strSortAsc = "[total_icash]" & strStdAsc
7330            End If
                '.Location_ID_lbl.Left = (.Location_ID.Left + (5& * lngTpp))
7340          Case "Location_ID"
7350            strSortAsc = "[Location_IDx]" & strStdAsc
7360          End Select
7370          If strSortNow = strSortAsc Then
7380            intCnt = CharCnt(strSortAsc, ",") + 1
7390            Select Case intCnt
                Case 1
7400              strTmp01 = strSortAsc & " DESC"
7410            Case 2
7420              intPos01 = InStr(strSortAsc, ",")
7430              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
7440              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
7450            Case 3
7460              intPos01 = InStr(strSortAsc, ",")
7470              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
7480              strTmp02 = Mid(strSortAsc, intPos01)
7490              intPos01 = InStr(2, strTmp02, ",")
7500              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
7510              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
7520            End Select
7530            strSortNow = strTmp01
7540            .Controls(strSortLbl).Caption = strArwDn
7550            .Controls(strSortLbl).ForeColor = CLR_DKRED
7560          Else
7570            strSortNow = strSortAsc
7580          End If
7590        End If
7600      End If
7610      .Controls(strSortLbl).Visible = True
7620      .Controls(strSortLine).Visible = True
7630      .OrderBy = strSortNow
7640      .OrderByOn = True
7650    End With

        ' ** Potential Sorts:
        ' **   [alphasort]
        ' **   [shortname]
        ' **   [total_shareface], [alphasort]
        ' **   [total_icash], [alphasort]
        ' **   [Location_IDx], [alphasort]

EXITP:
7660    Exit Sub

ERRH:
7670    Select Case ERR.Number
        Case Else
7680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7690    End Select
7700    Resume EXITP

End Sub

Public Function SortNow_Get() As String

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

7810    strRetVal = strSortNow

EXITP:
7820    SortNow_Get = strRetVal
7830    Exit Function

ERRH:
7840    strRetVal = vbNullString
7850    Select Case ERR.Number
        Case Else
7860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7870    End Select
7880    Resume EXITP

End Function
