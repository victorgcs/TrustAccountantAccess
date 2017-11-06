VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Int_Detail_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Int_Detail_Sub"

'VGC 07/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel on frmMap_Int_Detail}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Interest on frmMap_Int_Detail}
' **   OK:               O {cmdOK on frmMap_Int_Detail}
' **   Print Report:     P {cmdPrintReport on frmMap_Int_Detail}
' **   Reinvest:         R {cmdReinvest on frmMap_Int_Detail}
' **   Tax Codes:        T {cmbTaxCodes_Interest on frmMap_Int_Detail}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description on frmMap_Int_Detail}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots on frmMap_Int_Detail}

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
150       If strCallingForm = "frmMap_Int_Detail" Then

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
500       .total_icash_usd.ForeColor = CLR_DISABLED_FG
510       .total_icash_usd.BackColor = CLR_DISABLED_BG
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

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Comments:         M {description on frmMap_Int_Detail}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc Totals:    F5 {RecalcTots on frmMap_Int_Detail}

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
750           .Parent.cmdCancel_Click  ' ** Form Procedure: frmMap_Int_Detail.
760         End With
770       Case vbKeyF5
780         With Me
790           intRetVal = 0
800           .Parent.RecalcTots  ' ** Form Function: frmMap_Int_Detail.
810         End With
820       End Select
830     End If

        ' ** Ctrl keys.
840     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
850       Select Case intRetVal
          Case vbKeyM
860         With Me
870           intRetVal = 0
880           DoCmd.SelectObject acForm, .Parent.Name, False
890           .Parent.description.SetFocus
900         End With
910       Case vbKeyTab
920         With Me
930           intRetVal = 0
940           DoCmd.SelectObject acForm, .Parent.Name, False
950           .Parent.description.SetFocus
960         End With
970       Case vbKeyPageUp, vbKeyHome
980         intRetVal = 0
990         MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1000      Case vbKeyPageDown, vbKeyEnd
1010        intRetVal = 0
1020        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1030      End Select
1040    End If

        ' ** Ctrl-Shift keys.
1050    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1060      Select Case intRetVal
          Case vbKeyF
1070        With Me
1080          intRetVal = 0
1090          DoCmd.SelectObject acForm, .Parent.Name, False
1100          .Parent.FocusHolder.SetFocus
1110        End With
1120      Case vbKeyTab
1130        With Me
1140          intRetVal = 0
1150          DoCmd.SelectObject acForm, .Parent.Name, False
1160          .Parent.cmdCancel.SetFocus
1170        End With
1180      End Select
1190    End If

EXITP:
1200    KeyCode = intRetVal
1210    Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1310    With Me
1320      .TimerInterval = 0&
1330      blnCurrID = .Parent.curr_id.Visible
1340      lngCurrID = .Parent.curr_id
1350      IncludeCurrency  ' ** Procedure: Below.
1360      .Parent.IncludeCurrency True  ' ** Form Procedure: frmMap_Int_Detail.
1370    End With

EXITP:
1380    Exit Sub

ERRH:
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

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

Private Sub total_icash_usd_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "total_icash_usd_lbl_DblClick"

4410    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub total_icash_usd_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "total_icash_usd_KeyDown"

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
4620            .FocusHolder.SetFocus
4630          Else
4640            DoCmd.SelectObject acForm, .Parent.Name, False
4650            .Parent.description.SetFocus
4660          End If
4670        End With
4680      End Select
4690    End If

        ' ** Shift keys.
4700    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4710      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4720        With Me
4730          intRetVal = 0
4740          If .CurrentRecord > 1 Then
4750            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4760            .FocusHolder.SetFocus
4770          Else
4780            DoCmd.SelectObject acForm, .Parent.Name, False
4790            .Parent.cmdCancel.SetFocus
4800          End If
4810        End With
4820      End Select
4830    End If

EXITP:
4840    KeyCode = intRetVal
4850    Exit Sub

ERRH:
4860    intRetVal = 0
4870    Select Case ERR.Number
        Case Else
4880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4890    End Select
4900    Resume EXITP

End Sub

Public Sub IncludeCurrency()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim strTmp01 As String, intTmp02 As Integer

5010    With Me
5020      If blnCurrID = True And lngCurrID <> 150& Then
5030        .total_icash_usd.Visible = True
5040        .total_icash_usd_lbl.Visible = True
5050        .total_icash_usd_lbl_line1.Visible = True
5060        .total_icash_usd_lbl_line2.Visible = True
5070        strTmp01 = .Parent.total_cash_sum.Format
5080        intTmp02 = .Parent.total_cash_sum.DecimalPlaces
5090        .total_icash.Format = strTmp01
5100        .total_icash.DecimalPlaces = intTmp02
5110        .total_icash.BackColor = CLR_LTGRN
5120        If .Parent.total_cash_sum.FontName <> .total_icash.FontName Then
5130          .total_icash.FontName = .Parent.total_cash_sum.FontName
5140          .total_icash.FontSize = .Parent.total_cash_sum.FontSize
5150        End If
5160      Else
5170        .total_icash_usd.Visible = False
5180        .total_icash_usd.Left = 0&
5190        .total_icash_usd_lbl.Visible = False
5200        .total_icash_usd_lbl.Left = 0&
5210        .total_icash_usd_lbl_line1.Visible = False
5220        .total_icash_usd_lbl_line2.Visible = False
5230        .Width = .total_icash_alt_line.Left
5240      End If
5250    End With

EXITP:
5260    Exit Sub

ERRH:
5270    Select Case ERR.Number
        Case Else
5280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5290    End Select
5300    Resume EXITP

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
5840            .FindFirst "[accountno] = '" & varID & "'"
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

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [alphasort]"

6110    With Me
6120      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6130        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
6140      End If
6150      .Controls(strSortLbl).Visible = False
6160      .Controls(strSortLine).Visible = False
6170      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
6180      If strProc = "Form_Load" Then
6190        strCalled = "accountno"
6200        strSortNow = strSortOrig
6210        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
6220        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
6230        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6240        .Controls(strSortLbl).Left = lngSortLbl_Left
6250        .Controls(strSortLine).Left = lngSortLine_Left
6260        .Controls(strSortLine).Width = lngSortLine_Width
6270        .Controls(strSortLbl).Caption = strArwUp
6280        .Controls(strSortLbl).ForeColor = CLR_DKBLU
6290      Else
6300        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
6310        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
6320        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
6330        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
6340        .Controls(strSortLbl).Left = lngSortLbl_Left
6350        .Controls(strSortLine).Left = lngSortLine_Left
6360        .Controls(strSortLine).Width = lngSortLine_Width
6370        If strCalled = "accountno" Then
6380          strSortAsc = strSortOrig
6390          If strSortNow = strSortAsc Then
6400            strSortNow = "[alphasort] DESC"
6410            .Controls(strSortLbl).Caption = strArwDn
6420            .Controls(strSortLbl).ForeColor = CLR_DKRED
6430          Else
6440            strSortNow = strSortAsc
6450            .Controls(strSortLbl).Caption = strArwUp
6460            .Controls(strSortLbl).ForeColor = CLR_DKBLU
6470          End If
6480        Else
6490          .Controls(strSortLbl).Caption = strArwUp
6500          .Controls(strSortLbl).ForeColor = CLR_DKBLU
6510          Select Case strCalled
              Case "shortname"
6520            strSortAsc = "[shortname]"
6530          Case "total_shareface"
6540            strSortAsc = "[total_shareface]" & strStdAsc
6550          Case "total_icash"
6560            strSortAsc = "[total_icash]" & strStdAsc
6570          Case "total_icash_usd"
6580            strSortAsc = "[total_icash_usd]" & strStdAsc
6590          End Select
6600          If strSortNow = strSortAsc Then
6610            intCnt = CharCnt(strSortAsc, ",") + 1
6620            Select Case intCnt
                Case 1
6630              strTmp01 = strSortAsc & " DESC"
6640            Case 2
6650              intPos01 = InStr(strSortAsc, ",")
6660              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6670              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
6680            Case 3
6690              intPos01 = InStr(strSortAsc, ",")
6700              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
6710              strTmp02 = Mid(strSortAsc, intPos01)
6720              intPos01 = InStr(2, strTmp02, ",")
6730              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
6740              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
6750            End Select
6760            strSortNow = strTmp01
6770            .Controls(strSortLbl).Caption = strArwDn
6780            .Controls(strSortLbl).ForeColor = CLR_DKRED
6790          Else
6800            strSortNow = strSortAsc
6810          End If
6820        End If
6830      End If
6840      .Controls(strSortLbl).Visible = True
6850      .Controls(strSortLine).Visible = True
6860      .OrderBy = strSortNow
6870      .OrderByOn = True
6880    End With

        ' ** Potential Sorts:
        ' **   [alphasort]
        ' **   [shortname]
        ' **   [total_shareface], [alphasort]
        ' **   [total_icash], [alphasort]
        ' **   [total_icash_usd], [alphasort]

EXITP:
6890    Exit Sub

ERRH:
6900    Select Case ERR.Number
        Case Else
6910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6920    End Select
6930    Resume EXITP

End Sub

Public Function SortNow_Get() As String

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

7010    strRetVal = strSortNow

EXITP:
7020    SortNow_Get = strRetVal
7030    Exit Function

ERRH:
7040    strRetVal = vbNullString
7050    Select Case ERR.Number
        Case Else
7060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7070    End Select
7080    Resume EXITP

End Function
