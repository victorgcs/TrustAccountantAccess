VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Columns_TaxLot_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Columns_TaxLot_Sub"

'VGC 03/11/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Average:          A {opgChoice_optAverage on frmTaxLot}
' **   Cancel:           C {cmdCancel on frmTaxLot}
' **   FIFO:             F {opgChoice_optFIFO on frmTaxLot}
' **   High:             H {opgChoice_optHigh on frmTaxLot}
' **   Low:              L {opgChoice_optLow on frmTaxLot}
' **   OK:               O {cmdOK on frmTaxLot}
' **   Print:            P {cmdPrintReport on frmTaxLot}
' **   Shareface:        R {shareface}
' **   Specific:         S {opgChoice_optSpecific on frmTaxLot}
' **   Location:         T {Location_ID}
' **   Include Currency: U {chkIncludeCurrency on frmTaxLot}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** Location_ID and Location_Namex are the only real Location fields.

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Array: arr_varRec().
'Private lngRecs As Long, arr_varRec() As Variant
'Private Const R_ELEMS As Integer = 7  ' ** Array's first-element UBound().
'Private Const R_TID   As Integer = 0
'Private Const R_ASTNO As Integer = 1
'Private Const R_SHRS  As Integer = 2
'Private Const R_ADAT  As Integer = 3
'Private Const R_ZERO  As Integer = 4
'Private Const R_COST  As Integer = 5
'Private Const R_USD   As Integer = 6
'Private Const R_PPU   As Integer = 7
'Private Const R_CURID As Integer = 8

Private Const strSortOrig As String = "[assetdate]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngForm_Width As Long, lngCurrID_Left As Long, lngCurrIDLbl_Left As Long, lngCostUSD_Left As Long, lngCost_Left As Long
Private lngLocID_Left As Long, lngHoldingPeriod_Left As Long, lngPricePerUnit_Left As Long, lngTotLbl_Offset As Long

Private lngRecsCur As Long, lngTpp As Long
Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmJournal_Columns_TaxLot" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190         .form_width_line.Left = 0&  ' ** Just get it out of the way.

200         lngSortLbl_Width = .Sort_lbl.Width

210         lngForm_Width = .Width
220         lngLocID_Left = .Location_ID.Left
230         lngHoldingPeriod_Left = .Holding_Period.Left
240         lngCurrID_Left = .curr_id.Left
250         lngCurrIDLbl_Left = .curr_id_lbl.Left
260         lngPricePerUnit_Left = .priceperunit.Left
270         lngCostUSD_Left = .cost_usd.Left
280         lngCost_Left = .Cost.Left
290         lngTotLbl_Offset = ((.cost_sum.Left + .cost_sum.Width) - .cost_sum_lbl.Left)

            ' ** Holding_Period:
            ' ** =IIf(IsNull([assetdate]),"ST",IIf(DateDiff("d",[assetdate],IIf([AssetDateNull]=True,Now(),[AssetDateSale]))<=365,"ST","LT"))

300       Else
310         Cancel = -1
320       End If
330     End With

340     If Cancel = -1 Then
350       Beep
360       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
370     End If

EXITP:
380     Exit Sub

ERRH:
390     Select Case ERR.Number
        Case Else
400       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
410     End Select
420     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me
520       .assetdate.ForeColor = CLR_DISABLED_FG
530       .assetdate.BackColor = CLR_DISABLED_BG
540       .shareface.ForeColor = CLR_DISABLED_FG
550       .shareface.BackColor = CLR_DISABLED_BG
560       .Cost.ForeColor = CLR_DISABLED_FG
570       .Cost.BackColor = CLR_DISABLED_BG
580       .cost_usd.ForeColor = CLR_DISABLED_FG
590       .cost_usd.BackColor = CLR_DISABLED_BG
600       .averagepriceperunit.ForeColor = CLR_DISABLED_FG
610       .averagepriceperunit.BackColor = CLR_DISABLED_BG
620       .priceperunit.ForeColor = CLR_DISABLED_FG
630       .priceperunit.BackColor = CLR_DISABLED_BG
640       .curr_id.ForeColor = CLR_DISABLED_FG
650       .curr_id.BackColor = CLR_DISABLED_BG
660       .Holding_Period.ForeColor = CLR_DISABLED_FG
670       .Holding_Period.BackColor = CLR_DISABLED_BG
680     End With

690     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
700     Exit Sub

ERRH:
710     Select Case ERR.Number
        Case Else
720       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
730     End Select
740     Resume EXITP

End Sub

Private Sub Form_Current()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim strTmp01 As String

810     With Me

820       If .currsym_symbol = "CHF" Then
830         strTmp01 = "\C\H\F #,##0.00;(\C\H\F #,##0.00);\C\H\F 0.00;\C\H\F 0.00"
840       Else
850         strTmp01 = .currsym_symbol & "#,##0.00;(" & .currsym_symbol & "#,##0.00);" & .currsym_symbol & "0.00;" & .currsym_symbol & "0.00"
860       End If
870       .Cost.Format = strTmp01
880       .cost_sum.Format = strTmp01

890       Select Case .IsZero
          Case True
900         .FocusHolder.SetFocus
910         .Loc_Name_display.ForeColor = CLR_DISABLED_FG
920         .Loc_Name_display.BackColor = CLR_DISABLED_BG
930         .Location_ID.Locked = True
940         .Location_ID.ForeColor = CLR_DISABLED_FG
950         .Location_ID.BackColor = CLR_DISABLED_BG
960       Case False
970         If .Loc_Name_display.ForeColor <> CLR_VDKGRY Then
980           .Loc_Name_display.ForeColor = CLR_VDKGRY
990         End If
1000        If .Loc_Name_display.BackColor <> CLR_WHT Then
1010          .Loc_Name_display.BackColor = CLR_WHT
1020        End If
1030        If .Location_ID.Locked = True Then
1040          .Location_ID.Locked = False
1050        End If
1060        If .Location_ID.ForeColor <> CLR_VDKGRY Then
1070          .Location_ID.ForeColor = CLR_VDKGRY
1080        End If
1090        If .Location_ID.BackColor <> CLR_WHT Then
1100          .Location_ID.BackColor = CLR_WHT
1110        End If
1120      End Select

1130      Select Case .IsLess
          Case True

1140      Case False

1150      End Select

1160    End With

EXITP:
1170    Exit Sub

ERRH:
1180    Select Case ERR.Number
        Case Else
1190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1200    End Select
1210    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1320    intShiftDown = (Shift And acShiftMask) > 0
1330    intAltDown = (Shift And acAltMask) > 0
1340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Shareface:        R {shareface}
        ' **   Location:         T {Location_ID}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyUp
1370        intRetVal = 0
1380        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1390      Case vbKeyDown
1400        intRetVal = 0
1410        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1420      Case vbKeyEscape
1430        With Me
1440          intRetVal = 0
1450          .Parent.cmdCancel_Click  ' ** Form Procedure: frmTaxLot.
1460        End With
1470      End Select
1480    End If

        ' ** Alt keys.
1490    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1500      Select Case intRetVal
          Case vbKeyR
1510        With Me
1520          intRetVal = 0
1530          If .shareface.Enabled = True Then
1540            .shareface.SetFocus
1550          Else
1560            Beep
1570          End If
1580        End With
1590      Case vbKeyT
1600        With Me
1610          intRetVal = 0
1620          .Location_ID.SetFocus
1630        End With
1640      End Select
1650    End If

        ' ** Ctrl keys.
1660    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1670      Select Case intRetVal
          Case vbKeyS
1680        intRetVal = 0
1690        cmdSave_Click  ' ** Procedure: Below.
1700      Case vbKeyTab, vbKeyReturn
1710        With Me
1720          intRetVal = 0
1730          DoCmd.SelectObject acForm, .Parent.Name, False
1740          If .Parent.cmdPrintReport.Enabled = True Then
1750            .Parent.cmdPrintReport.SetFocus
1760          ElseIf .Parent.cmdOK.Enabled = True Then
1770            .Parent.cmdOK.SetFocus
1780          Else
1790            .Parent.cmdCancel.SetFocus
1800          End If
1810        End With
1820      Case vbKeyUp, vbKeyPageUp
1830        intRetVal = 0
1840        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1850      Case vbKeyDown, vbKeyPageDown
1860        intRetVal = 0
1870        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1880      End Select
1890    End If

        ' ** Ctrl-Shift keys.
1900    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1910      Select Case intRetVal
          Case vbKeyF
1920        With Me
1930          intRetVal = 0
1940          DoCmd.SelectObject acForm, .Parent.Name, False
1950          .Parent.FocusHolder.SetFocus
1960        End With
1970      Case vbKeyTab, vbKeyReturn
1980        With Me
1990          intRetVal = 0
2000          DoCmd.SelectObject acForm, .Parent.Name, False
2010          .Parent.opgChoice.SetFocus
2020        End With
2030      End Select
2040    End If

EXITP:
2050    KeyCode = intRetVal
2060    Exit Sub

ERRH:
2070    intRetVal = 0
2080    THAT_PROC = THIS_PROC
2090    That_Erl = Erl
2100    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2110    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControlName As String

2210    If DataErr > 0 Then
2220      Select Case DataErr
          Case 13, 2108, 2110, 2116, 2169, 2501, 3020, 3314, 7753, 7787, 8519
            ' ** 13  : Type mismatch.
            ' ** 2108: You must save the field before you execute the GoToControl action, the GoToControl method, or the SetFocus method.
            ' ** 2110: Microsoft Access can't move the focus to the control |.
            ' ** 2116: The value in the field or record violates the validation rule for the record or field.
            ' ** 2169: You can't save this record at this time.
            ' ** 2501: The '|' action was Canceled.
            ' ** 3020: Update or CancelUpdate without AddNew or Edit.
            ' ** 3314: The field '|' cannot contain a Null value because the Required property for this field is set to True.
            ' ** 7753: The value you entered doesn't meet the validation rule defined for the field or control.
            ' ** 7787: The record your are modifying has been changed by another user.
            ' ** 8519: You are about to delete | record(s).
            ' ** Do nothing.
2230      Case 2046  ' ** The command or action isn't available now (first or last record).
2240        DoCmd.CancelEvent
2250      Case 2105  ' ** You can't go to the specified record.
2260        Response = acDataErrContinue
2270      Case 2237  ' ** The text you entered isn't an item in the list.
2280        DoCmd.CancelEvent
2290      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
2300  On Error Resume Next
2310        strControlName = Me.ActiveControl.Name
2320  On Error GoTo ERRH
2330        If strControlName = "assetdate" Then
2340          MsgBox "The Date must be in the format MM/DD/YYYY to continue.", vbInformation + vbOKOnly, "Invalid Date"
2350          Me.assetdate.text = vbNullString
2360        End If
2370      Case 3101  ' ** The Microsoft Jet database engine cannot find a record in the table '|' with key matching field(s) '|'.
2380        DoCmd.CancelEvent
2390      Case 3162 ' ** You tried to assign the Null value to a variable that is not a Variant data type.
2400  On Error Resume Next
2410        strControlName = Me.ActiveControl.Name
2420  On Error GoTo ERRH
2430        MsgBox "The " & strControlName & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
2440      Case Else
2450        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2460        zErrorHandler "frmTaxLot_Sub", "Undefined Form Error", DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2470      End Select
2480    End If

EXITP:
2490    Exit Sub

ERRH:
2500    Select Case ERR.Number
        Case Else
2510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2520    End Select
2530    Resume EXITP

End Sub

Public Sub cmdSave_Click()
' ** This is a fake click event!

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2610    DoCmd.RunCommand acCmdSelectRecord
2620    DoCmd.RunCommand acCmdSaveRecord

EXITP:
2630    Exit Sub

ERRH:
2640    Select Case ERR.Number
        Case Else
2650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2660    End Select
2670    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

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
2790          .Location_ID.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2850        With Me
2860          intRetVal = 0
2870          If .CurrentRecord > 1 Then
2880            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
2890            .FocusHolder.SetFocus
2900          Else
2910            DoCmd.SelectObject acForm, .Parent.Name, False
2920            .Parent.opgChoice.setfocuso
2930          End If
2940        End With
2950      End Select
2960    End If

EXITP:
2970    KeyCode = intRetVal
2980    Exit Sub

ERRH:
2990    intRetVal = 0
3000    THAT_PROC = THIS_PROC
3010    That_Erl = Erl
3020    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3030    Resume EXITP

End Sub

Private Sub assetdate_display_lbl_DblClick(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "assetdate_display_lbl_DblClick"

3110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub shareface_lbl_DblClick(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "shareface_lbl_DblClick"

3210    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub cost_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_lbl_DblClick"

3310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub cost_usd_lbl_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cost_usd_lbl_DblClick"

3410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub priceperunit_lbl_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "priceperunit_lbl_DblClick"

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

Private Sub averagepriceperunit_lbl_DblClick(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "averagepriceperunit_lbl_DblClick"

3610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

3710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub Holding_Period_lbl_DblClick(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Holding_Period_lbl_DblClick"

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

Private Sub Loc_Name_display_GotFocus()
' ** This control sits on top of the Location_ID combo box.
' ** It's used to display the Loc_Name.
' ** CONTROL IS ENABLED, BUT LOCKED.

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_display_GotFocus"

3910    With Me
3920      If IsNull(.Location_ID) = True Then
3930        .Location_ID = 1&  ' ** {Unassigned}, {no entry}.
3940        .Location_Namex = "{Unassigned}"
3950        .Loc_Name_display.Requery
3960      End If
3970      .Location_ID.SetFocus
3980    End With

EXITP:
3990    Exit Sub

ERRH:
4000    Select Case ERR.Number
        Case Else
4010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4020    End Select
4030    Resume EXITP

End Sub

Private Sub Loc_Name_display_KeyDown(KeyCode As Integer, Shift As Integer)
' ** The focus should have already been moved to Location_ID before this can get hit.

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_display_KeyDown"

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
4190          .Location_ID.SetFocus
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
4290            .FocusHolder.SetFocus
4300          Else
4310            DoCmd.SelectObject acForm, .Parent.Name, False
4320            .Parent.opgChoice.SetFocus
4330          End If
4340        End With
4350      End Select
4360    End If

EXITP:
4370    KeyCode = intRetVal
4380    Exit Sub

ERRH:
4390    intRetVal = 0
4400    THAT_PROC = THIS_PROC
4410    That_Erl = Erl
4420    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4430    Resume EXITP

End Sub

Private Sub Location_ID_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_lbl_DblClick"

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

Private Sub Location_ID_GotFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_GotFocus"

4610    With Me
4620      .Location_ID.SelLength = 0
4630      .Location_ID.SelStart = 99
4640    End With

EXITP:
4650    Exit Sub

ERRH:
4660    Select Case ERR.Number
        Case Else
4670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4680    End Select
4690    Resume EXITP

End Sub

Private Sub Location_ID_AfterUpdate()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_AfterUpdate"

4710    With Me
4720      If IsNull(.Location_ID) = False Then
4730        .Location_Namex = .Location_ID.Column(1)
4740        .Loc_Name_display.Requery
4750      Else
4760        .Location_ID = 1&  ' ** {Unassigned}, {no entry}.
4770        .Location_Namex = "{Unassigned}"
4780        .Loc_Name_display.Requery
4790      End If
4800      cmdSave_Click  ' ** Procedure: Above.
4810      LocUpdate_TL  ' ** Procedure: Below.
4820    End With

EXITP:
4830    Exit Sub

ERRH:
4840    Select Case ERR.Number
        Case Else
4850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860    End Select
4870    Resume EXITP

End Sub

Private Sub Location_ID_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Location_ID_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4920    intShiftDown = (Shift And acShiftMask) > 0
4930    intAltDown = (Shift And acAltMask) > 0
4940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4970        With Me
4980          intRetVal = 0
4990          lngRecsCur = RecCnt  ' ** Function: Below.
5000          If .CurrentRecord < lngRecsCur Then
5010            DoCmd.RunCommand acCmdRecordsGoToNext
5020            .FocusHolder.SetFocus
5030          Else
5040            If .Parent.cmdPrintReport.Enabled = True Then
5050              .Parent.cmdPrintReport.SetFocus
5060            ElseIf .Parent.cmdOK.Enabled = True Then
5070              .Parent.cmdOK.SetFocus
5080            Else
5090              .Parent.cmdCancel.SetFocus
5100            End If
5110          End If
5120        End With
5130      End Select
5140    End If

        ' ** Shift keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5170        With Me
5180          intRetVal = 0
5190          If .CurrentRecord > 1 Then
5200            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5210            .FocusHolder.SetFocus
5220          Else
5230            DoCmd.SelectObject acForm, .Parent.Name, False
5240            .Parent.opgChoice.setfocuso
5250          End If
5260        End With
5270      End Select
5280    End If

EXITP:
5290    KeyCode = intRetVal
5300    Exit Sub

ERRH:
5310    intRetVal = 0
5320    THAT_PROC = THIS_PROC
5330    That_Erl = Erl
5340    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5350    Resume EXITP

End Sub

Private Sub LocUpdate_TL()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "LocUpdate_TL"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

5410    DoCmd.Hourglass True
5420    DoEvents

5430    Set dbs = CurrentDb
5440    With dbs
          ' ** Update ActiveAssets, by specified [actno], [astno], [astdat], [locid].
5450      Set qdf = .QueryDefs("qryLotInformation_28")
5460      With qdf.Parameters
5470        ![actno] = Me.accountno
5480        ![astno] = Me.assetno
5490        ![astdat] = Me.assetdate
5500        ![locid] = Me.Location_ID
5510      End With
5520      qdf.Execute dbFailOnError
5530      .Close
5540    End With

5550    DoCmd.Hourglass False

EXITP:
5560    Set qdf = Nothing
5570    Set dbs = Nothing
5580    Exit Sub

ERRH:
5590    DoCmd.Hourglass False
5600    Select Case ERR.Number
        Case Else
5610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5620    End Select
5630    Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnShow As Boolean)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngTmp01 As Long

5710    With Me
5720      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
5730        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
5740      End If
5750      Select Case blnShow
          Case True
5760        .Width = lngForm_Width
5770        .Location_ID.Left = lngLocID_Left
5780        .Loc_Name_display.Left = .Location_ID.Left
5790        .Location_ID_lbl.Left = .Location_ID.Left
5800        .Location_ID_lbl_line.Left = .Location_ID_lbl.Left
5810        .Holding_Period.Left = lngHoldingPeriod_Left
5820        .Holding_Period_box.Left = .Holding_Period.Left
5830        .Holding_Period_IsZero_box.Left = (.Holding_Period_box.Left - lngTpp)
5840        .Holding_Period_lbl.Left = .Holding_Period.Left
5850        .Holding_Period_lbl2.Left = .Holding_Period_lbl.Left
5860        .Holding_Period_lbl_line.Left = .Holding_Period_lbl.Left
5870        .curr_id.Left = lngCurrID_Left
5880        .curr_id_box.Left = .curr_id.Left
5890        .curr_id_IsZero_box.Left = (.curr_id_box.Left - lngTpp)
5900        .curr_id_lbl.Left = lngCurrIDLbl_Left
5910        .curr_id_lbl_line.Left = .curr_id.Left
5920        .curr_id.Visible = True
5930        .curr_id_box.Visible = True
5940        .curr_id_IsZero_box.Visible = True
5950        .curr_id_lbl.Visible = True
5960        .curr_id_lbl_line.Visible = True
5970        .priceperunit.Left = lngPricePerUnit_Left
5980        .priceperunit_box.Left = .priceperunit.Left
5990        .priceperunit_IsZero_box.Left = (.priceperunit_box.Left - lngTpp)
6000        .priceperunit_lbl.Left = .priceperunit.Left
6010        .priceperunit_lbl_line.Left = .priceperunit_lbl.Left
6020        .priceperunit_avg.Left = lngPricePerUnit_Left
6030        .priceperunit_avg_lbl.Left = (.priceperunit_avg.Left + .priceperunit_avg.Width)
6040        .priceperunit_avg_line.Left = .priceperunit_avg.Left
6050        .averagepriceperunit.Left = .priceperunit.Left
6060        .averagepriceperunit_avg.Left = .priceperunit_avg.Left
6070        .averagepriceperunit_lbl.Left = .priceperunit_lbl.Left
6080        .averagepriceperunit_lbl_line.Left = .averagepriceperunit_lbl.Left
6090        .cost_usd.Left = lngCostUSD_Left
6100        .cost_usd_box.Left = .cost_usd.Left
6110        .cost_usd_IsZero_box.Left = (.cost_usd_box.Left - lngTpp)
6120        .cost_usd_lbl.Left = .cost_usd.Left
6130        .cost_usd_lbl_line.Left = .cost_usd_lbl.Left
6140        .cost_usd_sum.Left = lngCostUSD_Left
6150        .cost_usd_sum_lbl.Left = ((.cost_usd_sum.Left + .cost_usd_sum.Width) - lngTotLbl_Offset)
6160        .cost_usd_sum_line.Left = .cost_usd_sum.Left
6170        .cost_usd.Visible = True
6180        .cost_usd_box.Visible = True
6190        .cost_usd_IsZero_box.Visible = True
6200        .cost_usd_lbl.Visible = True
6210        .cost_usd_lbl_line.Visible = True
6220        .cost_usd_sum.Visible = True
6230        .cost_usd_sum_lbl.Visible = True
6240        .cost_usd_sum_line.Visible = True
6250      Case False
6260        .curr_id.Visible = False
6270        .curr_id_box.Visible = False
6280        .curr_id_IsZero_box.Visible = False
6290        .curr_id_lbl.Visible = False
6300        .curr_id_lbl_line.Visible = False
6310        .cost_usd.Visible = False
6320        .cost_usd_box.Visible = False
6330        .cost_usd_IsZero_box.Visible = False
6340        .cost_usd_lbl.Visible = False
6350        .cost_usd_lbl_line.Visible = False
6360        .cost_usd_sum.Visible = False
6370        .cost_usd_sum_lbl.Visible = False
6380        .cost_usd_sum_line.Visible = False
6390        lngTmp01 = (lngPricePerUnit_Left - lngCostUSD_Left)
6400        .priceperunit.Left = (lngPricePerUnit_Left - lngTmp01)
6410        .priceperunit_box.Left = .priceperunit.Left
6420        .priceperunit_IsZero_box.Left = (.priceperunit_box.Left - lngTpp)
6430        .priceperunit_lbl.Left = .priceperunit.Left
6440        .priceperunit_lbl_line.Left = .priceperunit_lbl.Left
6450        .priceperunit_avg.Left = (lngPricePerUnit_Left - lngTmp01)
6460        .priceperunit_avg_lbl.Left = (.priceperunit_avg.Left + .priceperunit_avg.Width)
6470        .priceperunit_avg_line.Left = .priceperunit_avg.Left
6480        .averagepriceperunit.Left = .priceperunit.Left
6490        .averagepriceperunit_avg.Left = .priceperunit_avg.Left
6500        .averagepriceperunit_lbl.Left = .priceperunit_lbl.Left
6510        .averagepriceperunit_lbl_line.Left = .averagepriceperunit_lbl.Left
6520        lngTmp01 = (lngCurrID_Left - lngPricePerUnit_Left)
6530        .Holding_Period.Left = (.priceperunit.Left + lngTmp01)
6540        .Holding_Period_box.Left = .Holding_Period.Left
6550        .Holding_Period_IsZero_box.Left = (.Holding_Period_box.Left - lngTpp)
6560        .Holding_Period_lbl.Left = .Holding_Period.Left
6570        .Holding_Period_lbl2.Left = .Holding_Period_lbl.Left
6580        .Holding_Period_lbl_line.Left = .Holding_Period_lbl.Left
6590        lngTmp01 = (lngLocID_Left - lngHoldingPeriod_Left)
6600        .Location_ID.Left = (.Holding_Period.Left + lngTmp01)
6610        .Loc_Name_display.Left = .Location_ID.Left
6620        .Location_ID_lbl.Left = .Location_ID.Left
6630        .Location_ID_lbl_line.Left = .Location_ID_lbl.Left
6640        lngTmp01 = (lngCostUSD_Left - lngCost_Left)
6650        lngTmp01 = (lngTmp01 + (lngHoldingPeriod_Left - lngCurrID_Left))
6660        .IsZeroBGa.Width = (lngForm_Width - lngTmp01)
6670        .IsZeroBGb.Width = .IsZeroBGa.Width
6680        .Width = (lngForm_Width - lngTmp01)

6690      End Select
6700    End With

EXITP:
6710    Exit Sub

ERRH:
6720    Select Case ERR.Number
        Case Else
6730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6740    End Select
6750    Resume EXITP

End Sub

Public Function RecCnt() As Long

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

6810    lngRetVal = 0&

6820    With Me
6830      Set rst = .RecordsetClone
6840      With rst
6850        If .BOF = True And .EOF = True Then
              ' ** Let it return Zero.
6860        Else
6870          .MoveLast
6880          lngRetVal = .RecordCount
6890        End If
6900        .Close
6910      End With
6920    End With

EXITP:
6930    Set rst = Nothing
6940    RecCnt = lngRetVal
6950    Exit Function

ERRH:
6960    lngRetVal = 0&
6970    Select Case ERR.Number
        Case Else
6980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6990    End Select
7000    Resume EXITP

End Function

Public Sub MoveRec(lngRecsGoTo As AcCommand, Optional varRecID As Variant)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

7110    If IsMissing(varRecID) = True Then
7120      Select Case lngRecsGoTo
          Case acCmdRecordsGoToNew, acCmdRecordsGoToNext, acCmdRecordsGoToPrevious, acCmdRecordsGoToFirst, acCmdRecordsGoToLast
7130        DoCmd.RunCommand lngRecsGoTo
7140      Case acCmdSelectRecord
7150        DoCmd.RunCommand lngRecsGoTo
7160      End Select
7170    Else
7180      If IsNull(varRecID) = False Then
7190        If varRecID > 0& Then
7200  On Error Resume Next
7210          Set rst = Me.RecordsetClone
7220          If ERR.Number = 0 Then
7230  On Error GoTo ERRH
7240            With rst
7250  On Error Resume Next
7260              .FindFirst "[uniqueid] = '" & varRecID & "'"   ' ** Some combination of the events produced the error:
7270              If ERR.Number = 0 Then                         ' **   3420 - Object invalid or no longer set.
7280  On Error GoTo ERRH                                         ' ** on this line. The added On Error's should cover it.
7290                If .NoMatch = False Then
7300                  Me.Bookmark = .Bookmark
7310                End If
7320                .Close
7330              Else
7340  On Error GoTo ERRH
7350              End If
7360            End With
7370          Else
7380  On Error GoTo ERRH
7390          End If
7400        Else
7410          Select Case lngRecsGoTo
              Case acCmdRecordsGoToNew, acCmdRecordsGoToNext, acCmdRecordsGoToPrevious, acCmdRecordsGoToFirst, acCmdRecordsGoToLast
7420            DoCmd.RunCommand lngRecsGoTo
7430          End Select
7440        End If
7450      Else
7460        Select Case lngRecsGoTo
            Case acCmdRecordsGoToNew, acCmdRecordsGoToNext, acCmdRecordsGoToPrevious, acCmdRecordsGoToFirst, acCmdRecordsGoToLast
7470          DoCmd.RunCommand lngRecsGoTo
7480        End Select
7490      End If
7500    End If

EXITP:
7510    Set rst = Nothing
7520    Exit Sub

ERRH:
7530    Select Case ERR.Number
        Case Else
7540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7550    End Select
7560    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [assetdate]"

7610    With Me
7620      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7630        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7640      End If
7650      .Controls(strSortLbl).Visible = False
7660      .Controls(strSortLine).Visible = False
7670      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
7680      If strProc = "Form_Load" Then
7690        strCalled = "assetdate_display"
7700        strSortNow = strSortOrig
7710        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7720        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7730        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7740        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7750        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7760        .Controls(strSortLbl).Top = lngSortLbl_Top
7770        .Controls(strSortLbl).Left = lngSortLbl_Left
7780        .Controls(strSortLbl).Caption = strArwUp
7790        .Controls(strSortLbl).ForeColor = CLR_DKBLU
7800        .Controls(strSortLine).Top = lngSortLine_Top
7810        .Controls(strSortLine).Left = lngSortLine_Left
7820        .Controls(strSortLine).Width = lngSortLine_Width
7830      Else
7840        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
7850        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7860        lngSortLbl_Left = (.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width
7870        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7880        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7890        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7900        .Controls(strSortLbl).Top = lngSortLbl_Top
7910        .Controls(strSortLbl).Left = lngSortLbl_Left
7920        .Controls(strSortLine).Top = lngSortLine_Top
7930        .Controls(strSortLine).Left = lngSortLine_Left
7940        .Controls(strSortLine).Width = lngSortLine_Width
7950        If strCalled = "assetdate_display" Then
7960          strSortAsc = strSortOrig
7970          If strSortNow = strSortAsc Then
7980            strSortNow = "[assetdate] DESC"
7990            .Controls(strSortLbl).Caption = strArwDn
8000            .Controls(strSortLbl).ForeColor = CLR_DKRED
8010          Else
8020            strSortNow = strSortAsc
8030            .Controls(strSortLbl).Caption = strArwUp
8040            .Controls(strSortLbl).ForeColor = CLR_DKBLU
8050          End If
8060        Else
8070          .Controls(strSortLbl).Caption = strArwUp
8080          .Controls(strSortLbl).ForeColor = CLR_DKBLU
8090          Select Case strCalled
              Case "shareface"
8100            strSortAsc = "[shareface]" & strStdAsc
8110          Case "cost"
8120            strSortAsc = "[cost]" & strStdAsc
8130          Case "cost_usd"
8140            strSortAsc = "[cost_usd]" & strStdAsc
8150          Case "priceperunit"
8160            strSortAsc = "[priceperunit]" & strStdAsc
8170          Case "averagepriceperunit"
8180            strSortAsc = "[averagepriceperunit]" & strStdAsc
8190          Case "curr_id"
8200            strSortAsc = "[curr_code]" & strStdAsc
8210            lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
8220            lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
8230            lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
8240            .Controls(strSortLbl).Left = lngSortLbl_Left
8250            .Controls(strSortLine).Left = lngSortLine_Left
8260            .Controls(strSortLine).Width = lngSortLine_Width
8270          Case "Holding_Period"
8280            strSortAsc = "[Holding_Period]" & strStdAsc
8290          Case "Location_ID", "Loc_Name_display"
8300            strSortAsc = "[Loc_Name]" & strStdAsc
8310          End Select
8320          If strSortNow = strSortAsc Then
8330            intCnt = CharCnt(strSortAsc, ",") + 1
8340            Select Case intCnt
                Case 1
8350              strTmp01 = strSortAsc & " DESC"
8360            Case 2
8370              intPos01 = InStr(strSortAsc, ",")
8380              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8390              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
8400            Case 3
8410              intPos01 = InStr(strSortAsc, ",")
8420              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
8430              strTmp02 = Mid(strSortAsc, intPos01)
8440              intPos01 = InStr(2, strTmp02, ",")
8450              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
8460              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
8470            End Select
8480            strSortNow = strTmp01
8490            .Controls(strSortLbl).Caption = strArwDn
8500            .Controls(strSortLbl).ForeColor = CLR_DKRED
8510          Else
8520            strSortNow = strSortAsc
8530          End If
8540        End If
8550      End If
8560      .Controls(strSortLbl).Visible = True
8570      .Controls(strSortLine).Visible = True
8580      .OrderBy = strSortNow
8590      .OrderByOn = True
8600    End With

        ' ** Potential Sorts:
        ' **   [assetdate]
        ' **   [shareface], [assetdate]
        ' **   [cost], [assetdate]
        ' **   [cost_usd], [assetdate]
        ' **   [priceperunit], [assetdate]
        ' **   [averagepriceperunit], [assetdate]
        ' **   [curr_code], [assetdate]
        ' **   [Holding_Period], [assetdate]
        ' **   [Loc_Name], [assetdate]

EXITP:
8610    Exit Sub

ERRH:
8620    Select Case ERR.Number
        Case Else
8630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8640    End Select
8650    Resume EXITP

End Sub

Public Sub SortNow_Set(strNewSort As String)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Set"

8710    strSortNow = strNewSort

EXITP:
8720    Exit Sub

ERRH:
8730    Select Case ERR.Number
        Case Else
8740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8750    End Select
8760    Resume EXITP

End Sub

Public Function SortNow_Get() As String

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

8810    strRetVal = strSortNow

EXITP:
8820    SortNow_Get = strRetVal
8830    Exit Function

ERRH:
8840    strRetVal = vbNullString
8850    Select Case ERR.Number
        Case Else
8860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8870    End Select
8880    Resume EXITP

End Function
