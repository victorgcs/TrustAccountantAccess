VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssetPricing_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssetPricing_Sub"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose on frmAssetPricing}
' **   New Date:         D {Date_New on frmAssetPricing}
' **   Reset:            E {cmdReset on frmAssetPricing}
' **   Show Hidden:      H {chkShowHidden on frmAssetPricing}
' **   History:          I {cmdHistory on frmAssetPricing}
' **   Incl Currency:    L {chkIncludeCurrency on frmAssetPricing}
' **   Market Value:     M {marketvalue}
' **   Show AssetNo:     N {chkShowAssetNo on frmAssetPricing}
' **   Only Assets Held: O {opgShow_optHeldOnly on frmAssetPricing}
' **   Print:            P {cmdPrintReport on frmAssetPricing}
' **   Price:            R {cmdPrice on frmAssetPricing}
' **   Shorten:          S {cmdShorten}
' **   All Assets:       T {opgShow_optAll on frmAssetPricing}
' **   Update:           U {cmdUpdate on frmAssetPricing}
' **   Widen:            W {cmdWiden}
' **   Exit:             X {cmdClose on frmAssetPricing}
' **   Yield:            Y {yield_entry}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:          F5 {ForExTots on frmAssetPricing}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

'.._upd.ForeColor: 16774128  CLR_LTBLU

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[assettype], [description_masterasset_sort], [cusip]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private blnShorten_Focus As Boolean, blnShorten_MouseDown As Boolean
Private blnWiden_Focus As Boolean, blnWiden_MouseDown As Boolean

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private lngForm_Width As Long, lngTotDesc_Left As Long, lngTotDesc_Width As Long, lngTotDescAstno_Width As Long, lngAssetNext_Left As Long
Private lngYield_Left As Long, lngYieldPct_Offset As Long, lngMktValSym_Offset As Long
Private lngCurrID_Left As Long, lngCurrID_Width As Long, lngMktValCurUsdStr_Left As Long, lngMktValSymUsd_Left As Long
Private lngMktVal_Left As Long, lngMktValUsd_Left As Long, lngMktValCurStr_Left As Long
Private lngMktValUsd_Width As Long, lngMktValCurUsdStr_Width As Long
Private lngMktValSym_Left As Long

Private lngCurrID_PrevVal As Long, blnIsOpen As Boolean, blnFromTimer As Boolean
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String, lngRecsCur As Long, lngTpp As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAssetPricing" Then

160         blnIsOpen = True

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

190         blnShorten_Focus = False: blnShorten_MouseDown = False: blnFromTimer = False
200         blnWiden_Focus = False: blnWiden_MouseDown = False

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
210         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

            ' ** Move these out of the way.
220         .form_width_line.Left = 0&
230         .marketvaluecurrent.Left = 0&
240         .MarketValueCurrent_usd.Left = 0&
250         .curr_code.Left = 0&
260         .curr_rate1.Left = 0&
270         .curr_rate2.Left = 0&

280         lngSortLbl_Width = .Sort_lbl.Width

290         lngForm_Width = .Width
300         lngTotDesc_Left = .totdesc.Left
310         lngTotDesc_Width = .totdesc.Width
320         lngTotDescAstno_Width = (lngTotDesc_Width - (.assetno.Width + (4& * lngTpp)))
330         lngAssetNext_Left = .assettype.Left
340         lngYield_Left = .yield_entry.Left
350         lngYieldPct_Offset = (.yield_pct_lbl.Left - .yield_entry.Left)
360         lngMktVal_Left = .marketvalue.Left
370         lngMktValUsd_Left = .marketvalue_usd.Left
380         lngMktValCurStr_Left = .marketvaluecurrent_str.Left
390         lngMktValCurUsdStr_Left = .marketvaluecurrent_usd_str.Left
400         lngMktValSym_Left = .currsym_symbol.Left
410         lngMktValSymUsd_Left = .currsym_symbol_usd.Left
420         lngMktValSym_Offset = (.marketvalue.Left - .currsym_symbol.Left)
430         lngCurrID_Left = .curr_id.Left
440         lngCurrID_Width = ((.curr_id.Left + .curr_id.Width) - (.assettype.Left + .assettype.Width))
450         lngMktValUsd_Width = (.marketvalue_usd.Left + .marketvalue_usd.Width) - (.marketvalue.Left + .marketvalue.Width)
460         lngMktValCurUsdStr_Width = ((.marketvaluecurrent_usd_str.Left + .marketvaluecurrent_usd_str.Width) - _
              (.marketvaluecurrent_str.Left + .marketvaluecurrent_str.Width))

            ' ** Since there's nowhere to go, disable cmdShorten.
470         .cmdShorten.Enabled = False
480         .cmdShorten_raised_img_dis.Visible = True
490         .cmdShorten_raised_img.Visible = False
500         .cmdShorten_raised_semifocus_dots_img.Visible = False
510         .cmdShorten_raised_focus_img.Visible = False
520         .cmdShorten_raised_focus_dots_img.Visible = False
530         .cmdShorten_sunken_focus_dots_img.Visible = False

            'lngYield_Left
            'lngCurrID_Left
            'lngMktVal_Left
            'lngMktValUsd_Left
            'lngMktValCurStr_Left
            'lngMktValCurUsdStr_Left
            'lngMktValSym_Left
            'lngMktValSymUsd_Left

540       Else
550         Cancel = -1
560       End If
570     End With

580     If Cancel = -1 Then
590       DoCmd.Hourglass False
600       Beep
610       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
620     End If

EXITP:
630     Exit Sub

ERRH:
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub Form_Load()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

710     With Me

720       .assetno.ForeColor = CLR_DISABLED_FG
730       .assetno.BackColor = CLR_DISABLED_BG
740       .cusip.ForeColor = CLR_DISABLED_FG
750       .cusip.BackColor = CLR_DISABLED_BG
760       .totdesc.ForeColor = CLR_DISABLED_FG
770       .totdesc.BackColor = CLR_DISABLED_BG
780       .assettype.ForeColor = CLR_DISABLED_FG
790       .assettype.BackColor = CLR_DISABLED_BG
800       .curr_id.ForeColor = CLR_DISABLED_FG
810       .curr_id.BackColor = CLR_DISABLED_BG
820       .yield_entry.ForeColor = CLR_DISABLED_FG
830       .yield_box.BackColor = CLR_DISABLED_BG
840       .marketvaluecurrent_str.ForeColor = CLR_DISABLED_FG
850       .marketvaluecurrent_str.BackColor = CLR_DISABLED_BG
860       .marketvaluecurrent_usd_str.ForeColor = CLR_DISABLED_FG
870       .marketvaluecurrent_usd_str.BackColor = CLR_DISABLED_BG
880       .marketvalue.ForeColor = CLR_DISABLED_FG
890       .marketvalue_box.BackColor = CLR_DISABLED_BG
900       .marketvalue_loc_usd_box.ForeColor = CLR_DISABLED_BG
910       .currsym_symbol_loc_usd_box.ForeColor = CLR_DISABLED_BG
920       .marketvalue_usd.ForeColor = CLR_DISABLED_FG
930       .marketvalue_usd_box.BackColor = CLR_DISABLED_BG
940       lngRecsCur = .RecCnt  ' ** Function: Below.
950       If lngRecsCur > 0& Then
960         .Parent.Date_Current = .currentDate
970         .FocusHolder.SetFocus
980       Else
990         .Parent.Date_Current = Date
1000      End If
1010      DoCmd.SelectObject acForm, .Parent.Name, False
1020      .Parent.Date_New.SetFocus

1030    End With

1040    SortNow THIS_PROC

EXITP:
1050    Exit Sub

ERRH:
1060    Select Case ERR.Number
        Case Else
1070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1080    End Select
1090    Resume EXITP

End Sub

Private Sub Form_Current()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

1110    With Me
1120      If .rate > 0 And .yield <> .rate Then
1130        .yield = .rate
1140        .yield_entry = (.yield * 100#)
1150      End If
1160    End With

EXITP:
1170    Exit Sub

ERRH:
1180    Select Case ERR.Number
        Case 2424  ' ** The expression you entered has a field, control, or property name that Microsoft Access can't find.
          ' ** I have no idea what this is about!
1190    Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

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
        ' **   Market Value:     M {marketvalue}
        ' **   Shorten:          S {cmdShorten}
        ' **   Widen:            W {cmdWiden}
        ' **   Exit:             X {cmdClose on frmAssetPricing}
        ' **   Yield:            Y {yield_entry}

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:          F5 {ForExTots on frmAssetPricing}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyF5
1370        With Me
1380          intRetVal = 0
1390          .Parent.ForExTots_Set True  ' ** Form Procedure: frmAssetPricing.
1400        End With
1410      Case vbKeyUp
1420        intRetVal = 0
1430        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1440      Case vbKeyDown
1450        intRetVal = 0
1460        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1470      Case vbKeyEscape
1480        With Me
1490          intRetVal = 0
1500          .Parent.cmdClose_Click  ' ** Form Procedure: frmAssetPricing.
1510        End With
1520      End Select
1530    End If

        ' ** Alt keys.
1540    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1550      Select Case intRetVal
          Case vbKeyM
1560        With Me
1570          intRetVal = 0
1580          If .marketvalue.Enabled = True Then
1590            .marketvalue.SetFocus
1600          End If
1610        End With
1620      Case vbKeyS
            ' ** Invoke the shorten-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1630        With Me
1640          intRetVal = 0
1650          If .cmdShorten.Enabled = True Then
1660            cmdShorten_Click  ' ** Procedure: Below.
1670          Else
1680            Beep
1690          End If
1700        End With
1710      Case vbKeyW
            ' ** Invoke the widen-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1720        With Me
1730          intRetVal = 0
1740          If .cmdWiden.Enabled = True Then
1750            cmdWiden_Click  ' ** Procedure: Below.
1760          Else
1770            Beep
1780          End If
1790        End With
1800      Case vbKeyX
1810        With Me
1820          intRetVal = 0
1830          .Parent.cmdClose_Click  ' ** Form Procedure: frmAssetPricing.
1840        End With
1850      Case vbKeyY
1860        With Me
1870          intRetVal = 0
1880          If .yield_entry.Enabled = True Then
1890            .yield_entry.SetFocus
1900          End If
1910        End With
1920      End Select
1930    End If

        ' ** Ctrl keys.
1940    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1950      Select Case intRetVal
          Case vbKeyS
1960        intRetVal = 0
1970        cmdSave_Click  ' ** Procedure: Below.
1980      Case vbKeyTab, vbKeyReturn
1990        With Me
2000          intRetVal = 0
2010          DoCmd.SelectObject acForm, .Parent.Name, False
2020          .Parent.opgShow.SetFocus
2030        End With
2040      Case vbKeyPageUp
2050        intRetVal = 0
2060        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2070      Case vbKeyPageDown
2080        intRetVal = 0
2090        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2100      End Select
2110    End If

        ' ** Ctrl-Shift keys.
2120    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2130      Select Case intRetVal
          Case vbKeyF
2140        With Me
2150          intRetVal = 0
2160          DoCmd.SelectObject acForm, .Parent.Name, False
2170          .Parent.FocusHolder.SetFocus
2180        End With
2190      Case vbKeyTab, vbKeyReturn
2200        With Me
2210          intRetVal = 0
2220          DoCmd.SelectObject acForm, .Parent.Name, False
2230          .Parent.cmdClose.SetFocus
2240        End With
2250      End Select
2260    End If

EXITP:
2270    KeyCode = intRetVal
2280    Exit Sub

ERRH:
2290    intRetVal = 0
2300    THAT_PROC = THIS_PROC
2310    That_Erl = Erl
2320    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2330    Resume EXITP

End Sub

Private Sub Form_Timer()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

2410    With Me
2420      .TimerInterval = 0&
2430      If blnFromTimer = True Then
            ' ** This finally works!
2440        blnFromTimer = False
2450        .cmdShorten.Enabled = False
2460        .cmdShorten_raised_img_dis.Visible = True
2470        .cmdShorten_raised_img.Visible = False
2480        .cmdShorten_raised_semifocus_dots_img.Visible = False
2490        .cmdShorten_raised_focus_img.Visible = False
2500        .cmdShorten_raised_focus_dots_img.Visible = False
2510        .cmdShorten_sunken_focus_dots_img.Visible = False
2520      End If
2530    End With

EXITP:
2540    Exit Sub

ERRH:
2550    Select Case ERR.Number
        Case Else
2560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2570    End Select
2580    Resume EXITP

End Sub

Private Sub Form_Click()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Click"

2610    With Me
2620      .Parent.SetPricing True  ' ** Form Function: frmAssetPricing.
2630    End With

EXITP:
2640    Exit Sub

ERRH:
2650    Select Case ERR.Number
        Case Else
2660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2670    End Select
2680    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2710    If DataErr <> 0 Then
2720      DoCmd.Hourglass False
2730      Select Case DataErr
          Case 13, 2108, 2116, 2169, 2501, 3020, 3101, 3314, 7753, 8519
            ' ** 2501: The '|' action was Canceled.
2740        Response = acDataErrContinue
2750      Case 2046  ' ** The command or action '|' isn't available now.
            ' ** Ignore
2760        Response = acDataErrContinue
2770      Case 2110  ' ** Access can't move the focus to the control '|'.
2780        Response = acDataErrContinue
2790      Case 2113  ' ** 2113: The value you entered isn't valid for this field.
2800        Select Case Me.ActiveControl.Name
            Case "totdesc"
2810          MsgBox "Please enter a valid Asset.", vbInformation + vbOKOnly, "Invalid Entry"
2820        Case "cusip"
2830          MsgBox "Please enter a valid cusip.", vbInformation + vbOKOnly, "Invalid Entry"
2840        Case "yield", "yield_entry"
2850          MsgBox "Please enter a valid yield.", vbInformation + vbOKOnly, "Invalid Entry"
2860        Case "marketvalue"
2870          MsgBox "Please enter a valid market value.", vbInformation + vbOKOnly, "Invalid Entry"
2880        Case "Date_New"
2890          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
2900        End Select
2910        Response = acDataErrContinue
2920      Case 2279  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2930        MsgBox "Please enter a valid date (MM/DD/YYYY), Or" & vbCrLf & _
              "clear the contents of the date field to exit.", vbInformation + vbOKOnly, "Invalid Date"
2940        Response = acDataErrContinue
2950      Case 3162
2960        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
2970        Response = acDataErrContinue
2980      Case Else
2990        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3000        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3010      End Select
3020    End If

EXITP:
3030    Exit Sub

ERRH:
3040    DoCmd.Hourglass False
3050    Select Case ERR.Number
        Case Else
3060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3070    End Select
3080    Resume EXITP

End Sub

Private Sub Detail_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Click"

3110    Form_Click  ' ** Procedure: Above.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3210    With Me
3220      If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
3230        Select Case blnShorten_Focus
            Case True
3240          .cmdShorten_raised_semifocus_dots_img.Visible = True
3250          .cmdShorten_raised_img.Visible = False
3260        Case False
3270          .cmdShorten_raised_img.Visible = True
3280          .cmdShorten_raised_semifocus_dots_img.Visible = False
3290        End Select
3300        .cmdShorten_raised_focus_img.Visible = False
3310        .cmdShorten_raised_focus_dots_img.Visible = False
3320        .cmdShorten_sunken_focus_dots_img.Visible = False
3330        .cmdShorten_raised_img_dis.Visible = False
3340      End If
3350      If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
3360        Select Case blnWiden_Focus
            Case True
3370          .cmdWiden_raised_semifocus_dots_img.Visible = True
3380          .cmdWiden_raised_img.Visible = False
3390        Case False
3400          .cmdWiden_raised_img.Visible = True
3410          .cmdWiden_raised_semifocus_dots_img.Visible = False
3420        End Select
3430        .cmdWiden_raised_focus_img.Visible = False
3440        .cmdWiden_raised_focus_dots_img.Visible = False
3450        .cmdWiden_sunken_focus_dots_img.Visible = False
3460        .cmdWiden_raised_img_dis.Visible = False
3470      End If
3480    End With

EXITP:
3490    Exit Sub

ERRH:
3500    Select Case ERR.Number
        Case Else
3510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3520    End Select
3530    Resume EXITP

End Sub

Private Sub FormHeader_Click()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_Click"

3610    Form_Click  ' ** Procedure: Above.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

3710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub FormFooter_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_Click"

3810    Form_Click  ' ** Procedure: Above.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

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
3990          If .curr_id.Visible = True And .curr_id.Enabled = True Then
4000            .curr_id.SetFocus
4010          ElseIf .yield_entry.Enabled = True Then
4020            .yield_entry.SetFocus
4030          Else
4040            lngRecsCur = RecCnt  ' ** Function: Below.
4050            If .CurrentRecord < lngRecsCur Then
4060              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4070              .FocusHolder.SetFocus
4080            Else
4090              DoCmd.SelectObject acForm, .Parent.Name, False
4100              .Parent.opgShow.SetFocus
4110            End If
4120          End If
4130        End With
4140      End Select
4150    End If

        ' ** Shift keys.
4160    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4170      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4180        With Me
4190          intRetVal = 0
4200          If .CurrentRecord > 1 Then
4210            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4220            If .marketvalue.Enabled = True Then
4230              .marketvalue.SetFocus
4240            Else
4250              .FocusHolder.SetFocus
4260            End If
4270          Else
4280            DoCmd.SelectObject acForm, .Parent.Name, False
4290            .Parent.cmdClose.SetFocus
4300          End If
4310        End With
4320      End Select
4330    End If

EXITP:
4340    KeyCode = intRetVal
4350    Exit Sub

ERRH:
4360    intRetVal = 0
4370    THAT_PROC = THIS_PROC
4380    That_Erl = Erl
4390    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4400    Resume EXITP

End Sub

Public Sub cmdSave_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

4510    DoCmd.RunCommand acCmdSaveRecord

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub assetno_lbl_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "assetno_lbl_DblClick"

4610    SortNow THIS_PROC

EXITP:
4620    Exit Sub

ERRH:
4630    THAT_PROC = THIS_PROC
4640    That_Erl = Erl
4650    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4660    Resume EXITP

End Sub

Private Sub cusip_lbl_DblClick(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_lbl_DblClick"

4710    SortNow THIS_PROC

EXITP:
4720    Exit Sub

ERRH:
4730    THAT_PROC = THIS_PROC
4740    That_Erl = Erl
4750    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4760    Resume EXITP

End Sub

Private Sub cusip_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_KeyDown"

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
4890          If .curr_id.Visible = True And .curr_id.Enabled = True Then
4900            .curr_id.SetFocus
4910          ElseIf .yield_entry.Enabled = True Then
4920            .yield_entry.SetFocus
4930          Else
4940            lngRecsCur = RecCnt  ' ** Function: Below.
4950            If .CurrentRecord < lngRecsCur Then
4960              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4970              .FocusHolder.SetFocus
4980            Else
4990              DoCmd.SelectObject acForm, .Parent.Name, False
5000              .Parent.opgShow.SetFocus
5010            End If
5020          End If
5030        End With
5040      End Select
5050    End If

        ' ** Shift keys.
5060    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5070      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5080        With Me
5090          intRetVal = 0
5100          If .CurrentRecord > 1 Then
5110            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5120            If .marketvalue.Enabled = True Then
5130              .marketvalue.SetFocus
5140            Else
5150              .FocusHolder.SetFocus
5160            End If
5170          Else
5180            DoCmd.SelectObject acForm, .Parent.Name, False
5190            .Parent.cmdClose.SetFocus
5200          End If
5210        End With
5220      End Select
5230    End If

EXITP:
5240    KeyCode = intRetVal
5250    Exit Sub

ERRH:
5260    intRetVal = 0
5270    THAT_PROC = THIS_PROC
5280    That_Erl = Erl
5290    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5300    Resume EXITP

End Sub

Private Sub totdesc_lbl_DblClick(Cancel As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_DblClick"

5410    SortNow THIS_PROC

EXITP:
5420    Exit Sub

ERRH:
5430    THAT_PROC = THIS_PROC
5440    That_Erl = Erl
5450    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5460    Resume EXITP

End Sub

Private Sub totdesc_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_lbl_MouseMove"

5510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
5520    Exit Sub

ERRH:
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub

Private Sub totdesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "totdesc_MouseMove"

5610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: above.

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub assettype_lbl_DblClick(Cancel As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_lbl_DblClick"

5710    SortNow THIS_PROC

EXITP:
5720    Exit Sub

ERRH:
5730    THAT_PROC = THIS_PROC
5740    That_Erl = Erl
5750    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5760    Resume EXITP

End Sub

Private Sub curr_id_lbl_DblClick(Cancel As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_lbl_DblClick"

5810    SortNow THIS_PROC

EXITP:
5820    Exit Sub

ERRH:
5830    THAT_PROC = THIS_PROC
5840    That_Erl = Erl
5850    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5860    Resume EXITP

End Sub

Private Sub curr_id_DblClick(Cancel As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_DblClick"

5910    With Me
5920      Select Case .curr_date.Visible
          Case True
5930        .curr_date.Visible = False
5940      Case False
5950        .curr_date.Visible = True
5960      End Select
5970    End With

EXITP:
5980    Exit Sub

ERRH:
5990    Select Case ERR.Number
        Case Else
6000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6010    End Select
6020    Resume EXITP

End Sub

Private Sub curr_id_Enter()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_Enter"

6110    With Me
6120      lngRecsCur = RecCnt  ' ** Function: Below.
6130      If lngRecsCur > 0& Then
6140        Select Case IsNull(.curr_id)
            Case True
6150          lngCurrID_PrevVal = 150&  ' ** Default to 'USD'.
6160        Case False
6170  On Error Resume Next
6180          lngCurrID_PrevVal = .curr_id
6190          If ERR.Number <> 0 Then
6200  On Error GoTo ERRH
6210            lngCurrID_PrevVal = 0&
6220          Else
6230  On Error GoTo ERRH
6240          End If
6250        End Select
6260        If .curr_id = 150& Then
6270          .curr_id.BackColor = CLR_DISABLED_BG
6280        Else
6290          .curr_id.BackColor = CLR_VLTGRN
6300        End If
6310      End If
6320    End With

EXITP:
6330    Exit Sub

ERRH:
6340    THAT_PROC = THIS_PROC
6350    That_Erl = Erl
6360    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6370    Resume EXITP

End Sub

Private Sub curr_id_Change()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_Change"

6410    With Me
6420      If .Parent.cmdUpdate.Enabled = False Then
6430        .Parent.cmdUpdate.Enabled = True
6440      End If
6450      If .Parent.cmdReset.Enabled = False Then
6460        .Parent.cmdReset.Enabled = True
6470      End If
6480    End With

EXITP:
6490    Exit Sub

ERRH:
6500    THAT_PROC = THIS_PROC
6510    That_Erl = Erl
6520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6530    Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

6610    With Me
6620      .Parent.chkUpdated = True
6630      .AssetPricing_Changed = True
6640      If .Parent.cmdUpdate.Enabled = False Then
6650        .Parent.cmdUpdate.Enabled = True
6660      End If
6670      If .Parent.cmdReset.Enabled = False Then
6680        .Parent.cmdReset.Enabled = True
6690      End If
6700      If IsNull(.curr_id) = True Then
6710        Beep
6720        MsgBox "Currency cannot be blank." & vbCrLf & "Trust Accountant will default to 'USD'.", _
              vbInformation + vbOKOnly, "Entry Required"
6730        .curr_id = 150&  ' ** Default to 'USD'.
6740      End If
6750    End With

EXITP:
6760    Exit Sub

ERRH:
6770    THAT_PROC = THIS_PROC
6780    That_Erl = Erl
6790    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6800    Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

6910    With Me
6920      MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
6930      .curr_id.SetFocus
6940      Response = acDataErrContinue
6950    End With

EXITP:
6960    Exit Sub

ERRH:
6970    THAT_PROC = THIS_PROC
6980    That_Erl = Erl
6990    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7000    Resume EXITP

End Sub

Private Sub curr_id_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7120    intShiftDown = (Shift And acShiftMask) > 0
7130    intAltDown = (Shift And acAltMask) > 0
7140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7170        With Me
7180          intRetVal = 0
7190          If .yield_entry.Enabled = True Then
7200            .yield_entry.SetFocus
7210          ElseIf .marketvalue.Enabled = True Then
7220            .marketvalue.SetFocus
7230          Else
7240            lngRecsCur = RecCnt  ' ** Function: Below.
7250            If .CurrentRecord < lngRecsCur Then
7260              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
7270              .FocusHolder.SetFocus
7280            Else
7290              DoCmd.SelectObject acForm, .Parent.Name, False
7300              .Parent.opgShow.SetFocus
7310            End If
7320          End If
7330        End With
7340      End Select
7350    End If

        ' ** Shift keys.
7360    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7370      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7380        With Me
7390          intRetVal = 0
7400          .curr_id.SetFocus
7410          If .CurrentRecord > 1 Then
7420            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
7430            If .marketvalue.Enabled = True Then
7440              .marketvalue.SetFocus
7450            Else
7460              .FocusHolder.SetFocus
7470            End If
7480          Else
7490            DoCmd.SelectObject acForm, .Parent.Name, False
7500            .Parent.cmdClose.SetFocus
7510          End If
7520        End With
7530      End Select
7540    End If

EXITP:
7550    KeyCode = intRetVal
7560    Exit Sub

ERRH:
7570    intRetVal = 0
7580    THAT_PROC = THIS_PROC
7590    That_Erl = Erl
7600    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7610    Resume EXITP

End Sub

Private Sub yield_entry_lbl_DblClick(Cancel As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "yield_entry_lbl_DblClick"

7710    SortNow THIS_PROC

EXITP:
7720    Exit Sub

ERRH:
7730    THAT_PROC = THIS_PROC
7740    That_Erl = Erl
7750    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7760    Resume EXITP

End Sub

Private Sub yield_entry_Change()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "yield_entry_Change"

7810    With Me
7820      If .Parent.cmdUpdate.Enabled = False Then
7830        .Parent.cmdUpdate.Enabled = True
7840      End If
7850      If .Parent.cmdReset.Enabled = False Then
7860        .Parent.cmdReset.Enabled = True
7870      End If
7880    End With

EXITP:
7890    Exit Sub

ERRH:
7900    THAT_PROC = THIS_PROC
7910    That_Erl = Erl
7920    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7930    Resume EXITP

End Sub

Private Sub yield_entry_AfterUpdate()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "yield_entry_AfterUpdate"

8010    With Me
8020      .Parent.chkUpdated = True
8030      .AssetPricing_Changed = True
8040      If .Parent.cmdUpdate.Enabled = False Then
8050        .Parent.cmdUpdate.Enabled = True
8060      End If
8070      If .Parent.cmdReset.Enabled = False Then
8080        .Parent.cmdReset.Enabled = True
8090      End If
8100    End With

EXITP:
8110    Exit Sub

ERRH:
8120    THAT_PROC = THIS_PROC
8130    That_Erl = Erl
8140    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8150    Resume EXITP

End Sub

Private Sub yield_entry_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "yield_entry_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8220    intShiftDown = (Shift And acShiftMask) > 0
8230    intAltDown = (Shift And acAltMask) > 0
8240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8270        With Me
8280          intRetVal = 0
8290          If .marketvalue.Enabled = True Then
8300            .marketvalue.SetFocus
8310          Else
8320            lngRecsCur = RecCnt  ' ** Function: Below.
8330            If .CurrentRecord < lngRecsCur Then
8340              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
8350              .FocusHolder.SetFocus
8360            Else
8370              DoCmd.SelectObject acForm, .Parent.Name, False
8380              .Parent.opgShow.SetFocus
8390            End If
8400          End If
8410        End With
8420      End Select
8430    End If

        ' ** Shift keys.
8440    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8450      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8460        With Me
8470          intRetVal = 0
8480          If .curr_id.Visible = True And .curr_id.Enabled = True Then
8490            .curr_id.SetFocus
8500          Else
8510            If .CurrentRecord > 1 Then
8520              MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
8530              If .marketvalue.Enabled = True Then
8540                .marketvalue.SetFocus
8550              Else
8560                .FocusHolder.SetFocus
8570              End If
8580            Else
8590              DoCmd.SelectObject acForm, .Parent.Name, False
8600              .Parent.cmdClose.SetFocus
8610            End If
8620          End If
8630        End With
8640      End Select
8650    End If

EXITP:
8660    KeyCode = intRetVal
8670    Exit Sub

ERRH:
8680    intRetVal = 0
8690    THAT_PROC = THIS_PROC
8700    That_Erl = Erl
8710    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8720    Resume EXITP

End Sub

Private Sub yield_entry_Exit(Cancel As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "yield_entry_Exit"

        Dim intPos01 As Integer
        Dim strTmp01 As String, dblTmp02 As Double
        Dim blnContinue As Boolean

8810    blnContinue = True

8820    With Me

8830      If gblnMessage = False Then

8840        If IsNull(.yield_entry) = False Then
8850          strTmp01 = CStr(.yield_entry)
8860          strTmp01 = Trim(strTmp01)
8870          If Right(strTmp01, 1) = "%" Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
8880          If IsNumeric(strTmp01) = True Then
8890            dblTmp02 = CDbl(strTmp01)
8900            If dblTmp02 >= 100# Then
                  ' ** I'm not sure that we care where Yield is concerned.
8910            End If
8920            intPos01 = InStr(strTmp01, ".")
8930            If intPos01 > 0 Then
                  ' ** The user entered a decimal number.
8940              If Len(Mid(.yield_entry, (intPos01 + 1))) > 4 Then  ' ** It was 6, but now we're entering it as a percentage.
8950                blnContinue = False
8960                DoCmd.Hourglass False
8970                MsgBox "You can only enter a yield percentage with up to 4 decimal places.", vbInformation + vbOKOnly, "Invalid Entry"
8980                Cancel = -1
8990              Else
9000                If dblTmp02 <> 0# Then
9010                  If (dblTmp02 / 100#) <> .yield Then
9020                    .yield = dblTmp02 / 100#  ' ** Saved as a percentage, i.e., must be less than 1!
9030                    .AssetPricing_Changed = True
9040                    cmdSave_Click  ' ** Procedure: Above.
9050                  End If
9060                Else
9070                  If .yield <> 0# Then
9080                    .yield = 0#
9090                    .AssetPricing_Changed = True
9100                    cmdSave_Click  ' ** Procedure: Above.
9110                  End If
9120                End If
9130              End If
9140            Else
                  ' ** A whole number, with no decimals.
9150              .yield = .yield_entry
9160              .AssetPricing_Changed = True
9170              cmdSave_Click  ' ** Procedure: Above.
9180            End If
9190          Else
9200            DoCmd.Hourglass False
9210            MsgBox "Please enter a valid yield percentage.", vbInformation + vbOKOnly, "Invalid Entry"
9220            Cancel = -1
9230          End If
9240        Else
9250          .yield_entry = 0#
9260          .yield = 0#
9270          .AssetPricing_Changed = True
9280          cmdSave_Click  ' ** Procedure: Above.
9290        End If

9300        If blnContinue = True Then
9310          Select Case IsNull(.yield_entry)
              Case True
9320            .yield_entry = 0#
9330            .yield = 0#
9340            .AssetPricing_Changed = True
9350            DoEvents
9360            cmdSave_Click  ' ** Procedure: Above.
9370          Case False
9380            If IsNumeric(.yield_entry) = False Then
9390              .yield_entry = 0#
9400              .yield = 0#
9410              .AssetPricing_Changed = True
9420              cmdSave_Click  ' ** Procedure: Above.
9430            End If
9440          End Select
9450        End If

9460        DoCmd.Hourglass False

9470      End If  ' ** gblnMessage.

9480    End With  ' ** Me.

EXITP:
9490    Exit Sub

ERRH:
9500    DoCmd.Hourglass False
9510    THAT_PROC = THIS_PROC
9520    That_Erl = Erl
9530    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9540    Resume EXITP

End Sub

Private Sub marketvaluecurrent_str_lbl_DblClick(Cancel As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvaluecurrent_str_lbl_DblClick"

9610    SortNow THIS_PROC

EXITP:
9620    Exit Sub

ERRH:
9630    THAT_PROC = THIS_PROC
9640    That_Erl = Erl
9650    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9660    Resume EXITP

End Sub

Private Sub marketvaluecurrent_usd_str_lbl_DblClick(Cancel As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvaluecurrent_usd_str_lbl_DblClick"

9710    SortNow THIS_PROC

EXITP:
9720    Exit Sub

ERRH:
9730    THAT_PROC = THIS_PROC
9740    That_Erl = Erl
9750    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9760    Resume EXITP

End Sub

Private Sub marketvalue_lbl_DblClick(Cancel As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvalue_lbl_DblClick"

9810    SortNow THIS_PROC

EXITP:
9820    Exit Sub

ERRH:
9830    THAT_PROC = THIS_PROC
9840    That_Erl = Erl
9850    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9860    Resume EXITP

End Sub

Private Sub marketvalue_Enter()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "marketvalue_Enter"

9910    With Me

9920      If .curr_id = 150& Then
9930        If .marketvalue.Enabled = True And .marketvalue.Locked = False Then
9940          .marketvalue.BackColor = CLR_WHT
9950        Else
9960          .marketvalue.BackColor = CLR_DISABLED_BG
9970        End If
9980      Else
9990        .marketvalue.BackColor = CLR_VLTGRN
10000     End If

10010   End With

EXITP:
10020   Exit Sub

ERRH:
10030   THAT_PROC = THIS_PROC
10040   That_Erl = Erl
10050   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10060   Resume EXITP

End Sub

Private Sub marketvalue_Change()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "marketvalue_Change"

10110   With Me
10120     If .Parent.cmdUpdate.Enabled = False Then
10130       .Parent.cmdUpdate.Enabled = True
10140     End If
10150     If .Parent.cmdReset.Enabled = False Then
10160       .Parent.cmdReset.Enabled = True
10170     End If
10180   End With

EXITP:
10190   Exit Sub

ERRH:
10200   THAT_PROC = THIS_PROC
10210   That_Erl = Erl
10220   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10230   Resume EXITP

End Sub

Private Sub marketvalue_AfterUpdate()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "marketvalue_AfterUpdate"

10310   With Me
10320     .Parent.chkUpdated = True
10330     .AssetPricing_Changed = True
10340     If .Parent.cmdUpdate.Enabled = False Then
10350       .Parent.cmdUpdate.Enabled = True
10360     End If
10370     If .Parent.cmdReset.Enabled = False Then
10380       .Parent.cmdReset.Enabled = True
10390     End If
10400     Select Case IsNull(.marketvalue)
          Case True
10410       If IsNull(.marketvalue_usd) = False Then
10420         .marketvalue_usd = Null
10430       End If
10440     Case False
10450       If .curr_code = "USD" Then
10460         .marketvalue_usd = .marketvalue
10470       Else
10480         .marketvalue_usd = (.marketvalue * .curr_rate2)
10490       End If
10500     End Select
10510   End With

EXITP:
10520   Exit Sub

ERRH:
10530   THAT_PROC = THIS_PROC
10540   That_Erl = Erl
10550   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10560   Resume EXITP

End Sub

Private Sub marketvalue_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "marketvalue_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10620   intShiftDown = (Shift And acShiftMask) > 0
10630   intAltDown = (Shift And acAltMask) > 0
10640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10670       With Me
10680         intRetVal = 0
10690         lngRecsCur = RecCnt  ' ** Function: Below.
10700         If .CurrentRecord < lngRecsCur Then
10710           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
10720           If .curr_id.Visible = True And .curr_id.Enabled = True Then
10730             .curr_id.SetFocus
10740           ElseIf .yield_entry.Enabled = True Then
10750             .yield_entry.SetFocus
10760           Else
10770             .FocusHolder.SetFocus
10780           End If
10790         Else
10800           DoCmd.SelectObject acForm, .Parent.Name, False
10810           .Parent.opgShow.SetFocus
10820         End If
10830       End With
10840     End Select
10850   End If

        ' ** Shift keys.
10860   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10870     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10880       With Me
10890         intRetVal = 0
10900         If .yield_entry.Enabled = True Then
10910           .yield_entry.SetFocus
10920         ElseIf .curr_id.Visible = True And .curr_id.Enabled = True Then
10930           .curr_id.SetFocus
10940         Else
10950           If .CurrentRecord > 1 Then
10960             MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
10970             .FocusHolder.SetFocus
10980           Else
10990             DoCmd.SelectObject acForm, .Parent.Name, False
11000             .Parent.cmdClose.SetFocus
11010           End If
11020         End If
11030       End With
11040     End Select
11050   End If

EXITP:
11060   KeyCode = intRetVal
11070   Exit Sub

ERRH:
11080   intRetVal = 0
11090   THAT_PROC = THIS_PROC
11100   That_Erl = Erl
11110   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11120   Resume EXITP

End Sub

Private Sub marketvalue_Exit(Cancel As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "marketvalue_Exit"

        Dim intPos01 As Integer
        Dim blnRetVal As Boolean

11210   With Me
11220     blnRetVal = True
11230     Select Case IsNull(.marketvalue)
          Case True
11240       blnRetVal = False
11250     Case False
11260       intPos01 = InStr(1, .marketvalue, ".", vbTextCompare)
11270       If intPos01 > 0 Then
              ' ** The user entered a decimal number.
11280         If Len(Mid(.marketvalue, intPos01 + 1, Len(.marketvalue))) > 5 Then
11290           blnRetVal = False
11300           DoCmd.Hourglass False
11310           MsgBox "You can only enter a market value with up to 5 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
11320           .marketvalue = vbNullString
11330           DoCmd.CancelEvent
11340         End If
11350       End If
11360       If Val(Nz(.marketvalue, 0)) < 0 Then
11370         blnRetVal = False
11380         DoCmd.Hourglass False
11390         MsgBox "You cannot enter a negative market value.", vbInformation + vbOKOnly, "Invalid Entry"
11400         .marketvalue = vbNullString
11410         DoCmd.CancelEvent
11420       End If
11430     End Select
11440   End With

EXITP:
11450   Exit Sub

ERRH:
11460   DoCmd.Hourglass False
11470   THAT_PROC = THIS_PROC
11480   That_Erl = Erl
11490   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11500   Resume EXITP

End Sub

Private Sub marketvalue_usd_lbl_DblClick(Cancel As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "marketvalue_usd_lbl_DblClick"

11610   SortNow THIS_PROC

EXITP:
11620   Exit Sub

ERRH:
11630   THAT_PROC = THIS_PROC
11640   That_Erl = Erl
11650   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11660   Resume EXITP

End Sub

Private Sub cmdShorten_GotFocus()

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_GotFocus"

11710   With Me
11720     blnShorten_Focus = True
11730     .cmdShorten_raised_semifocus_dots_img.Visible = True
11740     .cmdShorten_raised_img.Visible = False
11750     .cmdShorten_raised_focus_img.Visible = False
11760     .cmdShorten_raised_focus_dots_img.Visible = False
11770     .cmdShorten_sunken_focus_dots_img.Visible = False
11780     .cmdShorten_raised_img_dis.Visible = False
11790   End With

EXITP:
11800   Exit Sub

ERRH:
11810   Select Case ERR.Number
        Case Else
11820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11830   End Select
11840   Resume EXITP

End Sub

Private Sub cmdShorten_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseDown"

11910   With Me
11920     blnShorten_MouseDown = True
11930     .cmdShorten_sunken_focus_dots_img.Visible = True
11940     .cmdShorten_raised_img.Visible = False
11950     .cmdShorten_raised_semifocus_dots_img.Visible = False
11960     .cmdShorten_raised_focus_img.Visible = False
11970     .cmdShorten_raised_focus_dots_img.Visible = False
11980     .cmdShorten_raised_img_dis.Visible = False
11990   End With

EXITP:
12000   Exit Sub

ERRH:
12010   Select Case ERR.Number
        Case Else
12020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12030   End Select
12040   Resume EXITP

End Sub

Public Sub cmdShorten_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_Click"

        Dim lngMinusedWidth As Long
        Dim lngDiff1 As Long, lngDiff2 As Long
        Dim blnSortHere As Boolean, blnHasCurr As Boolean

12110   With Me
12120     blnHasCurr = .Parent.chkIncludeCurrency
12130     lngMinusedWidth = lngTpp * WIDEN_MULT
12140     If (.totdesc.Width <= IIf(.Parent.chkShowAssetNo = True, lngTotDescAstno_Width, lngTotDesc_Width)) Then
12150       Beep
12160     ElseIf ((.totdesc.Width - lngMinusedWidth) < IIf(.Parent.chkShowAssetNo = True, lngTotDescAstno_Width, lngTotDesc_Width)) Then
12170       blnSortHere = False
12180       If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
12190       .totdesc.Width = IIf(.Parent.chkShowAssetNo = True, lngTotDescAstno_Width, lngTotDesc_Width)
12200       .totdesc_lbl.Width = IIf(.Parent.chkShowAssetNo = True, lngTotDescAstno_Width, lngTotDesc_Width)
12210       .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
12220       If blnSortHere = True Then
12230         .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
12240         .Sort_lbl.Left = (((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width) + (2& * lngTpp))
12250       End If
12260       blnSortHere = False
12270       Select Case blnHasCurr
            Case True
12280         lngDiff1 = (.curr_id.Left - .assettype.Left)
12290         If .Sort_line.Left = .assettype_lbl.Left Then blnSortHere = True
12300         .assettype.Left = lngAssetNext_Left
12310         .assettype_lbl.Left = lngAssetNext_Left
12320         .assettype_lbl_line.Left = .assettype_lbl.Left
12330         If blnSortHere = True Then
12340           .Sort_line.Left = .assettype_lbl.Left
12350           .Sort_lbl.Left = (((.assettype_lbl.Left + .assettype_lbl.Width) - lngSortLbl_Width) + (5& * lngTpp))
12360         End If
12370         blnSortHere = False
12380         lngDiff2 = (.yield_entry.Left - .curr_id.Left)
12390         If .Sort_line.Left = .curr_id.Left Then blnSortHere = True
12400         .curr_id.Left = (.assettype.Left + lngDiff1)
12410         .curr_id_usd_box.Left = (.curr_id.Left - lngTpp)
12420         .curr_id_forex_box.Left = (.curr_id.Left - lngTpp)
12430         .curr_id_lbl.Left = (.curr_id.Left - (4& * lngTpp))
12440         .curr_id_lbl_line.Left = .curr_id.Left
12450         .curr_date.Left = .curr_id_lbl_line.Left
12460         If blnSortHere = True Then
12470           .Sort_line.Left = .curr_id.Left
12480           .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (3& * lngTpp))
12490         End If
12500         blnSortHere = False
12510         lngDiff1 = (.marketvaluecurrent_str.Left - .yield_entry.Left)
12520         If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
12530         .yield_entry.Left = (.curr_id.Left + lngDiff2)
12540         .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
12550         .yield_box.Left = .yield_entry.Left
12560         .yield_upd.Left = (.yield_entry.Left - lngTpp)
12570         .yield_entry_lbl.Left = .yield_entry.Left
12580         .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
12590         If blnSortHere = True Then
12600           .Sort_line.Left = .yield_entry_lbl.Left
12610           .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
12620         End If
12630         blnSortHere = False
12640         lngDiff2 = (.marketvaluecurrent_usd_str.Left - .marketvaluecurrent_str.Left)
12650         If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
12660         .marketvaluecurrent_str.Left = (.yield_entry.Left + lngDiff1)
12670         .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
12680         .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
12690         .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
12700         .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
12710         .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
12720         .marketvaluecurrent_lbl_line.Left = .marketvaluecurrent_lbl.Left
12730         If blnSortHere = True Then
12740           .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
12750           .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
12760         End If
12770         blnSortHere = False
12780         lngDiff1 = (.currsym_symbol.Left - .marketvaluecurrent_usd_str.Left)
12790         If .Sort_line.Left = .marketvaluecurrent_usd_str_lbl.Left Then blnSortHere = True
12800         .marketvaluecurrent_usd_str.Left = (.marketvaluecurrent_str.Left + lngDiff2)
12810         .marketvaluecurrent_usd_str_lbl.Left = .marketvaluecurrent_usd_str.Left
12820         .marketvaluecurrent_usd_str_lbl_line.Left = .marketvaluecurrent_usd_str_lbl.Left
12830         If blnSortHere = True Then
12840           .Sort_line.Left = .marketvaluecurrent_usd_str_lbl.Left
12850           .Sort_lbl.Left = ((.marketvaluecurrent_usd_str_lbl.Left + .marketvaluecurrent_usd_str_lbl.Width) - lngSortLbl_Width)
12860         End If
12870         blnSortHere = False
12880         lngDiff2 = (.currsym_symbol_usd.Left - .currsym_symbol.Left)
12890         If .Sort_line.Left = .marketvalue_lbl.Left Then blnSortHere = True
12900         .currsym_symbol.Left = (.marketvaluecurrent_usd_str.Left + lngDiff1)
12910         .currsym_symbol_loc_usd_box.Left = (.currsym_symbol.Left - lngTpp)
12920         .currsym_symbol_loc_forex_box.Left = (.currsym_symbol.Left - lngTpp)
12930         .marketvalue.Left = (.currsym_symbol.Left + lngMktValSym_Offset)
12940         .marketvalue_loc_usd_box.Left = (.marketvalue.Left - lngTpp)
12950         .marketvalue_loc_forex_box.Left = (.marketvalue.Left - lngTpp)
12960         .marketvalue_box.Left = .currsym_symbol.Left
12970         .marketvalue_upd.Left = (.currsym_symbol.Left - lngTpp)
12980         .marketvalue_lbl.Left = .currsym_symbol.Left
12990         .marketvalue_lbl_line.Left = .marketvalue_lbl.Left
13000         .marketvaluenew_lbl.Left = .currsym_symbol.Left
13010         .marketvaluenew_lbl_line.Left = .marketvaluenew_lbl.Left
13020         If blnSortHere = True Then
13030           .Sort_line.Left = .currsym_symbol.Left
13040           .Sort_lbl.Left = ((.marketvalue_lbl.Left + .marketvalue_lbl.Width) - lngSortLbl_Width)
13050         End If
13060         blnSortHere = False
13070         If .Sort_line.Left = .marketvalue_usd_lbl.Left Then blnSortHere = True
13080         .currsym_symbol_usd.Left = (.currsym_symbol.Left + lngDiff2)
13090         .marketvalue_usd = (.currsym_symbol_usd.Left + lngMktValSym_Offset)
13100         .marketvalue_usd_box.Left = .currsym_symbol_usd.Left
13110         .marketvalue_usd_upd.Left = (.currsym_symbol_usd.Left - lngTpp)
13120         .marketvalue_usd_lbl.Left = .currsym_symbol_usd.Left
13130         .marketvalue_usd_lbl_line.Left = .marketvalue_usd_lbl.Left
13140         If blnSortHere = True Then
13150           .Sort_line.Left = .marketvalue_usd_lbl.Left
13160           .Sort_lbl.Left = ((.marketvalue_usd_lbl.Left + .marketvalue_usd_lbl.Width) - lngSortLbl_Width)
13170         End If
13180         .Width = lngForm_Width
13190       Case False
13200         lngDiff1 = (.yield_entry.Left - .assettype.Left)
13210         If .Sort_line.Left = .assettype_lbl.Left Then blnSortHere = True
13220         .assettype.Left = lngAssetNext_Left
13230         .assettype_lbl.Left = lngAssetNext_Left
13240         .assettype_lbl_line.Left = .assettype_lbl.Left
13250         If blnSortHere = True Then
13260           .Sort_line.Left = .assettype_lbl.Left
13270           .Sort_lbl.Left = (((.assettype_lbl.Left + .assettype_lbl.Width) - lngSortLbl_Width) + (5& * lngTpp))
13280         End If
13290         blnSortHere = False
13300         lngDiff2 = (.marketvaluecurrent_str.Left - .yield_entry.Left)
13310         If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
13320         .yield_entry.Left = (.assettype.Left + lngDiff1)
13330         .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
13340         .yield_box.Left = .yield_entry.Left
13350         .yield_upd.Left = (.yield_entry.Left - lngTpp)
13360         .yield_entry_lbl.Left = .yield_entry.Left
13370         .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
13380         If blnSortHere = True Then
13390           .Sort_line.Left = .yield_entry_lbl.Left
13400           .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
13410         End If
13420         blnSortHere = False
13430         lngDiff1 = (.currsym_symbol.Left - .marketvaluecurrent_str.Left)
13440         If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
13450         .marketvaluecurrent_str.Left = (.yield_entry.Left + lngDiff2)
13460         .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
13470         .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
13480         .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
13490         .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
13500         .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
13510         If blnSortHere = True Then
13520           .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
13530           .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
13540         End If
13550         blnSortHere = False
13560         If .Sort_line.Left = .marketvalue_lbl.Left Then blnSortHere = True
13570         .currsym_symbol.Left = (.marketvaluecurrent_str.Left + lngDiff1)
13580         .currsym_symbol_loc_usd_box.Left = (.currsym_symbol.Left - lngTpp)
13590         .currsym_symbol_loc_forex_box.Left = (.currsym_symbol.Left - lngTpp)
13600         .marketvalue.Left = (.currsym_symbol.Left + lngMktValSym_Offset)
13610         .marketvalue_loc_usd_box.Left = (.marketvalue.Left - lngTpp)
13620         .marketvalue_loc_forex_box.Left = (.marketvalue.Left - lngTpp)
13630         .marketvalue_box.Left = .currsym_symbol.Left
13640         .marketvalue_upd.Left = (.currsym_symbol.Left - lngTpp)
13650         .marketvalue_lbl.Left = .currsym_symbol.Left
13660         .marketvalue_lbl_line.Left = .marketvalue_lbl.Left
13670         .marketvaluenew_lbl.Left = .currsym_symbol.Left
13680         If blnSortHere = True Then
13690           .Sort_line.Left = .currsym_symbol.Left
13700           .Sort_lbl.Left = ((.marketvalue_lbl.Left + .marketvalue_lbl.Width) - lngSortLbl_Width)
13710         End If
13720         blnSortHere = False
13730         .Width = (lngForm_Width - ((.curr_id.Width + (4& * lngTpp)) + (.marketvaluecurrent_usd_str_lbl.Width + (4& * lngTpp)) + _
                (.marketvalue_usd_lbl.Width + (4& * lngTpp))))
13740       End Select
13750     Else
13760       blnSortHere = False
13770       If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
13780       .totdesc.Width = (.totdesc.Width - lngMinusedWidth)
13790       .totdesc_lbl.Width = .totdesc.Width
13800       .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
13810       If blnSortHere = True Then
13820         .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
13830         .Sort_lbl.Left = (((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width) + (2& * lngTpp))
13840       End If
13850       blnSortHere = False
13860       If .Sort_line.Left = .assettype_lbl.Left Then blnSortHere = True
13870       .assettype.Left = (.assettype.Left - lngMinusedWidth)
13880       .assettype_lbl.Left = .assettype.Left
13890       .assettype_lbl_line.Left = .assettype_lbl.Left
13900       If blnSortHere = True Then
13910         .Sort_line.Left = .assettype_lbl.Left
13920         .Sort_lbl.Left = (((.assettype_lbl.Left + .assettype_lbl.Width) - lngSortLbl_Width) + (5& * lngTpp))
13930       End If
13940       blnSortHere = False
13950       If blnHasCurr = True Then
13960         If .Sort_line.Left = .curr_id.Left Then blnSortHere = True
13970         .curr_id.Left = (.curr_id.Left - lngMinusedWidth)
13980         .curr_id_usd_box.Left = (.curr_id.Left - lngTpp)
13990         .curr_id_forex_box.Left = (.curr_id.Left - lngTpp)
14000         .curr_id_lbl.Left = (.curr_id.Left - (4& * lngTpp))
14010         .curr_id_lbl_line.Left = curr_id.Left
14020         .curr_date.Left = .curr_id_lbl_line.Left
14030         If blnSortHere = True Then
14040           .Sort_line.Left = .curr_id.Left
14050           .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (3& * lngTpp))
14060         End If
14070         blnSortHere = False
14080       End If
14090       If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
14100       .yield_entry.Left = (.yield_entry.Left - lngMinusedWidth)
14110       .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
14120       .yield_box.Left = .yield_entry.Left
14130       .yield_upd.Left = (.yield_entry.Left - lngTpp)
14140       .yield_entry_lbl.Left = .yield_entry.Left
14150       .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
14160       If blnSortHere = True Then
14170         .Sort_line.Left = .yield_entry_lbl.Left
14180         .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
14190       End If
14200       blnSortHere = False
14210       If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
14220       .marketvaluecurrent_str.Left = (.marketvaluecurrent_str.Left - lngMinusedWidth)
14230       .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
14240       .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
14250       .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
14260       .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
14270       .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
14280       .marketvaluecurrent_lbl_line.Left = .marketvaluecurrent_lbl.Left
14290       If blnSortHere = True Then
14300         .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
14310         .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
14320       End If
14330       blnSortHere = False
14340       If blnHasCurr = True Then
14350         If .Sort_line.Left = .marketvaluecurrent_usd_str_lbl.Left Then blnSortHere = True
14360         .marketvaluecurrent_usd_str.Left = (.marketvaluecurrent_usd_str.Left - lngMinusedWidth)
14370         .marketvaluecurrent_usd_str_lbl.Left = .marketvaluecurrent_usd_str.Left
14380         .marketvaluecurrent_usd_str_lbl_line.Left = .marketvaluecurrent_usd_str_lbl.Left
14390         If blnSortHere = True Then
14400           .Sort_line.Left = .marketvaluecurrent_usd_str_lbl.Left
14410           .Sort_lbl.Left = ((.marketvaluecurrent_usd_str_lbl.Left + .marketvaluecurrent_usd_str_lbl.Width) - lngSortLbl_Width)
14420         End If
14430         blnSortHere = False
14440       End If
14450       If .Sort_line.Left = .marketvalue_lbl.Left Then blnSortHere = True
14460       .currsym_symbol.Left = (.currsym_symbol.Left - lngMinusedWidth)
14470       .currsym_symbol_loc_usd_box.Left = (.currsym_symbol.Left - lngTpp)
14480       .currsym_symbol_loc_forex_box.Left = (.currsym_symbol.Left - lngTpp)
14490       .marketvalue.Left = (.currsym_symbol.Left + lngMktValSym_Offset)
14500       .marketvalue_loc_usd_box.Left = (.marketvalue.Left - lngTpp)
14510       .marketvalue_loc_forex_box.Left = (.marketvalue.Left - lngTpp)
14520       .marketvalue_box.Left = .currsym_symbol.Left
14530       .marketvalue_upd.Left = (.currsym_symbol.Left - lngTpp)
14540       .marketvalue_lbl.Left = .currsym_symbol.Left
14550       .marketvalue_lbl_line.Left = .marketvalue_lbl.Left
14560       .marketvaluenew_lbl.Left = .currsym_symbol.Left
14570       If blnHasCurr = True Then
14580         .marketvaluenew_lbl_line.Left = .marketvaluenew_lbl.Left
14590       End If
14600       If blnSortHere = True Then
14610         .Sort_line.Left = .marketvalue_lbl.Left
14620         .Sort_lbl.Left = ((.marketvalue_lbl.Left + .marketvalue_lbl.Width) - lngSortLbl_Width)
14630       End If
14640       blnSortHere = False
14650       If blnHasCurr = True Then
14660         If .Sort_line.Left = .marketvalue_usd_lbl.Left Then blnSortHere = True
14670         .currsym_symbol_usd.Left = (.currsym_symbol_usd.Left - lngMinusedWidth)
14680         .marketvalue_usd.Left = (.currsym_symbol_usd.Left + lngMktValSym_Offset)
14690         .marketvalue_usd_box.Left = .currsym_symbol_usd.Left
14700         .marketvalue_usd_upd.Left = (.currsym_symbol_usd.Left - lngTpp)
14710         .marketvalue_usd_lbl.Left = .currsym_symbol_usd.Left
14720         .marketvalue_usd_lbl_line.Left = .marketvalue_usd_lbl.Left
14730         If blnSortHere = True Then
14740           .Sort_line.Left = .marketvalue_usd_lbl.Left
14750           .Sort_lbl.Left = ((.marketvalue_usd_lbl.Left + .marketvalue_usd_lbl.Width) - lngSortLbl_Width)
14760         End If
14770         blnSortHere = False
14780       End If
14790       .Width = (.Width - lngMinusedWidth)
14800     End If
14810     .FocusHolder.SetFocus
14820     If .cmdWiden.Enabled = False Then
            ' ** Since cmdShorten was clicked, cmdWiden is certainly available.
14830       .cmdWiden.Enabled = True
14840       .cmdWiden_raised_img.Visible = True
14850       .cmdWiden_raised_semifocus_dots_img.Visible = False
14860       .cmdWiden_raised_focus_img.Visible = False
14870       .cmdWiden_raised_focus_dots_img.Visible = False
14880       .cmdWiden_sunken_focus_dots_img.Visible = False
14890       .cmdWiden_raised_img_dis.Visible = False
14900     End If
          ' ** This doesn't trigger when clicking, only when using shortcut!
14910     If (.totdesc.Width <= IIf(.Parent.chkShowAssetNo = True, lngTotDescAstno_Width, lngTotDesc_Width)) Then
            ' ** If there's nowhere to go, disable it.
14920       .FocusHolder2.SetFocus
14930       .cmdShorten.Enabled = False
14940       .cmdShorten_raised_img_dis.Visible = True
14950       .cmdShorten_raised_img.Visible = False
14960       .cmdShorten_raised_semifocus_dots_img.Visible = False
14970       .cmdShorten_raised_focus_img.Visible = False
14980       .cmdShorten_raised_focus_dots_img.Visible = False
14990       .cmdShorten_sunken_focus_dots_img.Visible = False
15000       blnFromTimer = True  ' ** I know this is backwards.
15010       .TimerInterval = 250&
15020     End If
15030   End With

EXITP:
15040   Exit Sub

ERRH:
15050   Select Case ERR.Number
        Case Else
15060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15070   End Select
15080   Resume EXITP

End Sub

Private Sub cmdShorten_KeyDown(KeyCode As Integer, Shift As Integer)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15120   intShiftDown = (Shift And acShiftMask) > 0
15130   intAltDown = (Shift And acAltMask) > 0
15140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15160     Select Case intRetVal
          Case vbKeyTab
15170       With Me
15180         intRetVal = 0
15190         .cmdWiden.SetFocus
15200       End With
15210     End Select
15220   End If

        ' ** Shift keys.
15230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15240     Select Case intRetVal
          Case vbKeyTab
15250       With Me
15260         intRetVal = 0
15270         If .marketvalue.Enabled = True Then
15280           .marketvalue.SetFocus
15290         Else
15300           .FocusHolder.SetFocus
15310         End If
15320       End With
15330     End Select
15340   End If

EXITP:
15350   KeyCode = intRetVal
15360   Exit Sub

ERRH:
15370   intRetVal = 0
15380   THAT_PROC = THIS_PROC
15390   That_Erl = Erl
15400   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15410   Resume EXITP

End Sub

Private Sub cmdShorten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseMove"

15510   With Me
15520     If blnShorten_MouseDown = False Then
15530       Select Case blnShorten_Focus
            Case True
15540         .cmdShorten_raised_focus_dots_img.Visible = True
15550         .cmdShorten_raised_focus_img.Visible = False
15560       Case False
15570         .cmdShorten_raised_focus_img.Visible = True
15580         .cmdShorten_raised_focus_dots_img.Visible = False
15590       End Select
15600       .cmdShorten_raised_img.Visible = False
15610       .cmdShorten_raised_semifocus_dots_img.Visible = False
15620       .cmdShorten_sunken_focus_dots_img.Visible = False
15630       .cmdShorten_raised_img_dis.Visible = False
15640     End If
15650     If .cmdWiden_raised_focus_dots_img.Visible = True Or .cmdWiden_raised_focus_img.Visible = True Then
15660       Select Case blnWiden_Focus
            Case True
15670         .cmdWiden_raised_semifocus_dots_img.Visible = True
15680         .cmdWiden_raised_img.Visible = False
15690       Case False
15700         .cmdWiden_raised_img.Visible = True
15710         .cmdWiden_raised_semifocus_dots_img.Visible = False
15720       End Select
15730       .cmdWiden_raised_focus_img.Visible = False
15740       .cmdWiden_raised_focus_dots_img.Visible = False
15750       .cmdWiden_sunken_focus_dots_img.Visible = False
15760       .cmdWiden_raised_img_dis.Visible = False
15770     End If
15780   End With

EXITP:
15790   Exit Sub

ERRH:
15800   Select Case ERR.Number
        Case Else
15810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15820   End Select
15830   Resume EXITP

End Sub

Private Sub cmdShorten_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_MouseUp"

15910   With Me
15920     .cmdShorten_raised_focus_dots_img.Visible = True
15930     .cmdShorten_raised_img.Visible = False
15940     .cmdShorten_raised_semifocus_dots_img.Visible = False
15950     .cmdShorten_raised_focus_img.Visible = False
15960     .cmdShorten_sunken_focus_dots_img.Visible = False
15970     .cmdShorten_raised_img_dis.Visible = False
15980     blnShorten_MouseDown = False
15990   End With

EXITP:
16000   Exit Sub

ERRH:
16010   Select Case ERR.Number
        Case Else
16020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16030   End Select
16040   Resume EXITP

End Sub

Private Sub cmdShorten_LostFocus()

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShorten_LostFocus"

16110   With Me
16120     .cmdShorten_raised_img.Visible = True
16130     .cmdShorten_raised_semifocus_dots_img.Visible = False
16140     .cmdShorten_raised_focus_img.Visible = False
16150     .cmdShorten_raised_focus_dots_img.Visible = False
16160     .cmdShorten_sunken_focus_dots_img.Visible = False
16170     .cmdShorten_raised_img_dis.Visible = False
16180     blnShorten_Focus = False
16190   End With

EXITP:
16200   Exit Sub

ERRH:
16210   Select Case ERR.Number
        Case Else
16220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16230   End Select
16240   Resume EXITP

End Sub

Private Sub cmdWiden_GotFocus()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_GotFocus"

16310   With Me
16320     blnWiden_Focus = True
16330     .cmdWiden_raised_semifocus_dots_img.Visible = True
16340     .cmdWiden_raised_img.Visible = False
16350     .cmdWiden_raised_focus_img.Visible = False
16360     .cmdWiden_raised_focus_dots_img.Visible = False
16370     .cmdWiden_sunken_focus_dots_img.Visible = False
16380     .cmdWiden_raised_img_dis.Visible = False
16390   End With

EXITP:
16400   Exit Sub

ERRH:
16410   Select Case ERR.Number
        Case Else
16420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16430   End Select
16440   Resume EXITP

End Sub

Private Sub cmdWiden_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseDown"

16510   With Me
16520     blnWiden_MouseDown = True
16530     .cmdWiden_sunken_focus_dots_img.Visible = True
16540     .cmdWiden_raised_img.Visible = False
16550     .cmdWiden_raised_semifocus_dots_img.Visible = False
16560     .cmdWiden_raised_focus_img.Visible = False
16570     .cmdWiden_raised_focus_dots_img.Visible = False
16580     .cmdWiden_raised_img_dis.Visible = False
16590   End With

EXITP:
16600   Exit Sub

ERRH:
16610   Select Case ERR.Number
        Case Else
16620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16630   End Select
16640   Resume EXITP

End Sub

Public Sub cmdWiden_Click()
' ** lngTpp * 4 = 60  : 1 grid point on form (my screen is 15 Twips Per Pixel).
' ** 60 * 4 = 240     : 4 grid points, a nice noticeable amount.
' ** 1440 Twips per Inch.
' ** Arbitrarily cap it at 10 in. wide.

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_Click"

        Dim lngAddedWidth As Long
        Dim blnSortHere As Boolean, blnHasCurr As Boolean

16710   With Me
16720     blnHasCurr = .Parent.chkIncludeCurrency
16730     lngAddedWidth = lngTpp * WIDEN_MULT
16740     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
16750       Beep
16760     Else
16770       .Width = (.Width + lngAddedWidth)
16780       blnSortHere = False
16790       If .Sort_line.Left = .assettype_lbl.Left Then blnSortHere = True
16800       .assettype.Left = (.assettype.Left + lngAddedWidth)
16810       .assettype_lbl.Left = .assettype.Left
16820       .assettype_lbl_line.Left = .assettype_lbl.Left
16830       If blnSortHere = True Then
16840         .Sort_line.Left = .assettype_lbl.Left
16850         .Sort_lbl.Left = (((.assettype_lbl.Left + .assettype_lbl.Width) - lngSortLbl_Width) + (5& * lngTpp))
16860       End If
16870       blnSortHere = False
16880       If blnHasCurr = True Then
16890         If .Sort_line.Left = .curr_id.Left Then blnSortHere = True
16900         .curr_id.Left = (.curr_id.Left + lngAddedWidth)
16910         .curr_id_usd_box.Left = (.curr_id.Left - lngTpp)
16920         .curr_id_forex_box.Left = (.curr_id.Left - lngTpp)
16930         .curr_id_lbl.Left = (.curr_id.Left - (4& * lngTpp))
16940         .curr_id_lbl_line.Left = .curr_id.Left
16950         .curr_date.Left = .curr_id_lbl_line.Left
16960         If blnSortHere = True Then
16970           .Sort_line.Left = curr_id.Left
16980           .Sort_lbl.Left = (((.curr_id_lbl.Left + .curr_id_lbl.Width) - lngSortLbl_Width) + (3& * lngTpp))
16990         End If
17000         blnSortHere = False
17010       End If
17020       If .Sort_line.Left = .yield_entry_lbl.Left Then blnSortHere = True
17030       .yield_entry.Left = (.yield_entry.Left + lngAddedWidth)
17040       .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
17050       .yield_box.Left = .yield_entry.Left
17060       .yield_upd.Left = (.yield_entry.Left - lngTpp)
17070       .yield_entry_lbl.Left = .yield_entry.Left
17080       .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
17090       If blnSortHere = True Then
17100         .Sort_line.Left = .yield_entry_lbl.Left
17110         .Sort_lbl.Left = ((.yield_entry_lbl.Left + .yield_entry_lbl.Width) - lngSortLbl_Width)
17120       End If
17130       blnSortHere = False
17140       If .Sort_line.Left = .marketvaluecurrent_str_lbl.Left Then blnSortHere = True
17150       .marketvaluecurrent_str.Left = (.marketvaluecurrent_str.Left + lngAddedWidth)
17160       .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
17170       .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
17180       .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
17190       .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
17200       .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
17210       If blnHasCurr = True Then
17220         .marketvaluecurrent_lbl_line.Left = .marketvaluecurrent_lbl.Left
17230       End If
17240       If blnSortHere = True Then
17250         .Sort_line.Left = .marketvaluecurrent_str_lbl.Left
17260         .Sort_lbl.Left = ((.marketvaluecurrent_str_lbl.Left + .marketvaluecurrent_str_lbl.Width) - lngSortLbl_Width)
17270       End If
17280       blnSortHere = False
17290       If blnHasCurr = True Then
17300         If .Sort_line.Left = .marketvaluecurrent_usd_str_lbl.Left Then blnSortHere = True
17310         .marketvaluecurrent_usd_str.Left = (.marketvaluecurrent_usd_str.Left + lngAddedWidth)
17320         .marketvaluecurrent_usd_str_lbl.Left = .marketvaluecurrent_usd_str.Left
17330         .marketvaluecurrent_usd_str_lbl_line.Left = .marketvaluecurrent_usd_str_lbl.Left
17340         If blnSortHere = True Then
17350           .Sort_line.Left = .marketvaluecurrent_usd_str_lbl.Left
17360           .Sort_lbl.Left = ((.marketvaluecurrent_usd_str_lbl.Left + .marketvaluecurrent_usd_str_lbl.Width) - lngSortLbl_Width)
17370         End If
17380         blnSortHere = False
17390       End If
17400       If .Sort_line.Left = .marketvalue_lbl.Left Then blnSortHere = True
17410       .currsym_symbol.Left = (.currsym_symbol.Left + lngAddedWidth)
17420       .currsym_symbol_loc_usd_box.Left = (.currsym_symbol.Left - lngTpp)
17430       .currsym_symbol_loc_forex_box.Left = (.currsym_symbol.Left - lngTpp)
17440       .marketvalue.Left = (.currsym_symbol.Left + lngMktValSym_Offset)
17450       .marketvalue_loc_usd_box.Left = (.marketvalue.Left - lngTpp)
17460       .marketvalue_loc_forex_box.Left = (.marketvalue.Left - lngTpp)
17470       .marketvalue_box.Left = .currsym_symbol.Left
17480       .marketvalue_upd.Left = (.currsym_symbol.Left - lngTpp)
17490       .marketvalue_lbl.Left = .currsym_symbol.Left
17500       .marketvalue_lbl_line.Left = .marketvalue_lbl.Left
17510       .marketvaluenew_lbl.Left = .marketvalue_lbl.Left
17520       If blnHasCurr = True Then
17530         .marketvaluenew_lbl_line.Left = .marketvaluenew_lbl.Left
17540       End If
17550       If blnSortHere = True Then
17560         .Sort_line.Left = .marketvalue_lbl.Left
17570         .Sort_lbl.Left = ((.marketvalue_lbl.Left + .marketvalue_lbl.Width) - lngSortLbl_Width)
17580       End If
17590       blnSortHere = False
17600       If blnHasCurr = True Then
17610         If .Sort_line.Left = .currsym_symbol_usd.Left Then blnSortHere = True
17620         .currsym_symbol_usd.Left = (.currsym_symbol_usd.Left + lngAddedWidth)
17630         .marketvalue_usd.Left = (.currsym_symbol_usd.Left + lngMktValSym_Offset)
17640         .marketvalue_usd_box.Left = .currsym_symbol_usd.Left
17650         .marketvalue_usd_upd.Left = (.currsym_symbol_usd.Left - lngTpp)
17660         .marketvalue_usd_lbl.Left = .currsym_symbol_usd.Left
17670         .marketvalue_usd_lbl_line.Left = .marketvalue_usd_lbl.Left
17680         If blnSortHere = True Then
17690           .Sort_line.Left = .marketvalue_usd_lbl.Left
17700           .Sort_lbl.Left = ((.marketvalue_usd_lbl.Left + .marketvalue_usd_lbl.Width) - lngSortLbl_Width)
17710         End If
17720         blnSortHere = False
17730       End If
17740       If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
17750       .totdesc.Width = (.totdesc.Width + lngAddedWidth)
17760       .totdesc_lbl.Width = .totdesc.Width
17770       .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
17780       If blnSortHere = True Then
17790         .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
17800         .Sort_lbl.Left = (((.totdesc_lbl.Left + .totdesc_lbl.Width) - lngSortLbl_Width) + (2& * lngTpp))
17810       End If
17820       blnSortHere = False
17830     End If
17840     .FocusHolder.SetFocus
17850     If .cmdShorten.Enabled = False Then
            ' ** Since cmdWiden was clicked, cmdShorten is certainly available.
17860       .cmdShorten.Enabled = True
17870       .cmdShorten_raised_img.Visible = True
17880       .cmdShorten_raised_semifocus_dots_img.Visible = False
17890       .cmdShorten_raised_focus_img.Visible = False
17900       .cmdShorten_raised_focus_dots_img.Visible = False
17910       .cmdShorten_sunken_focus_dots_img.Visible = False
17920       .cmdShorten_raised_img_dis.Visible = False
17930     End If
          ' ** If there's nowhere to go, disable it.
17940     If (.Width >= WIDEN_MAX) Or ((.Width + lngAddedWidth) > WIDEN_MAX) Then
17950       .cmdWiden.Enabled = False
17960       .cmdWiden_raised_img_dis.Visible = True
17970       .cmdWiden_raised_img.Visible = False
17980       .cmdWiden_raised_semifocus_dots_img.Visible = False
17990       .cmdWiden_raised_focus_img.Visible = False
18000       .cmdWiden_raised_focus_dots_img.Visible = False
18010       .cmdWiden_sunken_focus_dots_img.Visible = False
18020     End If
18030   End With

EXITP:
18040   Exit Sub

ERRH:
18050   Select Case ERR.Number
        Case Else
18060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18070   End Select
18080   Resume EXITP

End Sub

Private Sub cmdWiden_KeyDown(KeyCode As Integer, Shift As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18120   intShiftDown = (Shift And acShiftMask) > 0
18130   intAltDown = (Shift And acAltMask) > 0
18140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18160     Select Case intRetVal
          Case vbKeyTab
18170       With Me
18180         intRetVal = 0
18190         If .yield_entry.Enabled = True Then
18200           .yield_entry.SetFocus
18210         Else
18220           .FocusHolder.SetFocus
18230         End If
18240       End With
18250     End Select
18260   End If

        ' ** Shift keys.
18270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18280     Select Case intRetVal
          Case vbKeyTab
18290       With Me
18300         intRetVal = 0
18310         .cmdShorten.SetFocus
18320       End With
18330     End Select
18340   End If

EXITP:
18350   KeyCode = intRetVal
18360   Exit Sub

ERRH:
18370   intRetVal = 0
18380   THAT_PROC = THIS_PROC
18390   That_Erl = Erl
18400   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18410   Resume EXITP

End Sub

Private Sub cmdWiden_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseMove"

18510   With Me
18520     If blnWiden_MouseDown = False Then
18530       Select Case blnWiden_Focus
            Case True
18540         .cmdWiden_raised_focus_dots_img.Visible = True
18550         .cmdWiden_raised_focus_img.Visible = False
18560       Case False
18570         .cmdWiden_raised_focus_img.Visible = True
18580         .cmdWiden_raised_focus_dots_img.Visible = False
18590       End Select
18600       .cmdWiden_raised_img.Visible = False
18610       .cmdWiden_raised_semifocus_dots_img.Visible = False
18620       .cmdWiden_sunken_focus_dots_img.Visible = False
18630       .cmdWiden_raised_img_dis.Visible = False
18640     End If
18650     If .cmdShorten_raised_focus_dots_img.Visible = True Or .cmdShorten_raised_focus_img.Visible = True Then
18660       Select Case blnShorten_Focus
            Case True
18670         .cmdShorten_raised_semifocus_dots_img.Visible = True
18680         .cmdShorten_raised_img.Visible = False
18690       Case False
18700         .cmdShorten_raised_img.Visible = True
18710         .cmdShorten_raised_semifocus_dots_img.Visible = False
18720       End Select
18730       .cmdShorten_raised_focus_img.Visible = False
18740       .cmdShorten_raised_focus_dots_img.Visible = False
18750       .cmdShorten_sunken_focus_dots_img.Visible = False
18760       .cmdShorten_raised_img_dis.Visible = False
18770     End If
18780   End With

EXITP:
18790   Exit Sub

ERRH:
18800   Select Case ERR.Number
        Case Else
18810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18820   End Select
18830   Resume EXITP

End Sub

Private Sub cmdWiden_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_MouseUp"

18910   With Me
18920     .cmdWiden_raised_focus_dots_img.Visible = True
18930     .cmdWiden_raised_img.Visible = False
18940     .cmdWiden_raised_semifocus_dots_img.Visible = False
18950     .cmdWiden_raised_focus_img.Visible = False
18960     .cmdWiden_sunken_focus_dots_img.Visible = False
18970     .cmdWiden_raised_img_dis.Visible = False
18980     blnWiden_MouseDown = False
18990   End With

EXITP:
19000   Exit Sub

ERRH:
19010   Select Case ERR.Number
        Case Else
19020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19030   End Select
19040   Resume EXITP

End Sub

Private Sub cmdWiden_LostFocus()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWiden_LostFocus"

19110   With Me
19120     .cmdWiden_raised_img.Visible = True
19130     .cmdWiden_raised_semifocus_dots_img.Visible = False
19140     .cmdWiden_raised_focus_img.Visible = False
19150     .cmdWiden_raised_focus_dots_img.Visible = False
19160     .cmdWiden_sunken_focus_dots_img.Visible = False
19170     .cmdWiden_raised_img_dis.Visible = False
19180     blnWiden_Focus = False
19190   End With

EXITP:
19200   Exit Sub

ERRH:
19210   Select Case ERR.Number
        Case Else
19220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19230   End Select
19240   Resume EXITP

End Sub

Public Sub ShowAssetNo(blnShow As Boolean)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAssetNo"

        Dim blnSortHere As Boolean
        Dim lngTmp01 As Long, lngTmp02 As Long, blnTmp03 As Boolean

19310   With Me
19320     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
19330       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
19340     End If
19350     blnSortHere = False: blnTmp03 = False
19360     lngTmp02 = 0&
19370     Select Case blnShow
          Case True
19380       If .assetno.Visible = False Then
19390         lngTmp01 = (.assetno.Width + (4& * lngTpp))
19400         If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
19410         If .totdesc.Width > lngTotDesc_Width Then
19420           lngTmp02 = (.totdesc.Width - lngTotDesc_Width)
19430         End If
19440         .totdesc.Width = (lngTotDescAstno_Width + lngTmp02)
19450         .totdesc.Left = (lngTotDesc_Left + lngTmp01)
19460         .totdesc_lbl.Width = .totdesc.Width
19470         .totdesc_lbl.Left = .totdesc.Left
19480         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
19490         .totdesc_lbl_line.Left = .totdesc_lbl.Left
19500         If blnSortHere = True Then
19510           .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
19520           .Sort_line.Left = .totdesc_lbl.Left
19530         End If
19540         blnSortHere = False
19550         If .Sort_line.Left = .cusip_lbl.Left Then blnSortHere = True
19560         .cusip.Left = (.assetno.Left + lngTmp01)
19570         .IsHid1.Left = (.cusip.Left - (4& * lngTpp))
19580         .IsHid2.Left = .IsHid1.Left
19590         .cusip_lbl.Left = .cusip.Left
19600         .cusip_lbl_line.Left = .cusip_lbl.Left
19610         If blnSortHere = True Then
19620           .Sort_line.Left = .cusip_lbl.Left
19630           .Sort_lbl.Left = ((.cusip_lbl.Left + .cusip_lbl.Width) - lngSortLbl_Width)
19640         End If
19650         blnSortHere = False
19660         .assetno.Visible = True
19670         .assetno_lbl.Visible = True
19680         .assetno_lbl2.Visible = True
19690         .assetno_lbl_line.Visible = True
19700       End If
19710     Case False
19720       If .assetno.Visible = True Then
19730         If .Sort_line.Left = .assetno_lbl.Left Then blnSortHere = True
19740         .assetno.Visible = False
19750         .assetno_lbl.Visible = False
19760         .assetno_lbl2.Visible = False
19770         .assetno_lbl_line.Visible = False
19780         If blnSortHere = True Then
19790           blnTmp03 = True
19800         End If
19810         blnSortHere = False
19820         If .Sort_line.Left = .cusip_lbl.Left Then blnSortHere = True
19830         .cusip.Left = .assetno.Left
19840         .IsHid1.Left = 0&
19850         .IsHid2.Left = .IsHid1.Left
19860         .cusip_lbl.Left = .cusip.Left
19870         .cusip_lbl_line.Left = .cusip_lbl.Left
19880         If blnSortHere = True Then
19890           .Sort_line.Left = .cusip_lbl.Left
19900           .Sort_lbl.Left = ((.cusip_lbl.Left + .cusip_lbl.Width) - lngSortLbl_Width)
19910         End If
19920         blnSortHere = False
19930         If .Sort_line.Left = .totdesc_lbl.Left Then blnSortHere = True
19940         .totdesc.Left = lngTotDesc_Left
19950         If .totdesc.Width > lngTotDescAstno_Width Then
19960           lngTmp02 = (.totdesc.Width - lngTotDescAstno_Width)
19970         End If
19980         .totdesc.Width = (lngTotDesc_Width + lngTmp02)
19990         .totdesc_lbl.Left = .totdesc.Left
20000         .totdesc_lbl.Width = .totdesc.Width
20010         .totdesc_lbl_line.Left = .totdesc_lbl.Left
20020         .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
20030         If blnSortHere = True Then
20040           .Sort_line.Left = .totdesc_lbl.Left
20050           .Sort_line.Width = (.totdesc_lbl.Width + lngTpp)
20060         End If
20070         blnSortHere = False
20080         If blnTmp03 = True Then
20090           SortNow "Form_Load"  ' ** Procedure: Below.
20100         End If
20110       End If
20120     End Select
20130   End With

EXITP:
20140   Exit Sub

ERRH:
20150   Select Case ERR.Number
        Case Else
20160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20170   End Select
20180   Resume EXITP

End Sub

Public Sub IncludeCurrency_Sub(blnShow As Boolean)

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

20210   With Me
20220     Select Case blnShow
          Case True
20230       lngTmp01 = (.curr_id.Width + (4& * lngTpp))
20240       lngTmp04 = lngTmp01
20250       lngTmp02 = (.marketvaluecurrent_usd_str_lbl.Width + (4& * lngTpp))
20260       lngTmp04 = lngTmp04 + lngTmp02
20270       lngTmp03 = (.marketvalue_usd_lbl.Width + (4& * lngTpp))
20280       lngTmp04 = lngTmp04 + lngTmp03
20290       If .Width = ((.marketvalue.Left + .marketvalue.Width) + (4& * lngTpp)) Then
              ' ** Form width is where it should be.
20300         .Width = (.Width + lngTmp04)
20310       Else
20320         .Width = (((.marketvalue.Left + .marketvalue.Width) + (4& * lngTpp)) + lngTmp04)
20330       End If
20340       .currsym_symbol.Left = (.currsym_symbol.Left + (lngTmp01 + lngTmp02))
20350       .currsym_symbol_loc_usd_box.Left = (.currsym_symbol.Left - lngTpp)
20360       .currsym_symbol_loc_forex_box.Left = (.currsym_symbol.Left - lngTpp)
20370       .marketvalue.Left = (.marketvalue.Left + (lngTmp01 + lngTmp02))
20380       .marketvalue_loc_usd_box.Left = (.marketvalue.Left - lngTpp)
20390       .marketvalue_loc_forex_box.Left = (.marketvalue.Left - lngTpp)
20400       .marketvalue_box.Left = (.marketvalue_box.Left + (lngTmp01 + lngTmp02))
20410       .marketvalue_upd.Left = (.marketvalue_upd.Left + (lngTmp01 + lngTmp02))
20420       .marketvalue_lbl.Left = (.marketvalue_lbl.Left + (lngTmp01 + lngTmp02))
20430       .marketvalue_lbl_line.Left = (.marketvalue_lbl_line.Left + (lngTmp01 + lngTmp02))
20440       .marketvaluecurrent_str.Left = (.marketvaluecurrent_str.Left + lngTmp01)
20450       .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
20460       .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
20470       .marketvaluecurrent_str_lbl.Left = (.marketvaluecurrent_str_lbl.Left + lngTmp01)
20480       .marketvaluecurrent_str_lbl_line.Left = (.marketvaluecurrent_str_lbl_line.Left + lngTmp01)
20490       .yield_entry.Left = (.yield_entry.Left + lngTmp01)
20500       .yield_box.Left = (.yield_box.Left + lngTmp01)
20510       .yield_upd.Left = (.yield_upd.Left + lngTmp01)
20520       .yield_entry_lbl.Left = (.yield_entry_lbl.Left + lngTmp01)
20530       .yield_entry_lbl_line.Left = (.yield_entry_lbl_line.Left + lngTmp01)
20540       .yield_pct_lbl.Left = (.yield_pct_lbl.Left + lngTmp01)
20550       .curr_id.Left = (.yield_entry.Left - lngTmp01)
20560       .curr_id_usd_box.Left = (.curr_id.Left - lngTpp)
20570       .curr_id_forex_box.Left = (.curr_id.Left - lngTpp)
20580       .curr_id_lbl.Left = (.curr_id.Left - (4& * lngTpp))
20590       .curr_id_lbl_line.Left = .curr_id.Left
20600       .curr_date.Left = .curr_id_lbl_line.Left
20610       .marketvaluecurrent_usd_str.Left = (.currsym_symbol.Left - lngTmp02)
20620       .marketvaluecurrent_usd_str_lbl.Left = .marketvaluecurrent_usd_str.Left
20630       .marketvaluecurrent_usd_str_lbl_line.Left = .marketvaluecurrent_usd_str.Left
20640       .currsym_symbol_usd.Left = (.Width - lngTmp03)
20650       .marketvalue_usd.Left = (.currsym_symbol_usd.Left + lngMktValSym_Offset)
20660       .marketvalue_usd_box.Left = .currsym_symbol_usd.Left
20670       .marketvalue_usd_upd.Left = (.currsym_symbol_usd.Left - lngTpp)
20680       .marketvalue_usd_lbl.Left = .currsym_symbol_usd.Left
20690       .marketvalue_usd_lbl_line.Left = .marketvalue_usd_lbl.Left
20700       .marketvaluecurrent_lbl.Top = 0&
20710       .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
20720       .marketvaluecurrent_lbl.Width = ((.marketvaluecurrent_usd_str.Left + .marketvaluecurrent_usd_str.Width) - .marketvaluecurrent_str.Left)
20730       .marketvaluecurrent_lbl_line.Left = .marketvaluecurrent_lbl.Left
20740       .marketvaluenew_lbl.Top = 0&
20750       .marketvaluenew_lbl.Left = .currsym_symbol.Left
20760       .marketvaluenew_lbl.Width = ((.marketvalue_usd.Left + .marketvalue_usd.Width) - .currsym_symbol.Left)
20770       .marketvaluenew_lbl_line.Left = .marketvaluenew_lbl.Left
20780       .marketvaluecurrent_lbl.Caption = "Current Market Unit Value"
20790       .marketvaluecurrent_str_lbl.Caption = "Local"
20800       .marketvaluenew_lbl.Caption = "New &Market Unit Value"
20810       .marketvalue_lbl.Caption = "Local"
20820       .curr_id.Visible = True
20830       .curr_id_usd_box.Visible = True
20840       .curr_id_forex_box.Visible = True
20850       .curr_id_lbl.Visible = True
20860       .curr_id_lbl_line.Visible = True
20870       .marketvaluecurrent_usd_str.Visible = True
20880       .marketvaluecurrent_usd_str_lbl.Visible = True
20890       .marketvaluecurrent_usd_str_lbl_line.Visible = True
20900       .currsym_symbol_usd.Visible = True
20910       .marketvalue_usd.Visible = True
20920       .marketvalue_usd_box.Visible = True
            '.marketvalue_usd_upd.Visible = True
20930       .marketvalue_usd_lbl.Visible = True
20940       .marketvalue_usd_lbl_line.Visible = True
20950       .marketvaluecurrent_lbl_line.Visible = True
20960       .marketvaluenew_lbl_line.Visible = True
20970     Case False
20980 On Error Resume Next
20990       .FocusHolder.SetFocus
21000       If ERR.Number <> 0 Then
21010         .FocusHolder2.SetFocus
21020 On Error GoTo ERRH
21030       Else
21040 On Error GoTo ERRH
21050       End If
21060       .curr_id.Visible = False
21070       .curr_id_usd_box.Visible = False
21080       .curr_id_forex_box.Visible = False
21090       .curr_id_lbl.Visible = False
21100       .curr_id_lbl_line.Visible = False
21110       .curr_id.Left = 0&
21120       .curr_id_usd_box.Left = 0&
21130       .curr_id_forex_box.Left = 0&
21140       .curr_id_lbl.Left = 0&
21150       .curr_id_lbl_line.Left = 0&
21160       .curr_date.Visible = False
21170       .curr_date.Left = 0&
21180       .marketvaluecurrent_usd_str.Visible = False
21190       .marketvaluecurrent_usd_str_lbl.Visible = False
21200       .marketvaluecurrent_usd_str_lbl_line.Visible = False
21210       .marketvaluecurrent_usd_str.Left = 0&
21220       .marketvaluecurrent_usd_str_lbl.Left = 0&
21230       .marketvaluecurrent_usd_str_lbl_line.Left = 0&
21240       .currsym_symbol_usd.Visible = False
21250       .currsym_symbol_usd.Left = 0&
21260       .marketvalue_usd.Visible = False
21270       .marketvalue_usd_box.Visible = False
            '.marketvalue_usd_upd.Visible = False
21280       .marketvalue_usd_lbl.Visible = False
21290       .marketvalue_usd_lbl_line.Visible = False
21300       .marketvalue_usd.Left = 0&
21310       .marketvalue_usd_box.Left = 0&
21320       .marketvalue_usd_upd.Left = 0&
21330       .marketvalue_usd_lbl.Left = 0&
21340       .marketvalue_usd_lbl_line.Left = 0&
21350       .marketvaluecurrent_lbl_line.Visible = False
21360       .marketvaluenew_lbl_line.Visible = False
21370       .marketvaluecurrent_lbl_line.Left = 0&
21380       .marketvaluenew_lbl_line.Left = 0&
100         lngTmp01 = .totdesc.Width
110         lngTmp02 = 0&
120         If .assetno.Visible = True Then
130           If lngTmp01 > lngTotDescAstno_Width Then
                ' ** Column has been widened.
140             lngTmp02 = (lngTmp01 - lngTotDescAstno_Width)  ' ** Amount columns pushed right.
150           End If
160         Else
170           If lngTmp01 > lngTotDesc_Width Then
                ' ** Column has been widened.
180             lngTmp02 = (lngTmp01 - lngTotDesc_Width)  ' ** Amount columns pushed right.
190           End If
200         End If
21390       .yield_entry.Left = ((lngYield_Left + lngTmp02) - lngCurrID_Width)
21400       .yield_box.Left = .yield_entry.Left
21410       .yield_upd.Left = (.yield_box.Left - lngTpp)
21420       .yield_entry_lbl.Left = .yield_entry.Left
21430       .yield_entry_lbl_line.Left = .yield_entry_lbl.Left
21440       .yield_pct_lbl.Left = (.yield_entry.Left + lngYieldPct_Offset)
21450       .marketvaluecurrent_str.Left = ((lngMktValCurStr_Left + lngTmp02) - lngCurrID_Width)
21460       .marketvaluecurrent_str_usd_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
21470       .marketvaluecurrent_str_forex_box.Left = (.marketvaluecurrent_str.Left - lngTpp)
21480       .marketvaluecurrent_str_lbl.Left = .marketvaluecurrent_str.Left
21490       .marketvaluecurrent_str_lbl_line.Left = .marketvaluecurrent_str_lbl.Left
21500       .currsym_symbol.Left = ((lngMktValSym_Left + lngTmp02) - (lngCurrID_Width + lngMktValCurUsdStr_Width))
21510       .currsym_symbol_loc_usd_box.Left = (.currsym_symbol.Left - lngTpp)
21520       .currsym_symbol_loc_forex_box.Left = (.currsym_symbol.Left - lngTpp)
21530       .marketvalue.Left = ((lngMktVal_Left + lngTmp02) - (lngCurrID_Width + lngMktValCurUsdStr_Width))
21540       .marketvalue_loc_usd_box.Left = (.marketvalue.Left - lngTpp)
21550       .marketvalue_loc_forex_box.Left = (.marketvalue.Left - lngTpp)
21560       .marketvalue_box.Left = (.marketvalue.Left - lngMktValSym_Offset)
21570       .marketvalue_upd.Left = (.marketvalue_box.Left - lngTpp)
21580       .marketvalue_lbl.Left = .marketvalue_box.Left
21590       .marketvalue_lbl_line.Left = .marketvalue_box.Left
21600       .marketvaluecurrent_lbl.Top = (7& * lngTpp)
21610       .marketvaluecurrent_lbl.Left = .marketvaluecurrent_str_lbl.Left
21620       .marketvaluecurrent_lbl.Width = .marketvaluecurrent_str_lbl.Width
21630       .marketvaluecurrent_lbl.Caption = "Current Market"
21640       .marketvaluecurrent_str_lbl.Caption = "Unit Value"
21650       .marketvaluenew_lbl.Top = (7& * lngTpp)
21660       .marketvaluenew_lbl.Left = .marketvalue_lbl.Left
21670       .marketvaluenew_lbl.Width = .marketvalue_lbl.Width
21680       .marketvaluenew_lbl.Caption = "New &Market"
21690       .marketvalue_lbl.Caption = "Unit Value"
21700       .Width = (lngForm_Width - (lngCurrID_Width + lngMktValUsd_Width + lngMktValCurUsdStr_Width + lngTmp02))
21710     End Select
21720   End With

'lngTotDesc_Width
'lngTotDescAstno_Width

        'lngYield_Left
        'lngCurrID_Left
        'lngMktVal_Left
        'lngMktValUsd_Left
        'lngMktValCurStr_Left
        'lngMktValCurUsdStr_Left
        'lngMktValSym_Left
        'lngMktValSymUsd_Left

        'lngCurrID_Width
        'lngMktValUsd_Width
        'lngMktValCurUsdStr_Width

        'lngYieldPct_Offset
        'lngMktValSym_Offset

EXITP:
21730   Exit Sub

ERRH:
21740   Select Case ERR.Number
        Case Else
21750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21760   End Select
21770   Resume EXITP

End Sub

Public Sub IsOpen_Set(blnOpen As Boolean)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "IsOpen_Set"

21810   blnIsOpen = blnOpen

EXITP:
21820   Exit Sub

ERRH:
21830   Select Case ERR.Number
        Case Else
21840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21850   End Select
21860   Resume EXITP

End Sub

Public Function ForExTot_Get() As Long
' ** The Sum() field just takes too long to show an accurate count!

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "ForExTot_Get"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long
        Dim lngRetVal As Long

21910   With Me
21920     lngRetVal = 0&
21930     Set rst = .RecordsetClone
21940     With rst
21950       If .BOF = True And .EOF = True Then
              ' ** Shouldn't normally happen.
21960       Else
21970         .MoveLast
21980         lngRecs = .RecordCount
21990         .MoveFirst
22000         For lngX = 1& To lngRecs
22010           If ![curr_id] <> 150& Then
22020             lngRetVal = lngRetVal + 1&
22030           End If
22040           If lngX < lngRecs Then .MoveNext
22050         Next
22060       End If
22070     End With
22080   End With

EXITP:
22090   Set rst = Nothing
22100   ForExTot_Get = lngRetVal
22110   Exit Function

ERRH:
22120   lngRetVal = 0&
22130   Select Case ERR.Number
        Case Else
22140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22150   End Select
22160   Resume EXITP

End Function

Public Function RecCnt() As Long

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

22210   lngRetVal = 0&

22220   With Me
22230     Set rst = .RecordsetClone
22240     With rst
22250       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
22260       Else
22270         .MoveLast
22280         lngRetVal = .RecordCount
22290       End If
22300       .Close
22310     End With
22320   End With

EXITP:
22330   Set rst = Nothing
22340   RecCnt = lngRetVal
22350   Exit Function

ERRH:
22360   lngRetVal = 0&
22370   Select Case ERR.Number
        Case Else
22380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22390   End Select
22400   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

22510   Select Case intWhere
        Case acCmdRecordsGoToFirst
22520     DoCmd.RunCommand acCmdRecordsGoToFirst
22530   Case acCmdRecordsGoToLast
22540     DoCmd.RunCommand acCmdRecordsGoToLast
22550   Case acCmdRecordsGoToPrevious
22560     DoCmd.RunCommand acCmdRecordsGoToPrevious
22570   Case acCmdRecordsGoToNext
22580     DoCmd.RunCommand acCmdRecordsGoToNext
22590   Case Else
22600     If IsMissing(varID) = False Then
22610       With Me
22620         Set rst = .RecordsetClone
22630         With rst
22640           .FindFirst "[assetno] = " & CStr(varID)
22650           If .NoMatch = False Then
22660             Me.Bookmark = .Bookmark
22670           End If
22680           .Close
22690         End With
22700       End With
22710     End If
22720   End Select

EXITP:
22730   Set rst = Nothing
22740   Exit Sub

ERRH:
22750   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Ignore.
22760   Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
22770   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
22780   Case Else
22790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22800   End Select
22810   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [description_masterasset_sort], [assettype]"

22910   With Me
22920     .Controls(strSortLbl).Visible = False
22930     .Controls(strSortLine).Visible = False
22940     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
22950     If strProc = "Form_Load" Then
22960       strCalled = "assettype"
22970       strSortNow = strSortOrig
22980       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
22990       lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (5& * lngTpp))
23000       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
23010       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
23020       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
23030       .Controls(strSortLbl).Top = lngSortLbl_Top
23040       .Controls(strSortLbl).Left = lngSortLbl_Left
23050       .Controls(strSortLbl).Caption = strArwUp
23060       .Controls(strSortLbl).ForeColor = CLR_DKBLU
23070       .Controls(strSortLine).Top = lngSortLine_Top
23080       .Controls(strSortLine).Left = lngSortLine_Left
23090       .Controls(strSortLine).Width = lngSortLine_Width
23100     Else
23110       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
23120       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
23130       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
23140       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
23150       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
23160       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
23170       .Controls(strSortLbl).Top = lngSortLbl_Top
23180       .Controls(strSortLbl).Left = lngSortLbl_Left
23190       .Controls(strSortLine).Top = lngSortLine_Top
23200       .Controls(strSortLine).Left = lngSortLine_Left
23210       .Controls(strSortLine).Width = lngSortLine_Width
23220       If strCalled = "assettype" Then
23230         strSortAsc = strSortOrig
23240         If strSortNow = strSortAsc Then
23250           strSortNow = "[assettype] DESC, [description_masterasset_sort] DESC, [cusip] DESC"
23260           .Controls(strSortLbl).Caption = strArwDn
23270           .Controls(strSortLbl).ForeColor = CLR_DKRED
23280         Else
23290           strSortNow = strSortAsc
23300           .Controls(strSortLbl).Caption = strArwUp
23310           .Controls(strSortLbl).ForeColor = CLR_DKBLU
23320         End If
23330         lngSortLbl_Left = lngSortLbl_Left + (5& * lngTpp)
23340         .Controls(strSortLbl).Left = lngSortLbl_Left
23350       Else
23360         .Controls(strSortLbl).Caption = strArwUp
23370         .Controls(strSortLbl).ForeColor = CLR_DKBLU
23380         Select Case strCalled
              Case "assetno"
23390           strSortAsc = "[assetno]"
23400           lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
23410           lngSortLbl_Left = lngSortLbl_Left + (2& * lngTpp)
23420           lngSortLine_Top = (lngSortLine_Top - lngTpp)
23430           .Controls(strSortLbl).Top = lngSortLbl_Top
23440           .Controls(strSortLbl).Left = lngSortLbl_Left
23450           .Controls(strSortLine).Top = lngSortLine_Top
23460         Case "cusip"
23470           strSortAsc = "[cusip]" & strStdAsc
23480         Case "totdesc"
23490           strSortAsc = "[description_masterasset_sort], [cusip], [assettype]"
23500           lngSortLbl_Left = (lngSortLbl_Left + (2& * lngTpp))
23510           .Controls(strSortLbl).Left = lngSortLbl_Left
23520         Case "curr_id"
23530           strSortAsc = "[curr_code]" & strStdAsc
23540           lngSortLbl_Left = lngSortLbl_Left + (3& * lngTpp)
23550           lngSortLine_Left = .Controls(strCalled).Left
23560           lngSortLine_Width = (.Controls(strCalled).Width + lngTpp)
23570           .Controls(strSortLbl).Left = lngSortLbl_Left
23580           .Controls(strSortLine).Left = lngSortLine_Left
23590           .Controls(strSortLine).Width = lngSortLine_Width
23600         Case "yield", "yield_entry"
23610           strSortAsc = "[yield]" & strStdAsc
23620           lngSortLbl_Left = lngSortLbl_Left + (1& * lngTpp)
23630           .Controls(strSortLbl).Left = lngSortLbl_Left
23640         Case "marketvaluecurrent", "marketvaluecurrent_str"
23650           strSortAsc = "[marketvaluecurrent]" & strStdAsc
23660         Case "marketvaluecurrent_usd", "marketvaluecurrent_usd_str"
23670           strSortAsc = "[marketvaluecurrent_usd]" & strStdAsc
23680         Case "marketvalue"
23690           strSortAsc = "[marketvalue]" & strStdAsc
23700         Case "marketvalue_usd"
23710           strSortAsc = "[marketvalue_usd]" & strStdAsc
23720         End Select
23730         If strSortNow = strSortAsc Then
23740           intCnt = (CharCnt(strSortAsc, ",") + 1)  ' ** Module Function: modStringFuncs.
23750           Select Case intCnt
                Case 1
23760             strSortNow = strSortAsc & " DESC"
23770           Case 2
23780             intPos01 = InStr(strSortAsc, ",")
23790             strSortNow = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
23800           Case 3
23810             intPos01 = InStr(strSortAsc, ",")
23820             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
23830             strTmp02 = Mid(strSortAsc, intPos01)
23840             intPos01 = InStr(2, strTmp02, ",")
23850             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
23860             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
23870             strSortNow = strTmp01
23880           End Select
23890           .Controls(strSortLbl).Caption = strArwDn
23900           .Controls(strSortLbl).ForeColor = CLR_DKRED
23910         Else
23920           strSortNow = strSortAsc
23930         End If
23940       End If
23950     End If
23960     .Controls(strSortLbl).Visible = True
23970     .Controls(strSortLine).Visible = True
23980     .OrderBy = strSortNow
23990     .OrderByOn = True
24000   End With

        ' ** Possible sorts:
        ' **   [assettype], [description_masterasset_sort], [cusip]
        ' **   [assetno]
        ' **   [cusip], [description_masterasset_sort], [assettype]
        ' **   [description_masterasset_sort], [cusip], [assettype]
        ' **   [curr_code], [description_masterasset_sort], [assettype]
        ' **   [yield], [description_masterasset_sort], [assettype]
        ' **   [marketvaluecurrent], [description_masterasset_sort], [assettype]
        ' **   [marketvaluecurrent_usd], [description_masterasset_sort], [assettype]
        ' **   [marketvalue], [description_masterasset_sort], [assettype]
        ' **   [marketvalue_usd], [description_masterasset_sort], [assettype]

EXITP:
24010   Exit Sub

ERRH:
24020   Select Case ERR.Number
        Case Else
24030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24040   End Select
24050   Resume EXITP

End Sub

Public Function SortNow_Get() As String

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

24110   strRetVal = strSortNow

EXITP:
24120   SortNow_Get = strRetVal
24130   Exit Function

ERRH:
24140   strRetVal = strSortOrig
24150   Select Case ERR.Number
        Case Else
24160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24170   End Select
24180   Resume EXITP

End Function
