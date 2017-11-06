VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_Rate_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_Rate_Sub"

'VGC 04/17/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Active:           A {curr_active}
' **   Close:            C {cmdClose on frmCurrency_Rate}
' **   Code:             D {opgCurrencySource_optCurrCode on frmCurrency_Rate}
' **   Edit:             E {cmdEdit on frmCurrency_Rate}
' **   Find:             F {cmbCurrencies on frmCurrency_Rate}
' **   ISO:              I {opgCurrencySource_optCurrISO on frmCurrency_Rate}
' **   Excel:            L {cmdExcel on frmCurrency_Rate}
' **   Name:             M {opgCurrencySource_optCurrName on frmCurrency_Rate}
' **   Countries:        N {cmdCountries on frmCurrency_Rate}
' **   Word:             O {cmdWord on frmCurrency_Rate}
' **   Print:            P {cmdPrint on frmCurrency_Rate}
' **   Currency Name:    R {curr_name}
' **   Reset:            S {cmdResetFilter on frmCurrency_Rate}
' **   Date:             T {curr_date}
' **   Update:           U {cmdUpdate on frmCurrency_Rate}
' **   Convert 1:        V {curr_convert1a}
' **   Preview:          W {cmdPreview on frmCurrency_Rate}
' **   Exit:             X {cmdClose on frmCurrency_Rate}
' **   Country:          Y {opgCurrencySource_optCountry on frmCurrency_Rate}
' **   Convert 2:        Z {curr_convert2a}
' **   Rate 1:           1 {curr_rate1}
' **   Rate 2:           2 {curr_rate2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Exclude Active:   E {opgActive_optExclude on frmCurrency_Rate}
' **   Page Count:       G {chkPageOf on frmCurrency_Rate}
' **   Include Active:   I {opgActive_optInclude on frmCurrency_Rate}
' **   Include Notes:    N {chkIncludeNotes on frmCurrency_Rate}
' **   Only Active:      O {opgActive_optOnly on frmCurrency_Rate}
' **   Save:             S {cmdSave}
' **   Word:             W {opgCurrencySource_optCurrWord on frmCurrency_Rate}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Exclude Funds:    E {opgFund_optExclude on frmCurrency_Rate}
' **   Include Funds:    I {opgFund_optInclude on frmCurrency_Rate}
' **   Only Metals:      L {opgMetal_optOnly on frmCurrency_Rate}
' **   Include Metals:   N {opgMetal_optInclude on frmCurrency_Rate}
' **   Only Funds:       O {opgFund_optOnly on frmCurrency_Rate}
' **   Exclude Metals:   X {opgMetal_optExclude on frmCurrency_Rate}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   Exclude Units:    D {opgUnit_optExclude on frmCurrency_Rate}
' **   Exclude BMUs:     E {opgBMU_optExclude on frmCurrency_Rate}
' **   Include BMUs:     I {opgBMU_optInclude on frmCurrency_Rate}
' **   Only Misc:        L {opgAlt_optOnly on frmCurrency_Rate}
' **   Include Misc:     N {opgAlt_optInclude on frmCurrency_Rate}
' **   Only BMUs:        O {opgBMU_optOnly on frmCurrency_Rate}
' **   Include Units:    U {opgUnit_optInclude on frmCurrency_Rate}
' **   Exclude Misc:     X {opgAlt_optExclude on frmCurrency_Rate}
' **   Only Units:       Y {opgUnit_optOnly on frmCurrency_Rate}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[curr_name_sort]"
Private Const strDblClick As String = "_lbl_cmd_DblClick"
Private Const strSortLine2 As String = "Sort_line2"
Private Const strSortLine3 As String = "Sort_line3"
Private Const strSortLbl2 As String = "Sort_lbl2"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strControlName As String
Private strSortNow As String, lngSortLbl2_Top As Long, lngSortLbl2_Left As Long, lngSortLbl2_Width As Long
Private lngSortLine2_Top As Long, lngSortLine2_Left As Long, lngSortLine2_Width As Long
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
150       If strCallingForm = "frmCurrency_Rate" Then

160         lngSortLbl2_Width = .Sort_lbl2.Width

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         strControlName = vbNullString

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

          ' ** These are always disabled.
420       .curr_code.Enabled = False
430       .curr_code.Locked = True
440       .curr_code.ForeColor = CLR_DISABLED_FG
450       .curr_code.BackColor = CLR_DISABLED_BG
460       .curr_iso.Enabled = False
470       .curr_iso.Locked = True
480       .curr_iso.ForeColor = CLR_DISABLED_FG
490       .curr_iso.BackColor = CLR_DISABLED_BG
500       .curr_datemodified.Enabled = False
510       .curr_datemodified.Locked = True
520       .curr_datemodified.ForeColor = CLR_DISABLED_FG
530       .curr_datemodified.BackColor = CLR_DISABLED_BG

          ' ** These are always locked.
540       .curr_name.Enabled = True
550       .curr_name.Locked = True
560       .curr_name.ForeColor = CLR_DISABLED_FG
570       .curr_name.BackColor = CLR_DISABLED_BG
580       .curr_convert1b.Enabled = True
590       .curr_convert1b.Locked = True
600       .curr_convert1b.ForeColor = CLR_DISABLED_FG
610       .curr_convert1b_box.BackColor = CLR_DISABLED_BG
620       .curr_convert2b.Enabled = True
630       .curr_convert2b.Locked = True
640       .curr_convert2b.ForeColor = CLR_DISABLED_FG
650       .curr_convert2b_box.BackColor = CLR_DISABLED_BG

          ' ** These will be fully available.
660       .curr_notes.Enabled = True
670       .curr_notes.Locked = True
680       .curr_notes.ForeColor = CLR_DISABLED_FG
690       .curr_notes.BackColor = CLR_DISABLED_BG
700       .curr_active.Enabled = True
710       .curr_active.Locked = True
720       .curr_active_box1.ForeColor = MY_CLR_BGE 'CLR_IEI_BLU_DIS
730       .curr_active_box2.ForeColor = WIN_CLR_DISR 'CLR_LTBLU2
740       .curr_rate1.Enabled = True
750       .curr_rate1.Locked = True
760       .curr_rate1.ForeColor = CLR_DISABLED_FG
770       .curr_rate1.BackColor = CLR_DISABLED_BG
780       .curr_rate2.Enabled = True
790       .curr_rate2.Locked = True
800       .curr_rate2.ForeColor = CLR_DISABLED_FG
810       .curr_rate2.BackColor = CLR_DISABLED_BG
820       .curr_date.Enabled = True
830       .curr_date.Locked = True
840       .curr_date.ForeColor = CLR_DISABLED_FG
850       .curr_date.BackColor = CLR_DISABLED_BG
860       .curr_convert1a.Enabled = True
870       .curr_convert1a.Locked = True
880       .curr_convert1a.ForeColor = CLR_DISABLED_FG
890       .curr_convert1a.BackColor = CLR_DISABLED_BG
900       .curr_convert1a_box.BackColor = CLR_DISABLED_BG
910       .curr_convert2a.Enabled = True
920       .curr_convert2a.Locked = True
930       .curr_convert2a.ForeColor = CLR_DISABLED_FG
940       .curr_convert2a.BackColor = CLR_DISABLED_BG
950       .curr_convert2a_box.BackColor = CLR_DISABLED_BG
960       .currsym_symbol_box1a.Visible = False
970       .currsym_symbol_box1b.Visible = False
980       .currsym_symbol_box2a.Visible = False
990       .currsym_symbol_box2b.Visible = False

1000      .currsym_symbol1b_lbl_img_dis.Visible = True
1010      .currsym_symbol1b_lbl_img.Visible = False
1020      .currsym_symbol2a_lbl_img_dis.Visible = True
1030      .currsym_symbol2a_lbl_img.Visible = False

1040    End With

1050    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
1060    Exit Sub

ERRH:
1070    Select Case ERR.Number
        Case Else
1080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1090    End Select
1100    Resume EXITP

End Sub

Private Sub Form_Current()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim strDocName As String

1210    With Me
1220      .Parent.country_cnt = .country_cnt
1230      strDocName = "frmCurrency_Country"
1240      If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
1250        Forms(strDocName).cmbCurrencies = .curr_id
1260      End If
1270    End With

EXITP:
1280    Exit Sub

ERRH:
1290    Select Case ERR.Number
        Case Else
1300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1310    End Select
1320    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1420    intShiftDown = (Shift And acShiftMask) > 0
1430    intAltDown = (Shift And acAltMask) > 0
1440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Convert 1:        V {curr_convert1a}
        ' **   Convert 2:        Z {curr_convert2a}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Exclude Active:   E {opgActive_optExclude on frmCurrency_Rate}
        ' **   Page Count:       G {chkPageOf on frmCurrency_Rate}
        ' **   Include Active:   I {opgActive_optInclude on frmCurrency_Rate}
        ' **   Include Notes:    N {chkIncludeNotes on frmCurrency_Rate}
        ' **   Only Active:      O {opgActive_optOnly on frmCurrency_Rate}
        ' **   Save:             S {cmdSave}
        ' **   Word:             W {opgCurrencySource_optCurrWord on frmCurrency_Rate}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Exclude Funds:    E {opgFund_optExclude on frmCurrency_Rate}
        ' **   Include Funds:    I {opgFund_optInclude on frmCurrency_Rate}
        ' **   Only Metals:      L {opgMetal_optOnly on frmCurrency_Rate}
        ' **   Include Metals:   N {opgMetal_optInclude on frmCurrency_Rate}
        ' **   Only Funds:       O {opgFund_optOnly on frmCurrency_Rate}
        ' **   Exclude Metals:   X {opgMetal_optExclude on frmCurrency_Rate}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   Exclude Units:    D {opgUnit_optExclude on frmCurrency_Rate}
        ' **   Exclude BMUs:     E {opgBMU_optExclude on frmCurrency_Rate}
        ' **   Include BMUs:     I {opgBMU_optInclude on frmCurrency_Rate}
        ' **   Only Misc:        L {opgAlt_optOnly on frmCurrency_Rate}
        ' **   Include Misc:     N {opgAlt_optInclude on frmCurrency_Rate}
        ' **   Only BMUs:        O {opgBMU_optOnly on frmCurrency_Rate}
        ' **   Include Units:    U {opgUnit_optInclude on frmCurrency_Rate}
        ' **   Exclude Misc:     X {opgAlt_optExclude on frmCurrency_Rate}
        ' **   Only Units:       Y {opgUnit_optOnly on frmCurrency_Rate}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        With Me
1480          intRetVal = 0
1490          .Parent.cmdClose_Click  ' ** Form Procedure: frmCurrency_Rate.
1500        End With
1510      End Select
1520    End If

        ' ** Alt keys.
1530    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1540      Select Case intRetVal
          Case vbKeyV
1550        With Me
1560          intRetVal = 0
1570          .curr_convert1a.SetFocus
1580        End With
1590      Case vbKeyX
1600        With Me
1610          intRetVal = 0
1620          .Parent.cmdClose_Click  ' ** Form Procedure: frmCurrency_Rate.
1630        End With
1640      Case vbKeyZ
1650        With Me
1660          intRetVal = 0
1670          .curr_convert2a.SetFocus
1680        End With
1690      End Select
1700    End If

        ' ** Ctrl keys.
1710    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1720      Select Case intRetVal
          Case vbKeyE
1730        With Me
1740          intRetVal = 0
1750          DoCmd.SelectObject acForm, .Parent.Name, False
1760          .Parent.opgActive.SetFocus
1770          .Parent.opgActive = .Parent.opgActive_optExclude.OptionValue
1780          .Parent.opgActive_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
1790        End With
1800      Case vbKeyG
1810        With Me
1820          intRetVal = 0
1830          DoCmd.SelectObject acForm, .Parent.Name, False
1840          .Parent.chkPageOf.SetFocus
1850          .Parent.chkPageOf = (Not .Parent.chkPageOf)
1860          .Parent.chkPageOf_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
1870        End With
1880      Case vbKeyI
1890        With Me
1900          intRetVal = 0
1910          DoCmd.SelectObject acForm, .Parent.Name, False
1920          .Parent.opgActive.SetFocus
1930          .Parent.opgActive = .Parent.opgActive_optInclude.OptionValue
1940          .Parent.opgActive_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
1950        End With
1960      Case vbKeyN
1970        With Me
1980          intRetVal = 0
1990          DoCmd.SelectObject acForm, .Parent.Name, False
2000          .Parent.chkIncludeNotes.SetFocus
2010          .Parent.chkIncludeNotes = (Not .Parent.chkIncludeNotes)
2020          .Parent.chkIncludeNotes_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
2030        End With
2040      Case vbKeyO
2050        With Me
2060          intRetVal = 0
2070          DoCmd.SelectObject acForm, .Parent.Name, False
2080          .Parent.opgActive.SetFocus
2090          .Parent.opgActive = .Parent.opgActive_optOnly.OptionValue
2100          .Parent.opgActive_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
2110        End With
2120      Case vbKeyS
2130        With Me
2140          intRetVal = 0
2150          cmdSave_Click  ' ** Form Procedure: frmCurrency_Rate.
2160        End With
2170      Case vbKeyW
2180        With Me
2190          intRetVal = 0
2200          DoCmd.SelectObject acForm, .Parent.Name, False
2210          .Parent.opgCurrencySource.SetFocus
2220          .Parent.opgCurrencySource = .Parent.opgCurrencySource_optCurrWord.OptionValue
2230          .Parent.opgCurrencySource_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
2240        End With
2250      Case vbKeyTab, vbKeyReturn
2260        With Me
2270          intRetVal = 0
2280          DoCmd.SelectObject acForm, .Parent.Name, False
2290          If .Parent.opgActive.Enabled = True Then
2300            .Parent.opgActive.SetFocus
2310          ElseIf .Parent.cmdPreview.Enabled = True Then
2320            .Parent.cmdPreview.SetFocus
2330          Else
2340            .Parent.opgActive.SetFocus
2350          End If
2360        End With
2370      Case vbKeyUp, vbKeyPageUp
2380        intRetVal = 0
2390        MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmCurrency_Rate.
2400      Case vbKeyDown, vbKeyPageDown
2410        intRetVal = 0
2420        MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmCurrency_Rate.
2430      End Select
2440    End If

        ' ** Ctrl-Shift keys.
2450    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2460      Select Case intRetVal
          Case vbKeyE
2470        With Me
2480          intRetVal = 0
2490          DoCmd.SelectObject acForm, .Parent.Name, False
2500          .Parent.opgFund.SetFocus
2510          .Parent.opgFund = .Parent.opgFund_optExclude.OptionValue
2520          .Parent.opgFund_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
2530        End With
2540      Case vbKeyF
2550        With Me
2560          intRetVal = 0
2570          DoCmd.SelectObject acForm, .Parent.Name, False
2580          .Parent.FocusHolder.SetFocus
2590        End With
2600      Case vbKeyI
2610        With Me
2620          intRetVal = 0
2630          DoCmd.SelectObject acForm, .Parent.Name, False
2640          .Parent.opgFund.SetFocus
2650          .Parent.opgFund = .Parent.opgFund_optInclude.OptionValue
2660          .Parent.opgFund_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
2670        End With
2680      Case vbKeyL
2690        With Me
2700          intRetVal = 0
2710          DoCmd.SelectObject acForm, .Parent.Name, False
2720          .Parent.opgMetal.SetFocus
2730          .Parent.opgMetal = .Parent.opgMetal_optOnly.OptionValue
2740          .Parent.opgMetal_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
2750        End With
2760      Case vbKeyN
2770        With Me
2780          intRetVal = 0
2790          DoCmd.SelectObject acForm, .Parent.Name, False
2800          .Parent.opgMetal.SetFocus
2810          .Parent.opgMetal = .Parent.opgMetal_optInclude.OptionValue
2820          .Parent.opgMetal_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
2830        End With
2840      Case vbKeyO
2850        With Me
2860          intRetVal = 0
2870          DoCmd.SelectObject acForm, .Parent.Name, False
2880          .Parent.opgFund.SetFocus
2890          .Parent.opgFund = .Parent.opgFund_optOnly.OptionValue
2900          .Parent.opgFund_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
2910        End With
2920      Case vbKeyX
2930        With Me
2940          intRetVal = 0
2950          DoCmd.SelectObject acForm, .Parent.Name, False
2960          .Parent.opgMetal.SetFocus
2970          .Parent.opgMetal = .Parent.opgMetal_optExclude.OptionValue
2980          .Parent.opgMetal_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
2990        End With
3000      Case vbKeyTab, vbKeyReturn
3010        With Me
3020          intRetVal = 0
3030          DoCmd.SelectObject acForm, .Parent.Name, False
3040          .Parent.cmdClose.SetFocus
3050        End With
3060      End Select
3070    End If

        ' ** Ctrl-Alt keys.
3080    If intCtrlDown And intAltDown And (Not intShiftDown) Then
3090      Select Case intRetVal
          Case vbKeyD
3100        With Me
3110          intRetVal = 0
3120          DoCmd.SelectObject acForm, .Parent.Name, False
3130          .Parent.opgUnit.SetFocus
3140          .Parent.opgUnit = .Parent.opgUnit_optExclude.OptionValue
3150          .Parent.opgUnit_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
3160        End With
3170      Case vbKeyE
3180        With Me
3190          intRetVal = 0
3200          DoCmd.SelectObject acForm, .Parent.Name, False
3210          .Parent.opgBMU.SetFocus
3220          .Parent.opgBMU = .Parent.opgBMU_optExclude.OptionValue
3230          .Parent.opgBMU_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
3240        End With
3250      Case vbKeyI
3260        With Me
3270          intRetVal = 0
3280          DoCmd.SelectObject acForm, .Parent.Name, False
3290          .Parent.opgBMU.SetFocus
3300          .Parent.opgBMU = .Parent.opgBMU_optInclude.OptionValue
3310          .Parent.opgBMU_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
3320        End With
3330      Case vbKeyL
3340        With Me
3350          intRetVal = 0
3360          DoCmd.SelectObject acForm, .Parent.Name, False
3370          .Parent.opgAlt.SetFocus
3380          .Parent.opgAlt = .Parent.opgAlt_optOnly.OptionValue
3390          .Parent.opgAlt_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
3400        End With
3410      Case vbKeyN
3420        With Me
3430          intRetVal = 0
3440          DoCmd.SelectObject acForm, .Parent.Name, False
3450          .Parent.opgAlt.SetFocus
3460          .Parent.opgAlt = .Parent.opgAlt_optInclude.OptionValue
3470          .Parent.opgAlt_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
3480        End With
3490      Case vbKeyO
3500        With Me
3510          intRetVal = 0
3520          DoCmd.SelectObject acForm, .Parent.Name, False
3530          .Parent.opgBMU.SetFocus
3540          .Parent.opgBMU = .Parent.opgBMU_optOnly.OptionValue
3550          .Parent.opgBMU_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
3560        End With
3570      Case vbKeyU
3580        With Me
3590          intRetVal = 0
3600          DoCmd.SelectObject acForm, .Parent.Name, False
3610          .Parent.opgUnit.SetFocus
3620          .Parent.opgUnit = .Parent.opgUnit_optInclude.OptionValue
3630          .Parent.opgUnit_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
3640        End With
3650      Case vbKeyX
3660        With Me
3670          intRetVal = 0
3680          DoCmd.SelectObject acForm, .Parent.Name, False
3690          .Parent.opgAlt.SetFocus
3700          .Parent.opgAlt = .Parent.opgAlt_optExclude.OptionValue
3710          .Parent.opgAlt_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
3720        End With
3730      Case vbKeyY
3740        With Me
3750          intRetVal = 0
3760          DoCmd.SelectObject acForm, .Parent.Name, False
3770          .Parent.opgUnit.SetFocus
3780          .Parent.opgUnit = .Parent.opgUnit_optOnly.OptionValue
3790          .Parent.opgUnit_AfterUpdate  ' ** Form Procedure: frmCurrency_Rate.
3800        End With
3810      End Select
3820    End If

EXITP:
3830    KeyCode = intRetVal
3840    Exit Sub

ERRH:
3850    intRetVal = 0
3860    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3870    Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

4010    With Me
4020      .curr_datemodified = Now()
4030      .currratestage_datemodified = Now()
4040      .curr_username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
4050      .currratestage_username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
4060      .currratestage_changed = True
4070      If IsNull(.curr_rate1) = True Then
4080        .curr_rate1 = 0#
4090      End If
4100      If IsNull(.curr_rate2) = True Then
4110        .curr_rate2 = 0#
4120      End If
4130      .Parent.currratestage_changed = True
4140      If .Parent.cmdUpdate.Enabled = False Then
4150        .Parent.cmdUpdate.Enabled = True
4160        .Parent.cmdEdit.Enabled = False
4170      End If
4180    End With

EXITP:
4190    Exit Sub

ERRH:
4200    Select Case ERR.Number
        Case Else
4210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4220    End Select
4230    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

4310    If DataErr <> 0 Then
4320      Select Case DataErr
          Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
            ' ** Probably a bad date. Ignore.
4330        Response = acDataErrContinue
4340      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
4350        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
4360        Response = acDataErrContinue
4370        DoCmd.CancelEvent
4380      Case Else
4390        Response = acDataErrDisplay
4400      End Select
4410    End If

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Public Sub cmdSave_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

4510    With Me

4520      DoCmd.RunCommand acCmdSaveRecord
4530      DoEvents

4540    End With

EXITP:
4550    Exit Sub

ERRH:
4560    Select Case ERR.Number
        Case Else
4570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4580    End Select
4590    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

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
4710            .curr_name.SetFocus
4720          Else
4730            DoCmd.SelectObject acForm, .Parent.Name, False
4740            If .Parent.opgActive.Enabled = True Then
4750              .Parent.opgActive.SetFocus
4760            ElseIf .Parent.cmdPreview.Enabled = True Then
4770              .Parent.cmdPreview.SetFocus
4780            Else
4790              .Parent.opgActive.SetFocus
4800            End If
4810          End If
4820        End With
4830      End Select
4840    End If

        ' ** Shift keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4870        With Me
4880          intRetVal = 0
4890          If .CurrentRecord > 1 Then
4900            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4910            .curr_active.SetFocus
4920          Else
4930            DoCmd.SelectObject acForm, .Parent.Name, False
4940            .Parent.cmdClose.SetFocus
4950          End If
4960        End With
4970      End Select
4980    End If

EXITP:
4990    KeyCode = intRetVal
5000    Exit Sub

ERRH:
5010    intRetVal = 0
5020    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5030    Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub curr_code_lbl_cmd_DblClick(Cancel As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_lbl_cmd_DblClick"

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

Private Sub curr_iso_lbl_cmd_DblClick(Cancel As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_iso_lbl_cmd_DblClick"

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

Private Sub curr_name_lbl_cmd_DblClick(Cancel As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_lbl_cmd_DblClick"

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

Private Sub curr_name_GotFocus()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_GotFocus"

5410    With Me
5420      .curr_name.SelLength = 0
5430      .curr_name.SelStart = 99
5440    End With

EXITP:
5450    Exit Sub

ERRH:
5460    Select Case ERR.Number
        Case Else
5470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5480    End Select
5490    Resume EXITP

End Sub

Private Sub curr_name_KeyDown(KeyCode As Integer, Shift As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5520    intShiftDown = (Shift And acShiftMask) > 0
5530    intAltDown = (Shift And acAltMask) > 0
5540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5570        With Me
5580          intRetVal = 0
5590          .curr_rate1.SetFocus
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5650        With Me
5660          intRetVal = 0
5670          If .CurrentRecord > 1 Then
5680            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5690            .curr_active.SetFocus
5700          Else
5710            DoCmd.SelectObject acForm, .Parent.Name, False
5720            .Parent.cmdClose.SetFocus
5730          End If
5740        End With
5750      End Select
5760    End If

EXITP:
5770    KeyCode = intRetVal
5780    Exit Sub

ERRH:
5790    intRetVal = 0
5800    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5810    Case Else
5820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5830    End Select
5840    Resume EXITP

End Sub

Private Sub curr_rate1_lbl_cmd_DblClick(Cancel As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_lbl_cmd_DblClick"

5910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5920    Exit Sub

ERRH:
5930    Select Case ERR.Number
        Case Else
5940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5950    End Select
5960    Resume EXITP

End Sub

Private Sub curr_rate1_BeforeUpdate(Cancel As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_BeforeUpdate"

6010    With Me
6020      If IsNull(.curr_rate1) = False Then
6030        If .curr_rate1 < 0 Then
6040          MsgBox "Rate must positive.", vbInformation + vbOKOnly, "Invalid Entry"
6050          Cancel = -1
6060        End If
6070      End If
6080    End With

EXITP:
6090    Exit Sub

ERRH:
6100    Select Case ERR.Number
        Case Else
6110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6120    End Select
6130    Resume EXITP

End Sub

Private Sub curr_rate1_AfterUpdate()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_AfterUpdate"

6210    With Me
6220      Select Case IsNull(.curr_rate1)
          Case True
6230        .curr_rate1 = 0#
6240        .curr_convert1a = Null
6250        .curr_convert1b = Null
6260      Case False
6270        If IsNull(.curr_convert1a) = False Then
6280          If .curr_convert1a <> 0# Then
6290            .curr_convert1b = (.curr_convert1a * .curr_rate1)
6300          End If
6310        End If
6320      End Select
6330      cmdSave_Click  ' ** Procedure: Above.
6340    End With

EXITP:
6350    Exit Sub

ERRH:
6360    Select Case ERR.Number
        Case Else
6370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6380    End Select
6390    Resume EXITP

End Sub

Private Sub curr_rate1_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6420    intShiftDown = (Shift And acShiftMask) > 0
6430    intAltDown = (Shift And acAltMask) > 0
6440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6470        With Me
6480          intRetVal = 0
6490          .curr_rate2.SetFocus
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6550        With Me
6560          intRetVal = 0
6570          .curr_name.SetFocus
6580        End With
6590      End Select
6600    End If

EXITP:
6610    KeyCode = intRetVal
6620    Exit Sub

ERRH:
6630    intRetVal = 0
6640    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6650    Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
6660    Case Else
6670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6680    End Select
6690    Resume EXITP

End Sub

Private Sub curr_rate2_lbl_cmd_DblClick(Cancel As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_lbl_cmd_DblClick"

6710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6720    Exit Sub

ERRH:
6730    Select Case ERR.Number
        Case Else
6740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6750    End Select
6760    Resume EXITP

End Sub

Private Sub curr_rate2_BeforeUpdate(Cancel As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_BeforeUpdate"

6810    With Me
6820      If IsNull(.curr_rate2) = False Then
6830        If .curr_rate2 < 0 Then
6840          MsgBox "Rate must positive.", vbInformation + vbOKOnly, "Invalid Entry"
6850          Cancel = -1
6860        End If
6870      End If
6880    End With

EXITP:
6890    Exit Sub

ERRH:
6900    Select Case ERR.Number
        Case Else
6910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6920    End Select
6930    Resume EXITP

End Sub

Private Sub curr_rate2_AfterUpdate()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_AfterUpdate"

7010    With Me
7020      Select Case IsNull(.curr_rate2)
          Case True
7030        .curr_rate2 = 0#
7040        .curr_convert2a = Null
7050        .curr_convert2b = Null
7060      Case False
7070        If IsNull(.curr_convert2a) = False Then
7080          If .curr_convert2a <> 0# Then
7090            .curr_convert2b = (.curr_convert2a * .curr_rate2)
7100          End If
7110        End If
7120      End Select
7130      cmdSave_Click  ' ** Procedure: Above.
7140    End With

EXITP:
7150    Exit Sub

ERRH:
7160    Select Case ERR.Number
        Case Else
7170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7180    End Select
7190    Resume EXITP

End Sub

Private Sub curr_rate2_KeyDown(KeyCode As Integer, Shift As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7220    intShiftDown = (Shift And acShiftMask) > 0
7230    intAltDown = (Shift And acAltMask) > 0
7240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7270        With Me
7280          intRetVal = 0
7290          .curr_date.SetFocus
7300        End With
7310      End Select
7320    End If

        ' ** Shift keys.
7330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7350        With Me
7360          intRetVal = 0
7370          .curr_rate1.SetFocus
7380        End With
7390      End Select
7400    End If

EXITP:
7410    KeyCode = intRetVal
7420    Exit Sub

ERRH:
7430    intRetVal = 0
7440    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7450    Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
7460    Case Else
7470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7480    End Select
7490    Resume EXITP

End Sub

Private Sub curr_date_lbl_cmd_DblClick(Cancel As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_lbl_cmd_DblClick"

7510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7520    Exit Sub

ERRH:
7530    Select Case ERR.Number
        Case Else
7540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7550    End Select
7560    Resume EXITP

End Sub

Private Sub curr_date_AfterUpdate()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_AfterUpdate"

7610    With Me
7620      Select Case IsDate(.curr_date)
          Case True
7630        cmdSave_Click  ' ** Procedure: Above.
7640      Case False
7650        .curr_date = Null
7660        cmdSave_Click  ' ** Procedure: Above.
7670      End Select
7680    End With

EXITP:
7690    Exit Sub

ERRH:
7700    Select Case ERR.Number
        Case Else
7710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7720    End Select
7730    Resume EXITP

End Sub

Private Sub curr_date_KeyDown(KeyCode As Integer, Shift As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7820    intShiftDown = (Shift And acShiftMask) > 0
7830    intAltDown = (Shift And acAltMask) > 0
7840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7870        With Me
7880          intRetVal = 0
7890          Select Case IsDate(.curr_date)
              Case True
7900            .curr_convert1a.SetFocus
7910          Case False
7920            If .curr_date.text = "__/__/____" Then
7930              .curr_convert1a.SetFocus
7940            Else
7950              Select Case IsDate(.curr_date.text)
                  Case True
7960                .curr_convert1a.SetFocus
7970              Case False
7980                MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
7990              End Select
8000            End If
8010          End Select
8020        End With
8030      End Select
8040    End If

        ' ** Shift keys.
8050    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8070        With Me
8080          intRetVal = 0
8090          Select Case IsDate(.curr_date)
              Case True
8100            .curr_rate2.SetFocus
8110          Case False
8120            If .curr_date.text = "__/__/____" Then
8130              .curr_rate2.SetFocus
8140            Else
8150              Select Case IsDate(.curr_date.text)
                  Case True
8160                .curr_rate2.SetFocus
8170              Case False
8180                MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
8190              End Select
8200            End If
8210          End Select
8220        End With
8230      End Select
8240    End If

EXITP:
8250    KeyCode = intRetVal
8260    Exit Sub

ERRH:
8270    intRetVal = 0
8280    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8290    Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
8300    Case Else
8310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8320    End Select
8330    Resume EXITP

End Sub

Private Sub curr_convert1a_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_convert1a_AfterUpdate"

8410    With Me
8420      Select Case IsNull(.curr_convert1a)
          Case True
            ' ** Leave it.
8430      Case False
8440        Select Case IsNull(.curr_rate1)
            Case True
8450          .curr_convert1b = Null
8460        Case False
8470          If .curr_rate1 > 0 Then
8480            .curr_convert1b = (.curr_convert1a * .curr_rate1)
8490          Else
8500            .curr_convert1b = Null
8510          End If
8520        End Select
8530      End Select
8540      cmdSave_Click  ' ** Procedure: Above.
8550    End With

EXITP:
8560    Exit Sub

ERRH:
8570    Select Case ERR.Number
        Case Else
8580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8590    End Select
8600    Resume EXITP

End Sub

Private Sub curr_convert1a_KeyDown(KeyCode As Integer, Shift As Integer)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_convert1a_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8720    intShiftDown = (Shift And acShiftMask) > 0
8730    intAltDown = (Shift And acAltMask) > 0
8740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8770        With Me
8780          intRetVal = 0
8790          .curr_convert1b.SetFocus
8800        End With
8810      End Select
8820    End If

        ' ** Shift keys.
8830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8850        With Me
8860          intRetVal = 0
8870          .curr_date.SetFocus
8880        End With
8890      End Select
8900    End If

EXITP:
8910    KeyCode = intRetVal
8920    Exit Sub

ERRH:
8930    intRetVal = 0
8940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8950    Case Else
8960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8970    End Select
8980    Resume EXITP

End Sub

Private Sub curr_convert1b_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_convert1b_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9020    intShiftDown = (Shift And acShiftMask) > 0
9030    intAltDown = (Shift And acAltMask) > 0
9040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9070        With Me
9080          intRetVal = 0
9090          .curr_convert2a.SetFocus
9100        End With
9110      End Select
9120    End If

        ' ** Shift keys.
9130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9150        With Me
9160          intRetVal = 0
9170          .curr_convert1a.SetFocus
9180        End With
9190      End Select
9200    End If

EXITP:
9210    KeyCode = intRetVal
9220    Exit Sub

ERRH:
9230    intRetVal = 0
9240    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
9250    Case Else
9260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9270    End Select
9280    Resume EXITP

End Sub

Private Sub curr_convert2a_AfterUpdate()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_convert2a_AfterUpdate"

9310    With Me
9320      Select Case IsNull(.curr_convert2a)
          Case True
            ' ** Leave it.
9330      Case False
9340        Select Case IsNull(.curr_rate2)
            Case True
9350          .curr_convert2b = Null
9360        Case False
9370          If .curr_rate2 > 0 Then
9380            .curr_convert2b = (.curr_convert2a * .curr_rate2)
9390          Else
9400            .curr_convert2b = Null
9410          End If
9420        End Select
9430      End Select
9440      cmdSave_Click  ' ** Procedure: Above.
9450    End With

EXITP:
9460    Exit Sub

ERRH:
9470    Select Case ERR.Number
        Case Else
9480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9490    End Select
9500    Resume EXITP

End Sub

Private Sub curr_convert2a_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_convert2a_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9670        With Me
9680          intRetVal = 0
9690          .curr_convert2b.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9750        With Me
9760          intRetVal = 0
9770          .curr_convert1b.SetFocus
9780        End With
9790      End Select
9800    End If

EXITP:
9810    KeyCode = intRetVal
9820    Exit Sub

ERRH:
9830    intRetVal = 0
9840    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
9850    Case Else
9860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9870    End Select
9880    Resume EXITP

End Sub

Private Sub curr_convert2b_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_convert2b_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9920    intShiftDown = (Shift And acShiftMask) > 0
9930    intAltDown = (Shift And acAltMask) > 0
9940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9970        With Me
9980          intRetVal = 0
9990          .curr_notes.SetFocus
10000       End With
10010     End Select
10020   End If

        ' ** Shift keys.
10030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10050       With Me
10060         intRetVal = 0
10070         .curr_convert2a.SetFocus
10080       End With
10090     End Select
10100   End If

EXITP:
10110   KeyCode = intRetVal
10120   Exit Sub

ERRH:
10130   intRetVal = 0
10140   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
10150   Case Else
10160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10170   End Select
10180   Resume EXITP

End Sub

Private Sub curr_word1_lbl_cmd_DblClick(Cancel As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_word1_lbl_cmd_DblClick"

10210   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
10220   Exit Sub

ERRH:
10230   Select Case ERR.Number
        Case Else
10240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10250   End Select
10260   Resume EXITP

End Sub

Private Sub curr_notes_lbl_cmd_DblClick(Cancel As Integer)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_lbl_cmd_DblClick"

10310   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
10320   Exit Sub

ERRH:
10330   Select Case ERR.Number
        Case Else
10340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10350   End Select
10360   Resume EXITP

End Sub

Private Sub curr_notes_GotFocus()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_GotFocus"

10410   With Me
10420     .curr_notes.SelLength = 0
10430     .curr_notes.SelStart = 299
10440   End With

EXITP:
10450   Exit Sub

ERRH:
10460   Select Case ERR.Number
        Case Else
10470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10480   End Select
10490   Resume EXITP

End Sub

Private Sub curr_notes_Enter()

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_Enter"

        Dim blnRetVal As Boolean

10510   blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

EXITP:
10520   Exit Sub

ERRH:
10530   Select Case ERR.Number
        Case Else
10540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10550   End Select
10560   Resume EXITP

End Sub

Private Sub curr_notes_AfterUpdate()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_AfterUpdate"

10610   cmdSave_Click  ' ** Procedure: Above.

EXITP:
10620   Exit Sub

ERRH:
10630   Select Case ERR.Number
        Case Else
10640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10650   End Select
10660   Resume EXITP

End Sub

Private Sub curr_notes_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10720   intShiftDown = (Shift And acShiftMask) > 0
10730   intAltDown = (Shift And acAltMask) > 0
10740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab
10770       With Me
10780         intRetVal = 0
10790         .curr_active.SetFocus
10800       End With
10810     End Select
10820   End If

        ' ** Shift keys.
10830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10840     Select Case intRetVal
          Case vbKeyTab
10850       With Me
10860         intRetVal = 0
10870         .curr_convert2b.SetFocus
10880       End With
10890     End Select
10900   End If

EXITP:
10910   KeyCode = intRetVal
10920   Exit Sub

ERRH:
10930   intRetVal = 0
10940   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
10950   Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
10960   Case Else
10970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10980   End Select
10990   Resume EXITP

End Sub

Private Sub curr_notes_Exit(Cancel As Integer)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_Exit"

        Dim blnRetVal As Boolean

11010   blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

EXITP:
11020   Exit Sub

ERRH:
11030   Select Case ERR.Number
        Case Else
11040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11050   End Select
11060   Resume EXITP

End Sub

Private Sub curr_active_lbl_cmd_DblClick(Cancel As Integer)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_active_lbl_cmd_DblClick"

11110   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
11120   Exit Sub

ERRH:
11130   Select Case ERR.Number
        Case Else
11140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11150   End Select
11160   Resume EXITP

End Sub

Private Sub curr_active_AfterUpdate()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_active_AfterUpdate"

11210   cmdSave_Click  ' ** Procedure: Above.

EXITP:
11220   Exit Sub

ERRH:
11230   Select Case ERR.Number
        Case Else
11240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11250   End Select
11260   Resume EXITP

End Sub

Private Sub curr_active_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_active_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11320   intShiftDown = (Shift And acShiftMask) > 0
11330   intAltDown = (Shift And acAltMask) > 0
11340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11370       With Me
11380         intRetVal = 0
11390         lngRecsCur = RecCnt  ' ** Function: Below.
11400         If .CurrentRecord < lngRecsCur Then
11410           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
11420           .curr_name.SetFocus
11430         Else
11440           DoCmd.SelectObject acForm, .Parent.Name, False
11450           If .Parent.opgActive.Enabled = True Then
11460             .Parent.opgActive.SetFocus
11470           ElseIf .Parent.cmdPreview.Enabled = True Then
11480             .Parent.cmdPreview.SetFocus
11490           Else
11500             .Parent.opgActive.SetFocus
11510           End If
11520         End If
11530       End With
11540     End Select
11550   End If

        ' ** Shift keys.
11560   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11570     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11580       With Me
11590         intRetVal = 0
11600         .curr_notes.SetFocus
11610       End With
11620     End Select
11630   End If

EXITP:
11640   KeyCode = intRetVal
11650   Exit Sub

ERRH:
11660   intRetVal = 0
11670   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
11680   Case Else
11690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11700   End Select
11710   Resume EXITP

End Sub

Private Sub curr_fund_lbl_cmd_DblClick(Cancel As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_fund_lbl_cmd_DblClick"

11810   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
11820   Exit Sub

ERRH:
11830   Select Case ERR.Number
        Case Else
11840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11850   End Select
11860   Resume EXITP

End Sub

Private Sub curr_metal_lbl_cmd_DblClick(Cancel As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_metal_lbl_cmd_DblClick"

11910   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
11920   Exit Sub

ERRH:
11930   Select Case ERR.Number
        Case Else
11940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11950   End Select
11960   Resume EXITP

End Sub

Private Sub curr_bmu_lbl_cmd_DblClick(Cancel As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_bmu_lbl_cmd_DblClick"

12010   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
12020   Exit Sub

ERRH:
12030   Select Case ERR.Number
        Case Else
12040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12050   End Select
12060   Resume EXITP

End Sub

Private Sub curr_alt_lbl_cmd_DblClick(Cancel As Integer)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_alt_lbl_cmd_DblClick"

12110   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
12120   Exit Sub

ERRH:
12130   Select Case ERR.Number
        Case Else
12140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12150   End Select
12160   Resume EXITP

End Sub

Private Sub curr_unit_lbl_cmd_DblClick(Cancel As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_unit_lbl_cmd_DblClick"

12210   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
12220   Exit Sub

ERRH:
12230   Select Case ERR.Number
        Case Else
12240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12250   End Select
12260   Resume EXITP

End Sub

Public Function RecCnt() As Long

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

12310   lngRetVal = 0&

12320   With Me
12330     Set rst = .RecordsetClone
12340     With rst
12350       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
12360       Else
12370         .MoveLast
12380         lngRetVal = .RecordCount
12390       End If
12400       .Close
12410     End With
12420   End With

EXITP:
12430   Set rst = Nothing
12440   RecCnt = lngRetVal
12450   Exit Function

ERRH:
12460   lngRetVal = 0&
12470   Select Case ERR.Number
        Case Else
12480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12490   End Select
12500   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

12610   Select Case intWhere
        Case acCmdRecordsGoToFirst
12620     DoCmd.RunCommand acCmdRecordsGoToFirst
12630   Case acCmdRecordsGoToLast
12640     DoCmd.RunCommand acCmdRecordsGoToLast
12650   Case acCmdRecordsGoToPrevious
12660     DoCmd.RunCommand acCmdRecordsGoToPrevious
12670   Case acCmdRecordsGoToNext
12680     DoCmd.RunCommand acCmdRecordsGoToNext
12690   Case Else
12700     If IsMissing(varID) = False Then
12710       With Me
12720         Set rst = .RecordsetClone
12730         With rst
12740           .FindFirst "[curr_id] = " & varID
12750           If .NoMatch = False Then
12760             Me.Bookmark = .Bookmark
12770           End If
12780           .Close
12790         End With
12800       End With
12810     End If
12820   End Select

EXITP:
12830   Set rst = Nothing
12840   Exit Sub

ERRH:
12850   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
12860   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
12870   Case Else
12880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12890   End Select
12900   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [curr_name_sort]"

13010   With Me
13020     .Controls(strSortLbl2).Visible = False
13030     .Controls(strSortLine2).Visible = False
13040     .Controls(strSortLine3).Visible = False
13050     .Controls(strSortLine2).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
13060     .Controls(strSortLine2).Width = lngTpp
13070     If strProc = "Form_Load" Then
13080       strCalled = "curr_name"
13090       strSortNow = strSortOrig
13100       lngSortLbl2_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
13110       lngSortLbl2_Left = (.Controls(strCalled & "_lbl").Left - (6& * lngTpp))
13120       If lngSortLbl2_Left < 0& Then lngSortLbl2_Left = 0&
13130       lngSortLine2_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
13140       lngSortLine2_Left = .Controls(strCalled & "_lbl").Left
13150       If lngSortLine2_Left < 0& Then lngSortLine2_Left = 0&
13160       lngSortLine2_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
13170       .Controls(strSortLbl2).Top = lngSortLbl2_Top
13180       .Controls(strSortLbl2).Left = lngSortLbl2_Left
13190       .Controls(strSortLbl2).Caption = strArwUp
13200       .Controls(strSortLbl2).ForeColor = CLR_DKBLU
13210       .Controls(strSortLine2).Top = lngSortLine2_Top
13220       .Controls(strSortLine3).Top = (lngSortLine2_Top + lngTpp)
13230       .Controls(strSortLine2).Left = lngSortLine2_Left
13240       .Controls(strSortLine3).Left = lngSortLine2_Left
13250       .Controls(strSortLine2).Width = lngSortLine2_Width
13260       .Controls(strSortLine3).Width = lngSortLine2_Width
13270     Else
13280       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
13290       Select Case strCalled
            Case "curr_fund", "curr_metal", "curr_bmu", "curr_alt", "curr_unit"
13300         lngSortLbl2_Top = (.Controls(strCalled & "_lbl_off").Top - lngTpp)
13310         lngSortLbl2_Left = .Controls(strCalled & "_lbl_off").Left
13320         If lngSortLbl2_Left < 0& Then lngSortLbl2_Left = 0&
13330         lngSortLine2_Top = (.Controls(strCalled & "_lbl_off").Top - lngTpp)
13340         lngSortLine2_Left = .Controls(strCalled & "_lbl_off").Left
13350         If lngSortLine2_Left < 0& Then lngSortLine2_Left = 0&
13360         lngSortLine2_Width = (.Controls(strCalled & "_lbl_off").Width + lngTpp)
13370       Case Else
13380         lngSortLbl2_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
13390         lngSortLbl2_Left = .Controls(strCalled & "_lbl").Left
13400         If lngSortLbl2_Left < 0& Then lngSortLbl2_Left = 0&
13410         lngSortLine2_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
13420         lngSortLine2_Left = .Controls(strCalled & "_lbl").Left
13430         If lngSortLine2_Left < 0& Then lngSortLine2_Left = 0&
13440         lngSortLine2_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
13450       End Select
13460       .Controls(strSortLbl2).Top = lngSortLbl2_Top
13470       .Controls(strSortLbl2).Left = lngSortLbl2_Left
13480       .Controls(strSortLine2).Top = lngSortLine2_Top
13490       .Controls(strSortLine3).Top = (lngSortLine2_Top + lngTpp)
13500       .Controls(strSortLine2).Left = lngSortLine2_Left
13510       .Controls(strSortLine3).Left = lngSortLine2_Left
13520       .Controls(strSortLine2).Width = lngSortLine2_Width
13530       .Controls(strSortLine3).Width = lngSortLine2_Width
13540       If strCalled = "curr_name" Then
13550         strSortAsc = strSortOrig
13560         If strSortNow = strSortAsc Then
13570           strSortNow = "[curr_name_sort] DESC"
13580           .Controls(strSortLbl2).Caption = strArwDn
13590           .Controls(strSortLbl2).ForeColor = CLR_DKRED
13600         Else
13610           strSortNow = strSortAsc
13620           .Controls(strSortLbl2).Caption = strArwUp
13630           .Controls(strSortLbl2).ForeColor = CLR_DKBLU
13640         End If
13650         lngSortLbl2_Left = (lngSortLbl2_Left - (6& * lngTpp))
13660         .Controls(strSortLbl2).Left = lngSortLbl2_Left
13670       Else
13680         .Controls(strSortLbl2).Caption = strArwUp
13690         .Controls(strSortLbl2).ForeColor = CLR_DKBLU
13700         Select Case strCalled
              Case "curr_code"
13710           strSortAsc = "[curr_code]"
13720           lngSortLbl2_Left = (lngSortLbl2_Left - (2& * lngTpp))
13730           .Controls(strSortLbl2).Left = lngSortLbl2_Left
13740         Case "curr_iso"
13750           strSortAsc = "[curr_iso]"
13760           lngSortLbl2_Left = (lngSortLbl2_Left - (3& * lngTpp))
13770           .Controls(strSortLbl2).Left = lngSortLbl2_Left
13780         Case "curr_rate1"
13790           strSortAsc = "[curr_rate1], [curr_word1]" & strStdAsc
13800           lngSortLbl2_Left = (lngSortLbl2_Left + (53& * lngTpp))
13810           lngSortLine2_Width = (lngSortLine2_Width - (53& * lngTpp))
13820           lngSortLine2_Left = (lngSortLine2_Left + (53& * lngTpp))
13830           .Controls(strSortLbl2).Left = lngSortLbl2_Left
13840           .Controls(strSortLine2).Left = lngSortLine2_Left
13850           .Controls(strSortLine3).Left = lngSortLine2_Left
13860           .Controls(strSortLine2).Width = lngSortLine2_Width
13870           .Controls(strSortLine3).Width = lngSortLine2_Width
13880         Case "curr_rate2"
13890           strSortAsc = "[curr_rate2], [curr_word1]" & strStdAsc
13900           lngSortLbl2_Left = (lngSortLbl2_Left + (53& * lngTpp))
13910           lngSortLine2_Width = (lngSortLine2_Width - (53& * lngTpp))
13920           lngSortLine2_Left = (lngSortLine2_Left + (53& * lngTpp))
13930           .Controls(strSortLbl2).Left = lngSortLbl2_Left
13940           .Controls(strSortLine2).Left = lngSortLine2_Left
13950           .Controls(strSortLine3).Left = lngSortLine2_Left
13960           .Controls(strSortLine2).Width = lngSortLine2_Width
13970           .Controls(strSortLine3).Width = lngSortLine2_Width
13980         Case "curr_date"
13990           strSortAsc = "[curr_date]" & strStdAsc
14000           lngSortLbl2_Left = (lngSortLbl2_Left + (37& * lngTpp))
14010           lngSortLine2_Width = (lngSortLine2_Width - (37& * lngTpp))
14020           lngSortLine2_Left = (lngSortLine2_Left + (37& * lngTpp))
14030           .Controls(strSortLbl2).Left = lngSortLbl2_Left
14040           .Controls(strSortLine2).Left = lngSortLine2_Left
14050           .Controls(strSortLine3).Left = lngSortLine2_Left
14060           .Controls(strSortLine2).Width = lngSortLine2_Width
14070           .Controls(strSortLine3).Width = lngSortLine2_Width
14080         Case "curr_word1"
14090           strSortAsc = "[curr_word1]" & strStdAsc
14100           lngSortLbl2_Left = (lngSortLbl2_Left - (6& * lngTpp))
14110           .Controls(strSortLbl2).Left = lngSortLbl2_Left
14120         Case "curr_notes"
14130           strSortAsc = "[curr_notes]" & strStdAsc
14140         Case "curr_active"
14150           strSortAsc = "[curr_active]" & strStdAsc
14160           lngSortLbl2_Top = (lngSortLbl2_Top + lngTpp)
14170           lngSortLbl2_Left = (lngSortLbl2_Left - (3& * lngTpp))
14180           lngSortLine2_Top = (lngSortLine2_Top + lngTpp)
14190           lngSortLine2_Left = (lngSortLine2_Left - lngTpp)
14200           lngSortLine2_Width = (lngSortLine2_Width + (2& * lngTpp))
14210           .Controls(strSortLbl2).Top = lngSortLbl2_Top
14220           .Controls(strSortLbl2).Left = lngSortLbl2_Left
14230           .Controls(strSortLine2).Top = lngSortLine2_Top
14240           .Controls(strSortLine3).Top = (lngSortLine2_Top + lngTpp)
14250           .Controls(strSortLine2).Left = lngSortLine2_Left
14260           .Controls(strSortLine3).Left = lngSortLine2_Left
14270           .Controls(strSortLine2).Width = lngSortLine2_Width
14280           .Controls(strSortLine3).Width = lngSortLine2_Width
14290         Case "curr_fund"
14300           strSortAsc = "[curr_fund]" & strStdAsc
                'lngSortLbl2_Left = (lngSortLbl2_Left - (4& * lngTpp))  'Off
14310           lngSortLbl2_Left = (lngSortLbl2_Left - (7& * lngTpp))  'On
14320           If lngSortLbl2_Left < 0& Then lngSortLbl2_Left = 0&
14330           .Controls(strSortLbl2).Left = lngSortLbl2_Left
14340         Case "curr_metal"
14350           strSortAsc = "[curr_metal]" & strStdAsc
                'lngSortLbl2_Left = (lngSortLbl2_Left - (2& * lngTpp))  'Off
14360           lngSortLbl2_Left = (lngSortLbl2_Left - (5& * lngTpp))  'On
14370           .Controls(strSortLbl2).Left = lngSortLbl2_Left
14380         Case "curr_bmu"
14390           strSortAsc = "[curr_bmu]" & strStdAsc
14400           lngSortLbl2_Left = (lngSortLbl2_Left - (5& * lngTpp))  'Off/On
14410           .Controls(strSortLbl2).Left = lngSortLbl2_Left
14420         Case "curr_alt"
14430           strSortAsc = "[curr_alt]" & strStdAsc
                'lngSortLbl2_Left = (lngSortLbl2_Left - (3& * lngTpp))  'Off
14440           lngSortLbl2_Left = (lngSortLbl2_Left - (5& * lngTpp))  'On
14450           .Controls(strSortLbl2).Left = lngSortLbl2_Left
14460         Case "curr_unit"
14470           strSortAsc = "[curr_unit]" & strStdAsc
                'lngSortLbl2_Left = (lngSortLbl2_Left - (4& * lngTpp))  'Off
14480           lngSortLbl2_Left = (lngSortLbl2_Left - (5& * lngTpp))  'On
14490           .Controls(strSortLbl2).Left = lngSortLbl2_Left
14500         End Select
14510         If strSortNow = strSortAsc Then
14520           intCnt = CharCnt(strSortAsc, ",") + 1
14530           Select Case intCnt
                Case 1
14540             strTmp01 = strSortAsc & " DESC"
14550           Case 2
14560             intPos01 = InStr(strSortAsc, ",")
14570             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
14580             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
14590           Case 3
14600             intPos01 = InStr(strSortAsc, ",")
14610             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
14620             strTmp02 = Mid(strSortAsc, intPos01)
14630             intPos01 = InStr(2, strTmp02, ",")
14640             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & "DESC"
14650             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
14660           Case 4
14670             intPos01 = InStr(strSortAsc, ",")
14680             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
14690             strTmp02 = Mid(strSortAsc, intPos01)
14700             intPos01 = InStr(2, strTmp02, ",")
14710             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & "DESC"
14720             strTmp02 = Mid(strTmp02, intPos01)
14730             intPos01 = InStr(2, strTmp02, ",")
14740             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
14750             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
14760           End Select
14770           strSortNow = strTmp01
14780           .Controls(strSortLbl2).Caption = strArwDn
14790           .Controls(strSortLbl2).ForeColor = CLR_DKRED
14800         Else
14810           strSortNow = strSortAsc
14820         End If
14830       End If
14840     End If
14850     .Controls(strSortLbl2).Visible = True
14860     .Controls(strSortLine2).Visible = True
14870     .Controls(strSortLine3).Visible = True
14880     .OrderBy = strSortNow
14890     .OrderByOn = True
14900   End With

EXITP:
14910   Exit Sub

ERRH:
14920   Select Case ERR.Number
        Case Else
14930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14940   End Select
14950   Resume EXITP

End Sub

Public Function SortNow_Get() As String

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

15010   strRetVal = strSortNow

EXITP:
15020   SortNow_Get = strRetVal
15030   Exit Function

ERRH:
15040   strRetVal = vbNullString
15050   Select Case ERR.Number
        Case Else
15060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15070   End Select
15080   Resume EXITP

End Function
