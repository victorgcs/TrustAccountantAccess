VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_Sub"

'VGC 04/17/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Active:           A {curr_active}
' **   Close:            C {cmdClose on frmCurrency}
' **   Edit:             E {cmdEdit on frmCurrency}
' **   Excel:            L {cmdExcel on frmCurrency}
' **   Countries:        N {cmdCountries on frmCurrency}
' **   Word:             O {cmdWord on frmCurrency}
' **   Print:            P {cmdPrint on frmCurrency}
' **   Currency Word:    R {curr_word1}
' **   Reset:            S {cmdResetFilter on frmCurrency}
' **   Rate Date:        T {curr_date}
' **   Update:           U {cmdUpdate on frmCurrency}
' **   Preview:          W {cmdPreview on frmCurrency}
' **   Exit:             X {cmdClose on frmCurrency}
' **   Rate 1:           1 {curr_rate1}
' **   Rate 2:           2 {curr_rate2}
' **   Country:            {country_name1}

' ** Shortcut Ctrl keys responsive from this form:
' **   Exclude Active:   E {opgActive_optExclude on frmCurrency}
' **   Include Active:   I {opgActive_optInclude on frmCurrency}
' **   Only Active:      O {opgActive_optOnly on frmCurrency}
' **   Save:             S {cmdSave}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Exclude Funds:    E {opgFund_optExclude on frmCurrency}
' **   Include Funds:    I {opgFund_optInclude on frmCurrency}
' **   Only Metals:      L {opgMetal_optOnly on frmCurrency}
' **   Include Metals:   N {opgMetal_optInclude on frmCurrency}
' **   Only Funds:       O {opgFund_optOnly on frmCurrency}
' **   Exclude Metals:   X {opgMetal_optExclude on frmCurrency}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   Exclude Units:    D {opgUnit_optExclude on frmCurrency}
' **   Exclude BMUs:     E {opgBMU_optExclude on frmCurrency}
' **   Include BMUs:     I {opgBMU_optInclude on frmCurrency}
' **   Only Misc:        L {opgAlt_optOnly on frmCurrency}
' **   Include Misc:     N {opgAlt_optInclude on frmCurrency}
' **   Only BMUs:        O {opgBMU_optOnly on frmCurrency}
' **   Include Units:    U {opgUnit_optInclude on frmCurrency}
' **   Exclude Misc:     X {opgAlt_optExclude on frmCurrency}
' **   Only Units:       Y {opgUnit_optOnly on frmCurrency}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[curr_word1], [country_name_sort]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLine2 As String = "Sort_line2"
Private Const strSortLine3 As String = "Sort_line3"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strSortLbl2 As String = "Sort_lbl2"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strControlName As String
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
150       If strCallingForm = "frmCurrency" Then

160         lngSortLbl_Width = .Sort_lbl.Width

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
420       .curr_code.ForeColor = CLR_DISABLED_FG
430       .curr_code.BackColor = CLR_DISABLED_BG
440       .curr_word1.ForeColor = CLR_DISABLED_FG
450       .curr_word1.BackColor = CLR_DISABLED_BG
460       .curr_word2.ForeColor = CLR_DISABLED_FG
470       .curr_word2_box.BackColor = CLR_DISABLED_BG
480       .curr_name1.ForeColor = CLR_DISABLED_FG
490       .curr_name1.BackColor = CLR_DISABLED_BG
500       .curr_name2.ForeColor = CLR_DISABLED_FG
510       .curr_name2_box.BackColor = CLR_DISABLED_BG
520       .curr_iso.ForeColor = CLR_DISABLED_FG
530       .curr_iso.BackColor = CLR_DISABLED_BG
540       .curr_rate1.ForeColor = CLR_DISABLED_FG
550       .curr_rate1.BackColor = CLR_DISABLED_BG
560       .curr_rate2.ForeColor = CLR_DISABLED_FG
570       .curr_rate2.BackColor = CLR_DISABLED_BG
580       .curr_date.ForeColor = CLR_DISABLED_FG
590       .curr_date.BackColor = CLR_DISABLED_BG
600       .curr_decimal.ForeColor = CLR_DISABLED_FG
610       .curr_decimal.BackColor = CLR_DISABLED_BG
620       .curr_notes.ForeColor = CLR_DISABLED_FG
630       .curr_notes.BackColor = CLR_DISABLED_BG
640       .curr_active_box1.ForeColor = MY_CLR_BGE
650       .curr_active_box2.ForeColor = WIN_CLR_DISR
660       .country_code3.ForeColor = CLR_DISABLED_FG
670       .country_code3.BackColor = CLR_DISABLED_BG
680       .country_name.ForeColor = CLR_DISABLED_FG
690       .country_name.BackColor = CLR_DISABLED_BG
700       .country_iso.ForeColor = CLR_DISABLED_FG
710       .country_iso.BackColor = CLR_DISABLED_BG
720       .country_active_box1.ForeColor = MY_CLR_BGE
730       .country_active_box2.ForeColor = WIN_CLR_DISR
740     End With

750     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
760     Exit Sub

ERRH:
770     Select Case ERR.Number
        Case Else
780       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
790     End Select
800     Resume EXITP

End Sub

Private Sub Form_Current()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim strDocName As String

910     With Me
920       .Parent.country_cnt = .country_cnt
930       strDocName = "frmCurrency_Country"
940       If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
950         Forms(strDocName).cmbCurrencies = .curr_id
960       End If
970     End With

EXITP:
980     Exit Sub

ERRH:
990     Select Case ERR.Number
        Case Else
1000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1010    End Select
1020    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Active:           A {curr_active}
        ' **   Currency Word:    R {curr_word1}
        ' **   Rate Date:        T {curr_date}
        ' **   Exit:             X {cmdClose on frmCurrency}
        ' **   Rate 1:           1 {curr_rate1}
        ' **   Rate 2:           2 {curr_rate2}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Exclude Active:   E {opgActive_optExclude on frmCurrency}
        ' **   Include Active:   I {opgActive_optInclude on frmCurrency}
        ' **   Only Active:      O {opgActive_optOnly on frmCurrency}
        ' **   Save:             S {cmdSave}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Exclude Funds:    E {opgFund_optExclude on frmCurrency}
        ' **   Include Funds:    I {opgFund_optInclude on frmCurrency}
        ' **   Only Metals:      L {opgMetal_optOnly on frmCurrency}
        ' **   Include Metals:   N {opgMetal_optInclude on frmCurrency}
        ' **   Only Funds:       O {opgFund_optOnly on frmCurrency}
        ' **   Exclude Metals:   X {opgMetal_optExclude on frmCurrency}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   Exclude Units:    D {opgUnit_optExclude on frmCurrency}
        ' **   Exclude BMUs:     E {opgBMU_optExclude on frmCurrency}
        ' **   Include BMUs:     I {opgBMU_optInclude on frmCurrency}
        ' **   Only Misc:        L {opgAlt_optOnly on frmCurrency}
        ' **   Include Misc:     N {opgAlt_optInclude on frmCurrency}
        ' **   Only BMUs:        O {opgBMU_optOnly on frmCurrency}
        ' **   Include Units:    U {opgUnit_optInclude on frmCurrency}
        ' **   Exclude Misc:     X {opgAlt_optExclude on frmCurrency}
        ' **   Only Units:       Y {opgUnit_optOnly on frmCurrency}

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyUp
1170        intRetVal = 0
1180        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1190      Case vbKeyDown
1200        intRetVal = 0
1210        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1220      Case vbKeyEscape
1230        With Me
1240          intRetVal = 0
1250          .Parent.cmdClose_Click  ' ** Form Procedure: frmCurrency.
1260        End With
1270      End Select
1280    End If

        ' ** Alt keys.
1290    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1300      Select Case intRetVal
          Case vbKeyA
1310        With Me
1320          intRetVal = 0
1330          .curr_active.SetFocus
1340          If .curr_active.Enabled = True And .curr_active.Locked = False Then
1350            .curr_active = (Not .curr_active)
1360            curr_active_AfterUpdate  ' ** Procedure: Below.
1370          End If
1380        End With
1390      Case vbKeyR
1400        With Me
1410          intRetVal = 0
1420          .curr_word1.SetFocus
1430        End With
1440      Case vbKeyT
1450        With Me
1460          intRetVal = 0
1470          .curr_date.SetFocus
1480        End With
1490      Case vbKeyX
1500        With Me
1510          intRetVal = 0
1520          .Parent.cmdClose_Click  ' ** Form Procedure: frmCurrency.
1530        End With
1540      Case vbKey1
1550        With Me
1560          intRetVal = 0
1570          .curr_rate1.SetFocus
1580        End With
1590      Case vbKey2
1600        With Me
1610          intRetVal = 0
1620          .curr_rate2.SetFocus
1630        End With
1640      End Select
1650    End If

        ' ** Ctrl keys.
1660    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1670      Select Case intRetVal
          Case vbKeyE
1680        With Me
1690          intRetVal = 0
1700          DoCmd.SelectObject acForm, .Parent.Name, False
1710          .Parent.opgActive.SetFocus
1720          .Parent.opgActive = .Parent.opgActive_optExclude.OptionValue
1730          .Parent.opgActive_AfterUpdate  ' ** Form Procedure: frmCurrency.
1740        End With
1750      Case vbKeyI
1760        With Me
1770          intRetVal = 0
1780          DoCmd.SelectObject acForm, .Parent.Name, False
1790          .Parent.opgActive.SetFocus
1800          .Parent.opgActive = .Parent.opgActive_optInclude.OptionValue
1810          .Parent.opgActive_AfterUpdate  ' ** Form Procedure: frmCurrency.
1820        End With
1830      Case vbKeyO
1840        With Me
1850          intRetVal = 0
1860          DoCmd.SelectObject acForm, .Parent.Name, False
1870          .Parent.opgActive.SetFocus
1880          .Parent.opgActive = .Parent.opgActive_optOnly.OptionValue
1890          .Parent.opgActive_AfterUpdate  ' ** Form Procedure: frmCurrency.
1900        End With
1910      Case vbKeyS
1920        With Me
1930          intRetVal = 0
1940          If .curr_active.Enabled = True And .curr_active.Locked = False Then
1950            cmdSave_Click  ' ** Procedure: Below.
1960          End If
1970        End With
1980      Case vbKeyTab, vbKeyReturn
1990        With Me
2000          intRetVal = 0
2010          DoCmd.SelectObject acForm, .Parent.Name, False
2020          .Parent.opgActive.SetFocus
2030        End With
2040      Case vbKeyUp, vbKeyPageUp
2050        intRetVal = 0
2060        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2070      Case vbKeyDown, vbKeyPageDown
2080        intRetVal = 0
2090        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2100      End Select
2110    End If

        ' ** Ctrl-Shift keys.
2120    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2130      Select Case intRetVal
          Case vbKeyE
2140        With Me
2150          intRetVal = 0
2160          DoCmd.SelectObject acForm, .Parent.Name, False
2170          .Parent.opgFund.SetFocus
2180          .Parent.opgFund = .Parent.opgFund_optExclude.OptionValue
2190          .Parent.opgFund_AfterUpdate  ' ** Form Procedure: frmCurrency.
2200        End With
2210      Case vbKeyF
2220        With Me
2230          intRetVal = 0
2240          DoCmd.SelectObject acForm, .Parent.Name, False
2250          .Parent.FocusHolder.SetFocus
2260        End With
2270      Case vbKeyI
2280        With Me
2290          intRetVal = 0
2300          DoCmd.SelectObject acForm, .Parent.Name, False
2310          .Parent.opgFund.SetFocus
2320          .Parent.opgFund = .Parent.opgFund_optInclude.OptionValue
2330          .Parent.opgFund_AfterUpdate  ' ** Form Procedure: frmCurrency.
2340        End With
2350      Case vbKeyL
2360        With Me
2370          intRetVal = 0
2380          DoCmd.SelectObject acForm, .Parent.Name, False
2390          .Parent.opgMetal.SetFocus
2400          .Parent.opgMetal = .Parent.opgMetal_optOnly.OptionValue
2410          .Parent.opgMetal_AfterUpdate  ' ** Form Procedure: frmCurrency.
2420        End With
2430      Case vbKeyN
2440        With Me
2450          intRetVal = 0
2460          DoCmd.SelectObject acForm, .Parent.Name, False
2470          .Parent.opgMetal.SetFocus
2480          .Parent.opgMetal = .Parent.opgMetal_optInclude.OptionValue
2490          .Parent.opgMetal_AfterUpdate  ' ** Form Procedure: frmCurrency.
2500        End With
2510      Case vbKeyO
2520        With Me
2530          intRetVal = 0
2540          DoCmd.SelectObject acForm, .Parent.Name, False
2550          .Parent.opgFund.SetFocus
2560          .Parent.opgFund = .Parent.opgFund_optOnly.OptionValue
2570          .Parent.opgFund_AfterUpdate  ' ** Form Procedure: frmCurrency.
2580        End With
2590      Case vbKeyX
2600        With Me
2610          intRetVal = 0
2620          DoCmd.SelectObject acForm, .Parent.Name, False
2630          .Parent.opgMetal.SetFocus
2640          .Parent.opgMetal = .Parent.opgMetal_optExclude.OptionValue
2650          .Parent.opgMetal_AfterUpdate  ' ** Form Procedure: frmCurrency.
2660        End With
2670      Case vbKeyTab, vbKeyReturn
2680        With Me
2690          intRetVal = 0
2700          DoCmd.SelectObject acForm, .Parent.Name, False
2710          .Parent.cmdClose.SetFocus
2720        End With
2730      End Select
2740    End If

        ' ** Ctrl-Alt keys.
2750    If intCtrlDown And intAltDown And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyD
2770        With Me
2780          intRetVal = 0
2790          DoCmd.SelectObject acForm, .Parent.Name, False
2800          .Parent.opgUnit.SetFocus
2810          .Parent.opgUnit = .Parent.opgUnit_optExclude.OptionValue
2820          .Parent.opgUnit_AfterUpdate  ' ** Form Procedure: frmCurrency.
2830        End With
2840      Case vbKeyE
2850        With Me
2860          intRetVal = 0
2870          DoCmd.SelectObject acForm, .Parent.Name, False
2880          .Parent.opgBMU.SetFocus
2890          .Parent.opgBMU = .Parent.opgBMU_optExclude.OptionValue
2900          .Parent.opgBMU_AfterUpdate  ' ** Form Procedure: frmCurrency.
2910        End With
2920      Case vbKeyI
2930        With Me
2940          intRetVal = 0
2950          DoCmd.SelectObject acForm, .Parent.Name, False
2960          .Parent.opgBMU.SetFocus
2970          .Parent.opgBMU = .Parent.opgBMU_optInclude.OptionValue
2980          .Parent.opgBMU_AfterUpdate  ' ** Form Procedure: frmCurrency.
2990        End With
3000      Case vbKeyL
3010        With Me
3020          intRetVal = 0
3030          DoCmd.SelectObject acForm, .Parent.Name, False
3040          .Parent.opgAlt.SetFocus
3050          .Parent.opgAlt = .Parent.opgAlt_optOnly.OptionValue
3060          .Parent.opgAlt_AfterUpdate  ' ** Form Procedure: frmCurrency.
3070        End With
3080      Case vbKeyN
3090        With Me
3100          intRetVal = 0
3110          DoCmd.SelectObject acForm, .Parent.Name, False
3120          .Parent.opgAlt.SetFocus
3130          .Parent.opgAlt = .Parent.opgAlt_optInclude.OptionValue
3140          .Parent.opgAlt_AfterUpdate  ' ** Form Procedure: frmCurrency.
3150        End With
3160      Case vbKeyO
3170        With Me
3180          intRetVal = 0
3190          DoCmd.SelectObject acForm, .Parent.Name, False
3200          .Parent.opgBMU.SetFocus
3210          .Parent.opgBMU = .Parent.opgBMU_optOnly.OptionValue
3220          .Parent.opgBMU_AfterUpdate  ' ** Form Procedure: frmCurrency.
3230        End With
3240      Case vbKeyU
3250        With Me
3260          intRetVal = 0
3270          DoCmd.SelectObject acForm, .Parent.Name, False
3280          .Parent.opgUnit.SetFocus
3290          .Parent.opgUnit = .Parent.opgUnit_optInclude.OptionValue
3300          .Parent.opgUnit_AfterUpdate  ' ** Form Procedure: frmCurrency.
3310        End With
3320      Case vbKeyX
3330        With Me
3340          intRetVal = 0
3350          DoCmd.SelectObject acForm, .Parent.Name, False
3360          .Parent.opgAlt.SetFocus
3370          .Parent.opgAlt = .Parent.opgAlt_optExclude.OptionValue
3380          .Parent.opgAlt_AfterUpdate  ' ** Form Procedure: frmCurrency.
3390        End With
3400      Case vbKeyY
3410        With Me
3420          intRetVal = 0
3430          DoCmd.SelectObject acForm, .Parent.Name, False
3440          .Parent.opgUnit.SetFocus
3450          .Parent.opgUnit = .Parent.opgUnit_optOnly.OptionValue
3460          .Parent.opgUnit_AfterUpdate  ' ** Form Procedure: frmCurrency.
3470        End With
3480      End Select
3490    End If

EXITP:
3500    KeyCode = intRetVal
3510    Exit Sub

ERRH:
3520    intRetVal = 0
3530    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3540    Case Else
3550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3560    End Select
3570    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

3610    With Me
3620      .curr_datemodified = Now()
3630      .currstage_datemodified = Now()
3640      .curr_username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
3650      .currstage_username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
3660      .currstage_changed = True
3670      If IsNull(.curr_rate1) = True Then
3680        .curr_rate1 = 0#
3690      End If
3700      If IsNull(.curr_rate2) = True Then
3710        .curr_rate2 = 0#
3720      End If
3730      .Parent.currstage_changed = True
3740      If .Parent.cmdUpdate.Enabled = False Then
3750        .Parent.cmdUpdate.Enabled = True
3760      End If
3770    End With

EXITP:
3780    Exit Sub

ERRH:
3790    Select Case ERR.Number
        Case Else
3800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3810    End Select
3820    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3910    With Me

3920      DoCmd.RunCommand acCmdSelectRecord
3930      DoCmd.RunCommand acCmdSaveRecord
3940      DoEvents

3950      If strControlName <> vbNullString Then
            ' ** Keep an eye out in case it doesn't come from an AfterUpdate() event.
3960        strControlName = Left(strControlName, (Len(strControlName) - Len("_AfterUpdate")))
3970        Select Case strControlName
            Case "curr_rate1"
3980          strControlName = "curr_rate2"
3990        Case "curr_rate2"
4000          strControlName = "curr_date"
4010        Case "curr_date"
4020          strControlName = "curr_active"
4030        Case "curr_active"
4040          strControlName = "curr_decimal"
4050        Case "curr_notes"
4060          strControlName = "country_code3"
4070        Case "country_active"
4080          strControlName = "curr_code"
4090        End Select
4100        If strControlName = "curr_code" Then
4110          lngRecsCur = RecCnt  ' ** Function: Below
4120          If .CurrentRecord < lngRecsCur Then
4130            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
4140            DoCmd.GoToControl strControlName
4150          Else
4160            DoCmd.SelectObject acForm, .Parent.Name, False
4170            If .Parent.cmdUpdate.Enabled = True Then
4180              .Parent.cmdUpdate.SetFocus
4190            Else
                  ' ** cmdUpdate should have been enabled!
4200              .Parent.cmdUpdate.Enabled = True
4210              .Parent.cmdUpdate.SetFocus
4220            End If
4230          End If
4240        Else
4250          DoCmd.GoToControl strControlName
4260        End If
4270        strControlName = vbNullString
4280      End If

4290    End With

EXITP:
4300    Exit Sub

ERRH:
4310    Select Case ERR.Number
        Case Else
4320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4330    End Select
4340    Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

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
4510            .curr_code.SetFocus
4520          Else
4530            DoCmd.SelectObject acForm, .Parent.Name, False
4540            .Parent.opgActive.SetFocus
4550          End If
4560        End With
4570      End Select
4580    End If

        ' ** Shift keys.
4590    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4600      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4610        With Me
4620          intRetVal = 0
4630          If .CurrentRecord > 1 Then
4640            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4650            .country_active.SetFocus
4660          Else
4670            DoCmd.SelectObject acForm, .Parent.Name, False
4680            .Parent.cmdClose.SetFocus
4690          End If
4700        End With
4710      End Select
4720    End If

EXITP:
4730    KeyCode = intRetVal
4740    Exit Sub

ERRH:
4750    intRetVal = 0
4760    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4770    Case Else
4780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4790    End Select
4800    Resume EXITP

End Sub

Private Sub curr_code_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_lbl_DblClick"

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

Private Sub curr_code_GotFocus()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_GotFocus"

5010    With Me
5020      .curr_code.SelLength = 0
5030      .curr_code.SelStart = 9
5040    End With

EXITP:
5050    Exit Sub

ERRH:
5060    Select Case ERR.Number
        Case Else
5070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5080    End Select
5090    Resume EXITP

End Sub

Private Sub curr_code_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_code_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5170        With Me
5180          intRetVal = 0
5190          .curr_word1.SetFocus
5200        End With
5210      End Select
5220    End If

        ' ** Shift keys.
5230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5250        With Me
5260          intRetVal = 0
5270          If .CurrentRecord > 1 Then
5280            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
5290            .country_active.SetFocus
5300          Else
5310            DoCmd.SelectObject acForm, .Parent.Name, False
5320            .Parent.cmdClose.SetFocus
5330          End If
5340        End With
5350      End Select
5360    End If

EXITP:
5370    KeyCode = intRetVal
5380    Exit Sub

ERRH:
5390    intRetVal = 0
5400    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5410    Case Else
5420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5430    End Select
5440    Resume EXITP

End Sub

Private Sub curr_word1_lbl_DblClick(Cancel As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_word1_lbl_DblClick"

5510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5520    Exit Sub

ERRH:
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub

Private Sub curr_word1_GotFocus()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_word1_GotFocus"

5610    With Me
5620      .curr_word1.SelLength = 0
5630      .curr_word1.SelStart = 99
5640    End With

EXITP:
5650    Exit Sub

ERRH:
5660    Select Case ERR.Number
        Case Else
5670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5680    End Select
5690    Resume EXITP

End Sub

Private Sub curr_word1_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_word1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5720    intShiftDown = (Shift And acShiftMask) > 0
5730    intAltDown = (Shift And acAltMask) > 0
5740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5770        With Me
5780          intRetVal = 0
5790          .curr_name1.SetFocus
5800        End With
5810      End Select
5820    End If

        ' ** Shift keys.
5830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5850        With Me
5860          intRetVal = 0
5870          .curr_code.SetFocus
5880        End With
5890      End Select
5900    End If

EXITP:
5910    KeyCode = intRetVal
5920    Exit Sub

ERRH:
5930    intRetVal = 0
5940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5950    Case Else
5960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5970    End Select
5980    Resume EXITP

End Sub

Private Sub curr_name1_lbl_DblClick(Cancel As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name1_lbl_DblClick"

6010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
6020    Exit Sub

ERRH:
6030    Select Case ERR.Number
        Case Else
6040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6050    End Select
6060    Resume EXITP

End Sub

Private Sub curr_name1_GotFocus()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name1_GotFocus"

6110    With Me
6120      .curr_name1.SelLength = 0
6130      .curr_name1.SelStart = 99
6140    End With

EXITP:
6150    Exit Sub

ERRH:
6160    Select Case ERR.Number
        Case Else
6170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6180    End Select
6190    Resume EXITP

End Sub

Private Sub curr_name1_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_name1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6220    intShiftDown = (Shift And acShiftMask) > 0
6230    intAltDown = (Shift And acAltMask) > 0
6240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6270        With Me
6280          intRetVal = 0
6290          .curr_iso.SetFocus
6300        End With
6310      End Select
6320    End If

        ' ** Shift keys.
6330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6350        With Me
6360          intRetVal = 0
6370          .curr_word1.SetFocus
6380        End With
6390      End Select
6400    End If

EXITP:
6410    KeyCode = intRetVal
6420    Exit Sub

ERRH:
6430    intRetVal = 0
6440    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6450    Case Else
6460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6470    End Select
6480    Resume EXITP

End Sub

Private Sub curr_iso_lbl_DblClick(Cancel As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_iso_lbl_DblClick"

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

Private Sub curr_iso_GotFocus()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_iso_GotFocus"

6610    With Me
6620      .curr_iso.SelLength = 0
6630      .curr_iso.SelStart = 9
6640    End With

EXITP:
6650    Exit Sub

ERRH:
6660    Select Case ERR.Number
        Case Else
6670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6680    End Select
6690    Resume EXITP

End Sub

Private Sub curr_iso_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_iso_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6720    intShiftDown = (Shift And acShiftMask) > 0
6730    intAltDown = (Shift And acAltMask) > 0
6740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6770        With Me
6780          intRetVal = 0
6790          .curr_rate1.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6850        With Me
6860          intRetVal = 0
6870          .curr_name1.SetFocus
6880        End With
6890      End Select
6900    End If

EXITP:
6910    KeyCode = intRetVal
6920    Exit Sub

ERRH:
6930    intRetVal = 0
6940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6950    Case Else
6960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6970    End Select
6980    Resume EXITP

End Sub

Private Sub curr_rate1_lbl_DblClick(Cancel As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_lbl_DblClick"

7010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Private Sub curr_rate1_GotFocus()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_GotFocus"

7110    With Me
7120      If .curr_rate1.Locked = True Then  ' ** If editable, let it highlight the whole thing.
7130        .curr_rate1.SelLength = 0
7140        .curr_rate1.SelStart = 9
7150      End If
7160    End With

EXITP:
7170    Exit Sub

ERRH:
7180    Select Case ERR.Number
        Case Else
7190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7200    End Select
7210    Resume EXITP

End Sub

Private Sub curr_rate1_BeforeUpdate(Cancel As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_BeforeUpdate"

7310    With Me
7320      Select Case IsNull(.curr_rate1)
          Case True
            ' ** We'll put in a Zero later.
7330      Case False
7340        If .curr_rate1 < 0# Then
7350          MsgBox "Rate must be a positive value.", vbInformation + vbOKOnly, "Invalid Entry"
7360          Cancel = -1
7370        End If
7380      End Select
7390    End With

EXITP:
7400    Exit Sub

ERRH:
7410    Select Case ERR.Number
        Case Else
7420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7430    End Select
7440    Resume EXITP

End Sub

Private Sub curr_rate1_AfterUpdate()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_AfterUpdate"

7510    strControlName = THIS_PROC
7520    cmdSave_Click  ' ** Procedure: Above.

EXITP:
7530    Exit Sub

ERRH:
7540    Select Case ERR.Number
        Case Else
7550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7560    End Select
7570    Resume EXITP

End Sub

Private Sub curr_rate1_KeyDown(KeyCode As Integer, Shift As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7620    intShiftDown = (Shift And acShiftMask) > 0
7630    intAltDown = (Shift And acAltMask) > 0
7640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7670        With Me
7680          intRetVal = 0
7690          If .curr_rate1 >= 0# Then
7700            .curr_rate2.SetFocus
7710          End If
7720        End With
7730      End Select
7740    End If

        ' ** Shift keys.
7750    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7770        With Me
7780          intRetVal = 0
7790          .curr_iso.SetFocus
7800        End With
7810      End Select
7820    End If

EXITP:
7830    KeyCode = intRetVal
7840    Exit Sub

ERRH:
7850    intRetVal = 0
7860    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
7870    Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
7880    Case Else
7890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7900    End Select
7910    Resume EXITP

End Sub

Private Sub curr_rate2_lbl_DblClick(Cancel As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_lbl_DblClick"

8010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8020    Exit Sub

ERRH:
8030    Select Case ERR.Number
        Case Else
8040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8050    End Select
8060    Resume EXITP

End Sub

Private Sub curr_rate2_GotFocus()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_GotFocus"

8110    With Me
8120      If .curr_rate2.Locked = True Then  ' ** If editable, let it highlight the whole thing.
8130        .curr_rate2.SelLength = 0
8140        .curr_rate2.SelStart = 9
8150      End If
8160    End With

EXITP:
8170    Exit Sub

ERRH:
8180    Select Case ERR.Number
        Case Else
8190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8200    End Select
8210    Resume EXITP

End Sub

Private Sub curr_rate2_BeforeUpdate(Cancel As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_BeforeUpdate"

8310    With Me
8320      Select Case IsNull(.curr_rate2)
          Case True
            ' ** We'll put in a Zero later.
8330      Case False
8340        If .curr_rate2 < 0# Then
8350          MsgBox "Rate must be a positive value.", vbInformation + vbOKOnly, "Invalid Entry"
8360          Cancel = -1
8370        End If
8380      End Select
8390    End With

EXITP:
8400    Exit Sub

ERRH:
8410    Select Case ERR.Number
        Case Else
8420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8430    End Select
8440    Resume EXITP

End Sub

Private Sub curr_rate2_AfterUpdate()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_AfterUpdate"

8510    strControlName = THIS_PROC
8520    cmdSave_Click  ' ** Procedure: Above.

EXITP:
8530    Exit Sub

ERRH:
8540    Select Case ERR.Number
        Case Else
8550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8560    End Select
8570    Resume EXITP

End Sub

Private Sub curr_rate2_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_rate2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8620    intShiftDown = (Shift And acShiftMask) > 0
8630    intAltDown = (Shift And acAltMask) > 0
8640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8670        With Me
8680          intRetVal = 0
8690          If .curr_rate2 >= 0# Then
8700            .curr_date.SetFocus
8710          End If
8720        End With
8730      End Select
8740    End If

        ' ** Shift keys.
8750    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8770        With Me
8780          intRetVal = 0
8790          .curr_rate1.SetFocus
8800        End With
8810      End Select
8820    End If

EXITP:
8830    KeyCode = intRetVal
8840    Exit Sub

ERRH:
8850    intRetVal = 0
8860    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8870    Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
8880    Case Else
8890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8900    End Select
8910    Resume EXITP

End Sub

Private Sub curr_date_lbl_DblClick(Cancel As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_lbl_DblClick"

9010    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
9020    Exit Sub

ERRH:
9030    Select Case ERR.Number
        Case Else
9040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9050    End Select
9060    Resume EXITP

End Sub

Private Sub curr_date_GotFocus()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_GotFocus"

9110    With Me
9120      .curr_date.SelLength = 0
9130      .curr_date.SelStart = 0
9140    End With

EXITP:
9150    Exit Sub

ERRH:
9160    Select Case ERR.Number
        Case Else
9170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9180    End Select
9190    Resume EXITP

End Sub

Private Sub curr_date_BeforeUpdate(Cancel As Integer)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_BeforeUpdate"

9210    With Me
9220      Select Case IsNull(.curr_date)
          Case True
            ' ** Leave it Null.
9230      Case False
9240        If IsDate(.curr_date) = False Then
9250          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
9260          .curr_date = Null
9270          Cancel = -1
9280        End If
9290      End Select
9300    End With

EXITP:
9310    Exit Sub

ERRH:
9320    Select Case ERR.Number
        Case Else
9330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9340    End Select
9350    Resume EXITP

End Sub

Private Sub curr_date_AfterUpdate()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_AfterUpdate"

9410    strControlName = THIS_PROC
9420    cmdSave_Click  ' ** Procedure: Above.

EXITP:
9430    Exit Sub

ERRH:
9440    Select Case ERR.Number
        Case Else
9450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9460    End Select
9470    Resume EXITP

End Sub

Private Sub curr_date_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_date_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9520    intShiftDown = (Shift And acShiftMask) > 0
9530    intAltDown = (Shift And acAltMask) > 0
9540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9570        With Me
9580          intRetVal = 0
9590          Select Case IsNull(.curr_date)
              Case True
9600            .curr_active.SetFocus
9610          Case False
9620            If IsDate(.curr_date) = True Then
9630              .curr_active.SetFocus
9640            End If
9650          End Select
9660        End With
9670      End Select
9680    End If

        ' ** Shift keys.
9690    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9700      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9710        With Me
9720          intRetVal = 0
9730          .curr_rate2.SetFocus
9740        End With
9750      End Select
9760    End If

EXITP:
9770    KeyCode = intRetVal
9780    Exit Sub

ERRH:
9790    intRetVal = 0
9800    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
9810    Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
9820    Case Else
9830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9840    End Select
9850    Resume EXITP

End Sub

Private Sub curr_active_lbl_DblClick(Cancel As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_active_lbl_DblClick"

9910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
9920    Exit Sub

ERRH:
9930    Select Case ERR.Number
        Case Else
9940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9950    End Select
9960    Resume EXITP

End Sub

Public Sub curr_active_AfterUpdate()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_active_AfterUpdate"

10010   With Me
10020     Select Case .curr_active
          Case True
10030       If .country_active = False Then
10040         .country_active = True
10050       End If
10060     Case False
            ' ** Leave country_active alone.
10070     End Select
10080     strControlName = THIS_PROC
10090     cmdSave_Click  ' ** Procedure: Above.
10100   End With

EXITP:
10110   Exit Sub

ERRH:
10120   Select Case ERR.Number
        Case Else
10130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10140   End Select
10150   Resume EXITP

End Sub

Private Sub curr_active_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_active_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10220   intShiftDown = (Shift And acShiftMask) > 0
10230   intAltDown = (Shift And acAltMask) > 0
10240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10270       With Me
10280         intRetVal = 0
10290         .curr_decimal.SetFocus
10300       End With
10310     End Select
10320   End If

        ' ** Shift keys.
10330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10350       With Me
10360         intRetVal = 0
10370         .curr_date.SetFocus
10380       End With
10390     End Select
10400   End If

EXITP:
10410   KeyCode = intRetVal
10420   Exit Sub

ERRH:
10430   intRetVal = 0
10440   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
10450   Case Else
10460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10470   End Select
10480   Resume EXITP

End Sub

Private Sub curr_decimal_lbl_cmd_DblClick(Cancel As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_decimal_lbl_cmd_DblClick"

10510   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
10520   Exit Sub

ERRH:
10530   Select Case ERR.Number
        Case Else
10540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10550   End Select
10560   Resume EXITP

End Sub

Private Sub curr_decimal_GotFocus()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_decimal_GotFocus"

10610   With Me
10620     .curr_decimal.SelLength = 0
10630     .curr_decimal.SelStart = 9
10640   End With

EXITP:
10650   Exit Sub

ERRH:
10660   Select Case ERR.Number
        Case Else
10670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10680   End Select
10690   Resume EXITP

End Sub

Private Sub curr_decimal_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_decimal_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
10770       With Me
10780         intRetVal = 0
10790         .curr_notes.SetFocus
10800       End With
10810     End Select
10820   End If

        ' ** Shift keys.
10830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10850       With Me
10860         intRetVal = 0
10870         .curr_active.SetFocus
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
10950   Case Else
10960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10970   End Select
10980   Resume EXITP

End Sub

Private Sub curr_notes_lbl_cmd_DblClick(Cancel As Integer)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_lbl_cmd_DblClick"

11010   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
11020   Exit Sub

ERRH:
11030   Select Case ERR.Number
        Case Else
11040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11050   End Select
11060   Resume EXITP

End Sub

Private Sub curr_notes_GotFocus()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_GotFocus"

11110   With Me
11120     .curr_notes.SelLength = 0
11130     .curr_notes.SelStart = 0
11140   End With

EXITP:
11150   Exit Sub

ERRH:
11160   Select Case ERR.Number
        Case Else
11170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11180   End Select
11190   Resume EXITP

End Sub

Private Sub curr_notes_AfterUpdate()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_AfterUpdate"

11210   strControlName = THIS_PROC
11220   cmdSave_Click  ' ** Procedure: Above.

EXITP:
11230   Exit Sub

ERRH:
11240   Select Case ERR.Number
        Case Else
11250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11260   End Select
11270   Resume EXITP

End Sub

Private Sub curr_notes_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_notes_KeyDown"

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
11390         .country_code3.SetFocus
11400       End With
11410     End Select
11420   End If

        ' ** Shift keys.
11430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11450       With Me
11460         intRetVal = 0
11470         .curr_decimal.SetFocus
11480       End With
11490     End Select
11500   End If

EXITP:
11510   KeyCode = intRetVal
11520   Exit Sub

ERRH:
11530   intRetVal = 0
11540   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
11550   Case Else
11560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11570   End Select
11580   Resume EXITP

End Sub

Private Sub country_code3_lbl_DblClick(Cancel As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "country_code3_lbl_DblClick"

11610   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
11620   Exit Sub

ERRH:
11630   Select Case ERR.Number
        Case Else
11640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11650   End Select
11660   Resume EXITP

End Sub

Private Sub country_code3_GotFocus()

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "country_code3_GotFocus"

11710   With Me
11720     .country_code3.SelLength = 0
11730     .country_code3.SelStart = 9
11740   End With

EXITP:
11750   Exit Sub

ERRH:
11760   Select Case ERR.Number
        Case Else
11770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11780   End Select
11790   Resume EXITP

End Sub

Private Sub country_code3_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "country_code3_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11820   intShiftDown = (Shift And acShiftMask) > 0
11830   intAltDown = (Shift And acAltMask) > 0
11840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11870       With Me
11880         intRetVal = 0
11890         .country_name.SetFocus
11900       End With
11910     End Select
11920   End If

        ' ** Shift keys.
11930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11950       With Me
11960         intRetVal = 0
11970         .curr_notes.SetFocus
11980       End With
11990     End Select
12000   End If

EXITP:
12010   KeyCode = intRetVal
12020   Exit Sub

ERRH:
12030   intRetVal = 0
12040   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
12050   Case Else
12060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12070   End Select
12080   Resume EXITP

End Sub

Private Sub country_name_lbl_DblClick(Cancel As Integer)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "country_name_lbl_DblClick"

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

Private Sub country_name_GotFocus()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "country_name_GotFocus"

12210   With Me
12220     .country_name.SelLength = 0
12230     .country_name.SelStart = 0
12240   End With

EXITP:
12250   Exit Sub

ERRH:
12260   Select Case ERR.Number
        Case Else
12270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12280   End Select
12290   Resume EXITP

End Sub

Private Sub country_name_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "country_name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12320   intShiftDown = (Shift And acShiftMask) > 0
12330   intAltDown = (Shift And acAltMask) > 0
12340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12370       With Me
12380         intRetVal = 0
12390         .country_iso.SetFocus
12400       End With
12410     End Select
12420   End If

        ' ** Shift keys.
12430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12450       With Me
12460         intRetVal = 0
12470         .country_code3.SetFocus
12480       End With
12490     End Select
12500   End If

EXITP:
12510   KeyCode = intRetVal
12520   Exit Sub

ERRH:
12530   intRetVal = 0
12540   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
12550   Case Else
12560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12570   End Select
12580   Resume EXITP

End Sub

Private Sub country_iso_lbl_DblClick(Cancel As Integer)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "country_iso_lbl_DblClick"

12610   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
12620   Exit Sub

ERRH:
12630   Select Case ERR.Number
        Case Else
12640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12650   End Select
12660   Resume EXITP

End Sub

Private Sub country_iso_GotFocus()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "country_iso_GotFocus"

12710   With Me
12720     .country_iso.SelLength = 0
12730     .country_iso.SelStart = 9
12740   End With

EXITP:
12750   Exit Sub

ERRH:
12760   Select Case ERR.Number
        Case Else
12770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12780   End Select
12790   Resume EXITP

End Sub

Private Sub country_iso_KeyDown(KeyCode As Integer, Shift As Integer)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "country_iso_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12820   intShiftDown = (Shift And acShiftMask) > 0
12830   intAltDown = (Shift And acAltMask) > 0
12840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12870       With Me
12880         intRetVal = 0
12890         .country_active.SetFocus
12900       End With
12910     End Select
12920   End If

        ' ** Shift keys.
12930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12950       With Me
12960         intRetVal = 0
12970         .country_name.SetFocus
12980       End With
12990     End Select
13000   End If

EXITP:
13010   KeyCode = intRetVal
13020   Exit Sub

ERRH:
13030   intRetVal = 0
13040   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
13050   Case Else
13060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13070   End Select
13080   Resume EXITP

End Sub

Private Sub country_active_lbl_DblClick(Cancel As Integer)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "country_active_lbl_DblClick"

13110   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
13120   Exit Sub

ERRH:
13130   Select Case ERR.Number
        Case Else
13140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13150   End Select
13160   Resume EXITP

End Sub

Private Sub country_active_AfterUpdate()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "country_active_AfterUpdate"

13210   With Me
13220     Select Case .country_active
          Case True
            ' ** Leave curr_active along.
13230     Case False
13240       If .curr_active = True Then
13250         .curr_active = False
13260       End If
13270     End Select
13280     strControlName = THIS_PROC
13290     cmdSave_Click  ' ** Procedure: Above.
13300   End With

EXITP:
13310   Exit Sub

ERRH:
13320   Select Case ERR.Number
        Case Else
13330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13340   End Select
13350   Resume EXITP

End Sub

Private Sub country_active_KeyDown(KeyCode As Integer, Shift As Integer)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "country_active_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13420   intShiftDown = (Shift And acShiftMask) > 0
13430   intAltDown = (Shift And acAltMask) > 0
13440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13470       With Me
13480         intRetVal = 0
13490         lngRecsCur = RecCnt  ' ** Function: Below
13500         If .CurrentRecord < lngRecsCur Then
13510           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
13520           .curr_code.SetFocus
13530         Else
13540           DoCmd.SelectObject acForm, .Parent.Name, False
13550           .Parent.opgActive.SetFocus
13560         End If
13570       End With
13580     End Select
13590   End If

        ' ** Shift keys.
13600   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13610     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13620       With Me
13630         intRetVal = 0
13640         .country_iso.SetFocus
13650       End With
13660     End Select
13670   End If

EXITP:
13680   KeyCode = intRetVal
13690   Exit Sub

ERRH:
13700   intRetVal = 0
13710   Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
13720   Case Else
13730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13740   End Select
13750   Resume EXITP

End Sub

Private Sub curr_fund_lbl_cmd_DblClick(Cancel As Integer)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_fund_lbl_cmd_DblClick"

13810   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
13820   Exit Sub

ERRH:
13830   Select Case ERR.Number
        Case Else
13840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13850   End Select
13860   Resume EXITP

End Sub

Private Sub curr_bmu_lbl_cmd_DblClick(Cancel As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_bmu_lbl_cmd_DblClick"

13910   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
13920   Exit Sub

ERRH:
13930   Select Case ERR.Number
        Case Else
13940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13950   End Select
13960   Resume EXITP

End Sub

Private Sub curr_metal_lbl_cmd_DblClick(Cancel As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_metal_lbl_cmd_DblClick"

14010   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
14020   Exit Sub

ERRH:
14030   Select Case ERR.Number
        Case Else
14040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14050   End Select
14060   Resume EXITP

End Sub

Private Sub curr_alt_lbl_cmd_DblClick(Cancel As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_alt_lbl_cmd_DblClick"

14110   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
14120   Exit Sub

ERRH:
14130   Select Case ERR.Number
        Case Else
14140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14150   End Select
14160   Resume EXITP

End Sub

Private Sub curr_unit_lbl_cmd_DblClick(Cancel As Integer)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_unit_lbl_cmd_DblClick"

14210   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
14220   Exit Sub

ERRH:
14230   Select Case ERR.Number
        Case Else
14240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14250   End Select
14260   Resume EXITP

End Sub

Public Function RecCnt() As Long

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

14310   lngRetVal = 0&

14320   With Me
14330     Set rst = .RecordsetClone
14340     With rst
14350       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
14360       Else
14370         .MoveLast
14380         lngRetVal = .RecordCount
14390       End If
14400       .Close
14410     End With
14420   End With

EXITP:
14430   Set rst = Nothing
14440   RecCnt = lngRetVal
14450   Exit Function

ERRH:
14460   lngRetVal = 0&
14470   Select Case ERR.Number
        Case Else
14480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14490   End Select
14500   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

14610   Select Case intWhere
        Case acCmdRecordsGoToFirst
14620     DoCmd.RunCommand acCmdRecordsGoToFirst
14630   Case acCmdRecordsGoToLast
14640     DoCmd.RunCommand acCmdRecordsGoToLast
14650   Case acCmdRecordsGoToPrevious
14660     DoCmd.RunCommand acCmdRecordsGoToPrevious
14670   Case acCmdRecordsGoToNext
14680     DoCmd.RunCommand acCmdRecordsGoToNext
14690   Case Else
14700     If IsMissing(varID) = False Then
14710       With Me
14720         Set rst = .RecordsetClone
14730         With rst
14740           .FindFirst "[curr_id] = '" & varID & "'"
14750           If .NoMatch = False Then
14760             Me.Bookmark = .Bookmark
14770           End If
14780           .Close
14790         End With
14800       End With
14810     End If
14820   End Select

EXITP:
14830   Set rst = Nothing
14840   Exit Sub

ERRH:
14850   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
14860   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
14870   Case Else
14880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14890   End Select
14900   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim blnDetail As Boolean
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc1 As String = ", [curr_word1], [country_name_sort]"
        Const strStdAsc2 As String = ", [country_name_sort], [curr_word1]"

15010   With Me
15020     blnDetail = False
15030     .Controls(strSortLbl).Visible = False
15040     .Controls(strSortLbl2).Visible = False
15050     .Controls(strSortLine).Visible = False
15060     .Controls(strSortLine2).Visible = False
15070     .Controls(strSortLine3).Visible = False
15080     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
15090     .Controls(strSortLine2).Width = lngTpp
15100     .Controls(strSortLine3).Width = lngTpp
15110     If strProc = "Form_Load" Then
15120       strCalled = "curr_word1"
15130       strSortNow = strSortOrig
15140       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
15150       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
15160       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
15170       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
15180       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
15190       .Controls(strSortLbl).Top = lngSortLbl_Top
15200       .Controls(strSortLbl).Left = lngSortLbl_Left
15210       .Controls(strSortLbl).Caption = strArwUp
15220       .Controls(strSortLbl).ForeColor = CLR_DKBLU
15230       .Controls(strSortLine).Top = lngSortLine_Top
15240       .Controls(strSortLine).Left = lngSortLine_Left
15250       .Controls(strSortLine).Width = lngSortLine_Width
15260     Else
15270       If InStr(strProc, "_cmd") > 0 Then
15280         blnDetail = True
15290         strCalled = Left(strProc, (Len(strProc) - Len("_lbl_cmd_DblClick")))
15300       Else
15310         strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
15320       End If
15330       Select Case blnDetail
            Case True
15340         Select Case strCalled
              Case "curr_fund", "curr_metal", "curr_bmu", "curr_alt", "curr_unit"
15350           lngSortLine_Width = (.Controls(strCalled & "_box").Width - (3& * lngTpp))
15360           lngSortLbl_Top = (.Controls(strCalled & "_lbl_off").Top - (3& * lngTpp))
15370           lngSortLbl_Left = (.Controls(strCalled & "_lbl_off").Left - (5& * lngTpp))
15380           lngSortLine_Top = (.Controls(strCalled & "_lbl_off").Top - (3& * lngTpp))
15390           lngSortLine_Left = .Controls(strCalled & "_lbl_off").Left
15400         Case Else
15410           lngSortLine_Width = .Controls(strCalled & "_lbl").Width
15420           lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - (3& * lngTpp))
15430           lngSortLbl_Left = (.Controls(strCalled & "_lbl").Left - (5& * lngTpp))
15440           lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - (3& * lngTpp))
15450           lngSortLine_Left = .Controls(strCalled & "_lbl").Left
15460         End Select
15470         .Controls(strSortLbl2).Top = lngSortLbl_Top
15480         .Controls(strSortLbl2).Left = lngSortLbl_Left
15490         .Controls(strSortLine2).Top = lngSortLine_Top
15500         .Controls(strSortLine3).Top = (lngSortLine_Top + lngTpp)
15510         .Controls(strSortLine2).Left = lngSortLine_Left
15520         .Controls(strSortLine3).Left = lngSortLine_Left
15530         .Controls(strSortLine2).Width = lngSortLine_Width
15540         .Controls(strSortLine3).Width = lngSortLine_Width
15550       Case False
15560         lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
15570         lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
15580         lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
15590         lngSortLine_Left = .Controls(strCalled & "_lbl").Left
15600         lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
15610         .Controls(strSortLbl).Top = lngSortLbl_Top
15620         .Controls(strSortLbl).Left = lngSortLbl_Left
15630         .Controls(strSortLine).Top = lngSortLine_Top
15640         .Controls(strSortLine).Left = lngSortLine_Left
15650         .Controls(strSortLine).Width = lngSortLine_Width
15660       End Select
15670       If strCalled = "curr_word1" Then
15680         strSortAsc = strSortOrig
15690         If strSortNow = strSortAsc Then
15700           strSortNow = "[curr_word1] DESC, [country_name_sort] DESC"
15710           .Controls(strSortLbl).Caption = strArwDn
15720           .Controls(strSortLbl).ForeColor = CLR_DKRED
15730         Else
15740           strSortNow = strSortAsc
15750           .Controls(strSortLbl).Caption = strArwUp
15760           .Controls(strSortLbl).ForeColor = CLR_DKBLU
15770         End If
15780       Else
15790         .Controls(strSortLbl).Caption = strArwUp
15800         .Controls(strSortLbl2).Caption = strArwUp
15810         .Controls(strSortLbl).ForeColor = CLR_DKBLU
15820         .Controls(strSortLbl2).ForeColor = CLR_DKBLU
15830         Select Case strCalled
              Case "curr_code"
15840           strSortAsc = "[curr_code]"
15850           lngSortLbl_Left = (lngSortLbl_Left + (9& * lngTpp))
15860           .Controls(strSortLbl).Left = lngSortLbl_Left
15870         Case "curr_name1"
15880           strSortAsc = "[curr_name_sort]" & strStdAsc1
15890         Case "curr_iso"
15900           strSortAsc = "[curr_iso]"
15910           lngSortLbl_Left = (lngSortLbl_Left + (9& * lngTpp))
15920           .Controls(strSortLbl).Left = lngSortLbl_Left
15930         Case "curr_rate1"
15940           strSortAsc = "[curr_rate1]" & strStdAsc2
15950         Case "curr_rate2"
15960           strSortAsc = "[curr_rate2]" & strStdAsc2
15970         Case "curr_date"
15980           strSortAsc = "[curr_date]" & strStdAsc2
15990           lngSortLbl_Left = (lngSortLbl_Left + (1& * lngTpp))
16000           .Controls(strSortLbl).Left = lngSortLbl_Left
16010         Case "curr_active"
16020           strSortAsc = "[curr_active]" & strStdAsc1
16030           lngSortLbl_Left = (lngSortLbl_Left + (9& * lngTpp))
16040           .Controls(strSortLbl).Left = lngSortLbl_Left
16050         Case "curr_decimal"
16060           strSortAsc = "[curr_decimal]" & strStdAsc1
16070         Case "curr_notes"
16080           strSortAsc = "[curr_notes]" & strStdAsc1
16090         Case "country_code3"
16100           strSortAsc = "[country_code3]" & strStdAsc1  ' ** Some countries list more than one currency.
16110           lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
16120           .Controls(strSortLbl).Left = lngSortLbl_Left
16130         Case "country_name"
16140           strSortAsc = "[country_name_sort]" & strStdAsc1
16150         Case "country_iso"
16160           strSortAsc = "[country_iso]" & strStdAsc2
16170           lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
16180           .Controls(strSortLbl).Left = lngSortLbl_Left
16190         Case "country_active"
16200           strSortAsc = "[country_active]" & strStdAsc2
16210           lngSortLbl_Left = (lngSortLbl_Left + (5& * lngTpp))
16220           .Controls(strSortLbl).Left = lngSortLbl_Left
16230         Case "curr_fund"
16240           strSortAsc = "[curr_fund]" & strStdAsc1
16250         Case "curr_bmu"
16260           strSortAsc = "[curr_bmu]" & strStdAsc1
16270         Case "curr_metal"
16280           strSortAsc = "[curr_metal]" & strStdAsc1
16290         Case "curr_alt"
16300           strSortAsc = "[curr_alt]" & strStdAsc1
16310         Case "curr_unit"
16320           strSortAsc = "[curr_unit]" & strStdAsc1
16330         End Select
16340         If strSortNow = strSortAsc Then
16350           intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
16360           Select Case intCnt
                Case 1
16370             strTmp01 = strSortAsc & " DESC"
16380           Case 2
16390             intPos01 = InStr(strSortAsc, ",")
16400             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
16410           Case 3
16420             intPos01 = InStr(strSortAsc, ",")
16430             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
16440             strTmp02 = Mid(strSortAsc, intPos01)
16450             intPos01 = InStr(2, strTmp02, ",")
16460             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16470             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
16480           Case 4
16490             intPos01 = InStr(strSortAsc, ",")
16500             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
16510             strTmp02 = Mid(strSortAsc, intPos01)
16520             intPos01 = InStr(2, strTmp02, ",")
16530             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16540             strTmp02 = Mid(strTmp02, intPos01)
16550             intPos01 = InStr(2, strTmp02, ",")
16560             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
16570             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
16580           End Select
16590           strSortNow = strTmp01
16600           .Controls(strSortLbl).Caption = strArwDn
16610           .Controls(strSortLbl2).Caption = strArwDn
16620           .Controls(strSortLbl).ForeColor = CLR_DKRED
16630           .Controls(strSortLbl2).ForeColor = CLR_DKRED
16640         Else
16650           strSortNow = strSortAsc
16660         End If
16670       End If
16680     End If
16690     Select Case blnDetail
          Case True
16700       .Controls(strSortLbl2).Visible = True
16710       .Controls(strSortLine2).Visible = True
16720       .Controls(strSortLine3).Visible = True
16730     Case False
16740       .Controls(strSortLbl).Visible = True
16750       .Controls(strSortLine).Visible = True
16760     End Select
16770     .OrderBy = strSortNow
16780     .OrderByOn = True
16790   End With

EXITP:
16800   Exit Sub

ERRH:
16810   Select Case ERR.Number
        Case Else
16820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16830   End Select
16840   Resume EXITP

End Sub

Public Function SortNow_Get() As String

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

16910   strRetVal = strSortNow

EXITP:
16920   SortNow_Get = strRetVal
16930   Exit Function

ERRH:
16940   strRetVal = vbNullString
16950   Select Case ERR.Number
        Case Else
16960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16970   End Select
16980   Resume EXITP

End Function
