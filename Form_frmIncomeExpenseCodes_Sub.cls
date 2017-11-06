VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmIncomeExpenseCodes_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmIncomeExpenseCodes_Sub"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd on frmIncomeExpenseCodes}
' **   Close:            C {cmdClose on frmIncomeExpenseCodes}
' **   Edit:             E {cmdEdit on frmIncomeExpenseCodes}
' **   Print Report:     P {cmdPrintReport on frmIncomeExpenseCodes}
' **   Renumber:         R {cmdRenumber on frmIncomeExpenseCodes}
' **   Update:           U {cmdUpdate on frmIncomeExpenseCodes}
' **   Exit:             X {cmdClose on frmIncomeExpenseCodes}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Tax Codes:        T {ShowTaxCodes}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[revcode_TYPE], [revcode_SORTORDER]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long

Private blnEdited As Boolean
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmIncomeExpenseCodes" Then

160         RevCode_SetEdited THIS_NAME, False  ' ** Module Procedure: modRevCodeFuncs.

170         CLR_DISABLED_FG = CLR_DKGRY
180         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         lngSortLbl_Width = .Sort_lbl.Width

            ' ** Populate tmpRevCodeEdit table.
210         blnRetVal = RevCode_Setup(THIS_NAME)  ' ** Module Function: modRevCodeFuncs.
220         If blnRetVal = False Then
230           .Parent.ClosingForm True  ' ** Form Procedure: frmIncomeExpenseCodes.
240         End If

250       Else
260         Cancel = -1
270       End If
280     End With

290     If Cancel = -1 Then
300       Beep
310       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
320     ElseIf blnRetVal = False Then
330       Cancel = -1
340     End If

EXITP:
350     Exit Sub

ERRH:
360     Select Case ERR.Number
        Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
380     End Select
390     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me

420       .RecordSource = "qryRevCodes_02"
430       .Requery

          ' ** Controls with special 'disabled' BackColor:
440       .opgIncomeExpense_optIncome_box.ForeColor = CLR_IEI_BLU_DIS
450       .opgIncomeExpense_box.BackStyle = acBackStyleTransparent
460       .Legend_Income_box.BackColor = CLR_IEI_BLU_DIS
470       .opgIncomeExpense_optExpense_box.ForeColor = CLR_IEE_RED_DIS
480       .Legend_Expense_box.BackColor = CLR_IEE_RED_DIS
490       .opgIncomeExpense_def_box.ForeColor = .Detail.BackColor
500       .opgIncomeExpense_norm_box.ForeColor = .Detail.BackColor
510       .revcode_DESC.ForeColor = CLR_DISABLED_FG
520       .revcode_DESC.BackColor = CLR_DISABLED_BG
530       .revcode_DESC_def_box.ForeColor = CLR_DISABLED_BG
540       .revcode_DESC_norm_box.ForeColor = CLR_DISABLED_BG
550       .revcode_SORTORDER.ForeColor = CLR_DISABLED_FG
560       .revcode_SORTORDER.BackColor = CLR_DISABLED_BG
570       .revcode_SORTORDER_def_box.ForeColor = CLR_DISABLED_BG
580       .revcode_SORTORDER_norm_box.ForeColor = CLR_DISABLED_BG
590       .IsInc_Def.ForeColor = CLR_DISABLED_BG
600       .Rev_SortOrd_Inc_Def.ForeColor = CLR_DISABLED_FG
610       .IsInc_Norm.ForeColor = CLR_DISABLED_BG
620       .Rev_SortOrd_Inc_Norm.ForeColor = CLR_DISABLED_FG
630       .IsExp_Def.ForeColor = CLR_DISABLED_BG
640       .Rev_SortOrd_Exp_Def.ForeColor = CLR_DISABLED_FG
650       .IsExp_Norm.ForeColor = CLR_DISABLED_BG
660       .Rev_SortOrd_Exp_Norm.ForeColor = CLR_DISABLED_FG
670       .revcode_active_locked1.ForeColor = CLR_DISABLED_BG
680       .revcode_active_locked1.ForeColor = CLR_DISABLED_BG
690       .revcode_active_unlocked1.ForeColor = CLR_DISABLED_BG
700       .revcode_active_unlocked1.ForeColor = CLR_DISABLED_BG
710       .taxcode_type.ForeColor = CLR_DISABLED_FG
720       .taxcode_type.BackColor = CLR_DISABLED_BG
730       Select Case gblnLinkRevTaxCodes
          Case True
740         .taxcode_type.Visible = True
750         .TaxCode_Type_lbl.Visible = True
760         .TaxCode_Type_lbl2.Visible = True
770         .TaxCode_Type_lbl_line.Visible = True
780       Case False
790         .taxcode_type.Visible = False
800         .TaxCode_Type_lbl.Visible = False
810         .TaxCode_Type_lbl2.Visible = False
820         .TaxCode_Type_lbl_line.Visible = False
830       End Select
840       .IsDef_bg1.ForeColor = .Detail.BackColor
850       .IsDef_bg2.ForeColor = .Detail.BackColor

860       DoEvents

870     End With

880     SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
890     Exit Sub

ERRH:
900     Select Case ERR.Number
        Case Else
910       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
920     End Select
930     Resume EXITP

End Sub

Private Sub Form_Current()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim blnTmp01 As Boolean

1010    With Me
1020      If gblnClosing = False And Len(.RecordSource) > 0 Then
1030        .Parent.DefaultCodes_lbl.Visible = False
1040        .Parent.DefaultCodes_lbl.ForeColor = .Parent.ShortcutMenu_lbl.ForeColor
1050        Select Case .Parent.chkEditable
            Case True
1060          If .revcode_ID <= 6& Then
1070            .revcode_SORTORDER.BackColor = CLR_DISABLED_BG
1080          Else
1090            .revcode_SORTORDER.BackColor = CLR_WHT
1100          End If
1110          DoEvents
1120          If IsNull(.revcode_ID) Then
                ' ** New record.
1130            LockDetail False  ' ** Procedure: Below.
1140          Else
                ' ** Can't edit record #1 or #2 (Unspecified), which both have a sort order of 1, or
                ' ** the two new OC records, which both should have a sort order of 2.
1150            blnTmp01 = False
1160            If .revcode_SORTORDER < 3 Then
1170              blnTmp01 = True
1180            ElseIf .revcode_DESC = "Ordinary Dividend" Or .revcode_DESC = "Interest Income" Then
1190              blnTmp01 = True
1200            End If
1210            LockDetail blnTmp01  ' ** Procedure: Below.
1220          End If
1230        Case False
1240          .opgIncomeExpense.Locked = True
1250          .revcode_DESC.Locked = True
1260          .revcode_SORTORDER.Locked = True
1270          .revcode_SORTORDER.BackColor = CLR_DISABLED_BG
1280          .revcode_ACTIVE.Locked = True
1290        End Select
1300        If .revcode_SORTORDER < 3 Or .revcode_DESC = "Ordinary Dividend" Or .revcode_DESC = "Interest Income" Then
1310  On Error Resume Next
1320          If ERR.Number <> 0 Then  'HOW WOULD THERE BE AN ERROR IF WE DON'T DO ANYTHING?
1330  On Error GoTo ERRH
1340            DoCmd.SelectObject acForm, .Parent.Name, False
1350            .Parent.cmdAdd.SetFocus
1360            DoEvents
1370          Else
1380  On Error GoTo ERRH
1390          End If
1400        End If
1410      End If
1420    End With

EXITP:
1430    Exit Sub

ERRH:
1440    Select Case ERR.Number
        Case 2427  ' ** You entered an expression that has no value.
          ' ** RecordSource not yet present.
1450    Case Else
1460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1470    End Select
1480    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1520    intShiftDown = (Shift And acShiftMask) > 0
1530    intAltDown = (Shift And acAltMask) > 0
1540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Shortcut Ctrl-Shift keys  to other forms and functionality:
        ' **   Tax Codes:        T {ShowTaxCodes}

        ' ** Plain keys.
1550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1560      Select Case intRetVal
          Case vbKeyUp
1570        intRetVal = 0
1580        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1590      Case vbKeyDown
1600        intRetVal = 0
1610        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1620      Case vbKeyEscape
1630        With Me
1640          intRetVal = 0
1650          .Parent.cmdClose_Click  ' ** Form Procedure: frmIncomeExpenseCodes.
1660        End With
1670      End Select
1680    End If

        ' ** Alt keys.
1690    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1700      Select Case intRetVal
          Case vbKeyX
1710        With Me
1720          intRetVal = 0
1730          .Parent.cmdClose_Click  ' ** Form Procedure: frmIncomeExpenseCodes.
1740        End With
1750      End Select
1760    End If

        ' ** Ctrl keys.
1770    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1780      Select Case intRetVal
          Case vbKeyS
1790        intRetVal = 0
1800        cmdSave_Click  ' ** Procedure: Below.
1810      Case vbKeyTab, vbKeyReturn
1820        With Me
1830          intRetVal = 0
1840          DoCmd.SelectObject acForm, .Parent.Name, False
1850          If .Parent.cmdPrintReport.Enabled = True Then
1860            .Parent.cmdPrintReport.SetFocus
1870          ElseIf .Parent.cmdRenumber.Enabled = True Then
1880            .Parent.cmdRenumber.SetFocus
1890          ElseIf .Parent.cmdAdd.Enabled = True Then
1900            .Parent.cmdAdd.SetFocus
1910          ElseIf .Parent.cmdEdit.Enabled = True Then
1920            .Parent.cmdEdit.SetFocus
1930          ElseIf .Parent.cmdUpdate.Enabled = True Then
1940            .Parent.cmdUpdate.SetFocus
1950          Else
1960            .Parent.cmdClose.SetFocus
1970          End If
1980        End With
1990      Case vbKeyPageUp, vbKeyHome
2000        intRetVal = 0
2010        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2020      Case vbKeyPageDown, vbKeyEnd
2030        intRetVal = 0
2040        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2050      End Select
2060    End If

        ' ** Ctrl-Shift keys.
2070    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2080      Select Case intRetVal
          Case vbKeyF
2090        With Me
2100          intRetVal = 0
2110          DoCmd.SelectObject acForm, .Parent.Name, False
2120          .Parent.FocusHolder.SetFocus
2130        End With
2140      Case vbKeyT
2150        intRetVal = 0
2160        ShowTaxCodes  ' ** Procedure: Below.
2170      Case vbKeyTab, vbKeyReturn
2180        With Me
2190          intRetVal = 0
2200          DoCmd.SelectObject acForm, .Parent.Name, False
2210          .Parent.cmdClose.SetFocus
2220        End With
2230      End Select
2240    End If

EXITP:
2250    KeyCode = intRetVal
2260    Exit Sub

ERRH:
2270    intRetVal = 0
2280    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2290    Case Else
2300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2310    End Select
2320    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

2410    With Me
2420      If IsNull(.revcode_ID) = True Then
            ' ** New record, and AutoNumber is only on the original m_REVCODE table, not this temporary one.
2430        Set dbs = CurrentDb
            ' ** m_REVCODE, with Max(revcode_ID).
2440        Set qdf = dbs.QueryDefs("qryRevCodes_07")
2450        Set rst = qdf.OpenRecordset
2460        rst.MoveFirst
2470        .revcode_ID = rst![revcode_ID] + 1&
2480        .revcode_CHANGED = True
2490        Select Case .opgIncomeExpense
            Case .opgIncomeExpense_optIncome.OptionValue
2500          .revcode_TYPE_Description = "Income"
2510        Case .opgIncomeExpense_optExpense.OptionValue
2520          .revcode_TYPE_Description = "Expense"
2530        End Select
2540        rst.Close
2550        dbs.Close
2560        blnEdited = True
2570        With .Parent
2580          .EditingForm blnEdited  ' ** Form Procedure: frmIncomeExpenseCodes.
2590          If .cmdUpdate.Enabled = False Then
2600            .cmdUpdate.Enabled = True
2610          End If
2620        End With
2630      End If
2640    End With

EXITP:
2650    Set rst = Nothing
2660    Set qdf = Nothing
2670    Set dbs = Nothing
2680    Exit Sub

ERRH:
2690    Select Case ERR.Number
        Case Else
2700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2710    End Select
2720    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2810    If DataErr <> 0 Then
2820      Select Case DataErr
          Case 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237, 2427
            ' ** 2501: The '|' action was Canceled.
2830        Response = acDataErrContinue
2840      Case 2046  ' ** The command or action isn't available now (first or last record).
2850        MsgBox "There are no records to delete.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
2860        Response = acDataErrContinue
2870      Case 3162
2880        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
2890        Response = acDataErrContinue
2900      Case 3058
2910        MsgBox "You must enter the Description to continue.", vbInformation + vbOKOnly, "Entry Required"
2920        Me.revcode_DESC.SetFocus
2930        Response = acDataErrContinue
2940      Case Else
2950        zErrorHandler THIS_NAME, "Undefined Form Error", DataErr  ' ** Module Function: modErrorHandler.
2960      End Select
2970    End If

EXITP:
2980    Exit Sub

ERRH:
2990    Select Case ERR.Number
        Case Else
3000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3010    End Select
3020    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3110    With Me
3120      If IsNull(.revcode_ID) = False Then
3130        DoCmd.RunCommand acCmdSaveRecord
3140      End If
3150    End With

EXITP:
3160    Exit Sub

ERRH:
3170    Select Case ERR.Number
        Case Else
3180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3190    End Select
3200    Resume EXITP

End Sub

Private Sub opgIncomeExpense_lbl_DblClick(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_lbl_DblClick"

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

Private Sub opgIncomeExpense_lbl2_DblClick(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_lbl2_DblClick"

3410    opgIncomeExpense_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub opgIncomeExpense_lbl3_DblClick(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_lbl3_DblClick"

3510    opgIncomeExpense_lbl_DblClick Cancel  ' ** Procedure: Above.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub opgIncomeExpense_Click()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_Click"

3610    With Me
3620      If .Parent.chkEditable = True Then
3630        If .revcode_SORTORDER < 3 Or .revcode_DESC = "Ordinary Dividend" Or .revcode_DESC = "Interest Income" Then
3640          Select Case .Parent.DefaultCodes_lbl.Visible
              Case True
3650            .Parent.DefaultCodes_lbl.ForeColor = 255&
3660          Case False
3670            .Parent.DefaultCodes_lbl.Visible = True
3680          End Select
3690        End If
3700      End If
3710    End With

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub opgIncomeExpense_AfterUpdate()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_AfterUpdate"

3810    With Me
3820      blnEdited = True
3830      With .Parent
3840        .EditingForm blnEdited  ' ** Form Procedure: frmIncomeExpenseCodes.
3850        RevCode_SetEdited THIS_NAME, True  ' ** Module Procedure: modRevCodeFuncs.
3860        If .cmdUpdate.Enabled = False Then
3870          .cmdUpdate.Enabled = True
3880        End If
3890      End With
3900    End With

EXITP:
3910    Exit Sub

ERRH:
3920    Select Case ERR.Number
        Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Private Sub opgIncomeExpense_optIncome_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_optIncome_KeyDown"

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
4090          .revcode_DESC.SetFocus
4100        End With
4110      End Select
4120    End If

        ' ** Shift keys.
4130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4150        With Me
4160          intRetVal = 0
4170          If .CurrentRecord > 1 Then
4180            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4190            If .taxcode_type.Visible = True And .taxcode_type.Enabled = True Then
4200              .taxcode_type.SetFocus
4210            Else
4220              .revcode_ACTIVE.SetFocus
4230            End If
4240          Else
4250            DoCmd.SelectObject acForm, .Parent.Name, False
4260            .Parent.cmdClose.SetFocus
4270          End If
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

Private Sub opgIncomeExpense_optExpense_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_optExpense_KeyDown"

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
4490          .revcode_DESC.SetFocus
4500        End With
4510      End Select
4520    End If

        ' ** Shift keys.
4530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4550        With Me
4560          intRetVal = 0
4570          If .CurrentRecord > 1 Then
4580            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4590            If .taxcode_type.Visible = True And .taxcode_type.Enabled = True Then
4600              .taxcode_type.SetFocus
4610            Else
4620              .revcode_ACTIVE.SetFocus
4630            End If
4640          Else
4650            DoCmd.SelectObject acForm, .Parent.Name, False
4660            .Parent.cmdClose.SetFocus
4670          End If
4680        End With
4690      End Select
4700    End If

EXITP:
4710    KeyCode = intRetVal
4720    Exit Sub

ERRH:
4730    intRetVal = 0
4740    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4750    Case Else
4760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4770    End Select
4780    Resume EXITP

End Sub

Private Sub revcode_DESC_lbl_DblClick(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_lbl_DblClick"

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

Private Sub revcode_DESC_GotFocus()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_GotFocus"

4910    With Me
4920      If .Parent.chkEditable = True Then
4930        Select Case .IsDef
            Case True
4940          .revcode_DESC.BackColor = CLR_DISABLED_BG
4950        Case False
4960          .revcode_DESC.BackColor = CLR_WHT
4970        End Select
4980      End If
4990      .revcode_DESC.SelLength = 0
5000      .revcode_DESC.SelStart = 99
5010    End With

EXITP:
5020    Exit Sub

ERRH:
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub revcode_DESC_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

5150    With Me

5160      If .Parent.chkEditable = True Then
5170        If .revcode_SORTORDER < 3 Or .revcode_DESC = "Ordinary Dividend" Or .revcode_DESC = "Interest Income" Then

              ' ** Plain keys.
5180          If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5190            If (intRetVal >= vbKey0 And intRetVal <= vbKey9) Or (intRetVal >= vbKeyA And intRetVal <= vbKeyZ) Then
5200              If .Parent.DefaultCodes_lbl.Visible = True Then
5210                .Parent.DefaultCodes_lbl.ForeColor = 255&
5220              Else
5230                .Parent.DefaultCodes_lbl.Visible = True
5240              End If
5250            End If
5260          End If

              ' ** Shift keys.
5270          If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5280            If (intRetVal >= vbKey0 And intRetVal <= vbKey9) Or (intRetVal >= vbKeyA And intRetVal <= vbKeyZ) Then
5290              If .Parent.DefaultCodes_lbl.Visible = True Then
5300                .Parent.DefaultCodes_lbl.ForeColor = 255&
5310              Else
5320                .Parent.DefaultCodes_lbl.Visible = True
5330              End If
5340            End If
5350          End If

5360        End If
5370      End If

5380    End With

EXITP:
5390    KeyCode = intRetVal
5400    Exit Sub

ERRH:
5410    intRetVal = 0
5420    Select Case ERR.Number
        Case Else
5430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5440    End Select
5450    Resume EXITP

End Sub

Private Sub revcode_DESC_AfterUpdate()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_DESC_AfterUpdate"

5510    With Me
5520      blnEdited = True
5530      With .Parent
5540        .EditingForm blnEdited  ' ** Form Procedure: frmIncomeExpenseCodes.
5550        RevCode_SetEdited THIS_NAME, True  ' ** Module Procedure: modRevCodeFuncs.
5560        If .cmdUpdate.Enabled = False Then
5570          .cmdUpdate.Enabled = True
5580        End If
5590      End With
5600      cmdSave_Click  ' ** Procedure: Above.
5610    End With

EXITP:
5620    Exit Sub

ERRH:
5630    Select Case ERR.Number
        Case Else
5640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5650    End Select
5660    Resume EXITP

End Sub

Private Sub Rev_SortOrd_Inc_Def_GotFocus()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "Rev_SortOrd_Inc_Def_GotFocus"

5710    With Me
5720      .revcode_SORTORDER.SetFocus
5730    End With

EXITP:
5740    Exit Sub

ERRH:
5750    Select Case ERR.Number
        Case Else
5760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5770    End Select
5780    Resume EXITP

End Sub

Private Sub Rev_SortOrd_Inc_Norm_GotFocus()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Rev_SortOrd_Inc_Norm_GotFocus"

5810    With Me
5820      .revcode_SORTORDER.SetFocus
5830    End With

EXITP:
5840    Exit Sub

ERRH:
5850    Select Case ERR.Number
        Case Else
5860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5870    End Select
5880    Resume EXITP

End Sub

Private Sub Rev_SortOrd_Exp_Def_GotFocus()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "Rev_SortOrd_Exp_Def_GotFocus"

5910    With Me
5920      .revcode_SORTORDER.SetFocus
5930    End With

EXITP:
5940    Exit Sub

ERRH:
5950    Select Case ERR.Number
        Case Else
5960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5970    End Select
5980    Resume EXITP

End Sub

Private Sub Rev_SortOrd_Exp_Norm_GotFocus()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "Rev_SortOrd_Exp_Norm_GotFocus"

6010    With Me
6020      .revcode_SORTORDER.SetFocus
6030    End With

EXITP:
6040    Exit Sub

ERRH:
6050    Select Case ERR.Number
        Case Else
6060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6070    End Select
6080    Resume EXITP

End Sub

Private Sub revcode_SORTORDER_lbl_DblClick(Cancel As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_SORTORDER_lbl_DblClick"

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

Private Sub revcode_SORTORDER_BeforeUpdate(Cancel As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_SORTORDER_BeforeUpdate"

6210    With Me
6220      If gblnClosing = False Then
6230        If IsNull(.revcode_SORTORDER) Then
6240          Cancel = -1
6250        Else
6260          If Len(.revcode_SORTORDER) = 0 Then
6270            Cancel = -1
6280          End If
6290        End If
6300        If Cancel = -1 Then
6310          MsgBox "All Revenue codes must have a Sort Order.", vbExclamation + vbOKOnly, "Entry Required"
6320        Else
6330          If (.revcode_SORTORDER <= 0) Or (.revcode_SORTORDER > 98) Then
6340            Cancel = -1
6350            MsgBox "Codes must have a Sort Order between 3 and 98.", vbExclamation + vbOKOnly, "Invalid Entry"
6360          ElseIf .revcode_SORTORDER = 1 Or .revcode_SORTORDER = 2 Then
6370            Cancel = -1
6380            MsgBox "Sort Orders 1 and 2 are locked, and can't be changed.", vbExclamation + vbOKOnly, "Invalid Entry"
6390          ElseIf (.revcode_SORTORDER = 3 Or .revcode_SORTORDER = 4) And .opgIncomeExpense = REVTYP_INC Then
6400            Cancel = -1
6410            MsgBox "This Sort Order is locked, and can't be changed.", vbExclamation + vbOKOnly, "Invalid Entry"
                ' Or .revcode_DESC = "Ordinary Dividend" Or .revcode_DESC = "Interest Income"
6420          End If
6430        End If
6440      End If
6450    End With

EXITP:
6460    Exit Sub

ERRH:
6470    Select Case ERR.Number
        Case Else
6480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6490    End Select
6500    Resume EXITP

End Sub

Private Sub revcode_SORTORDER_AfterUpdate()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_SORTORDER_AfterUpdate"

        Dim lngX As Long

6610    With Me
6620      blnEdited = True
6630      With .Parent
6640        .EditingForm blnEdited  ' ** Form Procedure: frmIncomeExpenseCodes.
6650        RevCode_SetEdited THIS_NAME, True  ' ** Module Procedure: modRevCodeFuncs.
6660        If .cmdUpdate.Enabled = False Then
6670          .cmdUpdate.Enabled = True
6680        End If
6690      End With
6700      Select Case .revcode_TYPE
          Case REVTYP_INC
6710        .revcode_SORTORDER_I = .revcode_SORTORDER
6720      Case REVTYP_EXP
6730        .revcode_SORTORDER_E = .revcode_SORTORDER
6740      End Select
6750      For lngX = 0& To (glngRevOs - 1&)
6760        If garr_varRevO(REVO_ID, lngX) = .revcode_ID Then
6770          If .revcode_SORTORDER <> garr_varRevO(REVO_SORTORD, lngX) Then
6780            .revcode_CHANGED = True
6790            garr_varRevO(REVO_CHANGED, lngX) = True
6800          Else
6810            .revcode_CHANGED = False
6820            garr_varRevO(REVO_CHANGED, lngX) = False
6830          End If
6840          Exit For
6850        End If
6860      Next
6870      cmdSave_Click  ' ** Procedure: Above.
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

Private Sub revcode_SORTORDER_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_SORTORDER_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7020    intShiftDown = (Shift And acShiftMask) > 0
7030    intAltDown = (Shift And acAltMask) > 0
7040    intCtrlDown = (Shift And acCtrlMask) > 0

7050    With Me

7060      If .Parent.chkEditable = True Then
7070        If .revcode_SORTORDER < 3 Or .revcode_DESC = "Ordinary Dividend" Or .revcode_DESC = "Interest Income" Then

              ' ** Plain keys.
7080          If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7090            If (intRetVal >= vbKey0 And intRetVal <= vbKey9) Or (intRetVal >= vbKeyA And intRetVal <= vbKeyZ) Then
7100              If .Parent.DefaultCodes_lbl.Visible = True Then
7110                .Parent.DefaultCodes_lbl.ForeColor = 255&
7120              Else
7130                .Parent.DefaultCodes_lbl.Visible = True
7140              End If
7150            End If
7160          End If

              ' ** Shift keys.
7170          If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7180            If (intRetVal >= vbKey0 And intRetVal <= vbKey9) Or (intRetVal >= vbKeyA And intRetVal <= vbKeyZ) Then
7190              If .Parent.DefaultCodes_lbl.Visible = True Then
7200                .Parent.DefaultCodes_lbl.ForeColor = 255&
7210              Else
7220                .Parent.DefaultCodes_lbl.Visible = True
7230              End If
7240            End If
7250          End If

7260        End If
7270      End If

7280    End With

EXITP:
7290    KeyCode = intRetVal
7300    Exit Sub

ERRH:
7310    intRetVal = 0
7320    Select Case ERR.Number
        Case Else
7330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7340    End Select
7350    Resume EXITP

End Sub

Private Sub revcode_ACTIVE_lbl_DblClick(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ACTIVE_lbl_DblClick"

7410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
7420    Exit Sub

ERRH:
7430    Select Case ERR.Number
        Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Sub

Private Sub revcode_ACTIVE_Click()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ACTIVE_Click"

7510    With Me
7520      If .Parent.chkEditable = True Then
7530        If .revcode_SORTORDER < 3 Or .revcode_DESC = "Ordinary Dividend" Or .revcode_DESC = "Interest Income" Then
7540          If .Parent.DefaultCodes_lbl.Visible = True Then
7550            .Parent.DefaultCodes_lbl.ForeColor = 255&
7560          Else
7570            .Parent.DefaultCodes_lbl.Visible = True
7580          End If
7590        End If
7600      End If
7610    End With

EXITP:
7620    Exit Sub

ERRH:
7630    Select Case ERR.Number
        Case Else
7640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7650    End Select
7660    Resume EXITP

End Sub

Private Sub revcode_ACTIVE_AfterUpdate()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ACTIVE_AfterUpdate"

7710    With Me
7720      blnEdited = True
7730      With .Parent
7740        .EditingForm blnEdited  ' ** Form Procedure: frmIncomeExpenseCodes.
7750        RevCode_SetEdited THIS_NAME, True  ' ** Module Procedure: modRevCodeFuncs.
7760        If .cmdUpdate.Enabled = False Then
7770          .cmdUpdate.Enabled = True
7780        End If
7790      End With
7800      cmdSave_Click  ' ** Procedure: Above.
7810    End With

EXITP:
7820    Exit Sub

ERRH:
7830    Select Case ERR.Number
        Case Else
7840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7850    End Select
7860    Resume EXITP

End Sub

Private Sub revcode_ACTIVE_KeyDown(KeyCode As Integer, Shift As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_ACTIVE_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7920    intShiftDown = (Shift And acShiftMask) > 0
7930    intAltDown = (Shift And acAltMask) > 0
7940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7970        With Me
7980          intRetVal = 0
7990          lngRecsCur = RecCnt  ' ** Function: Below.
8000          If .taxcode_type.Visible = True And .taxcode_type.Enabled = True Then
8010            .taxcode_type.SetFocus
8020          Else
8030            If .CurrentRecord < lngRecsCur Then
8040              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
8050              .opgIncomeExpense.SetFocus
8060            Else
8070              DoCmd.SelectObject acForm, .Parent.Name, False
8080              If .Parent.cmdRenumber.Enabled = True Then
8090                .Parent.cmdRenumber.SetFocus
8100              ElseIf .Parent.cmdAdd.Enabled = True Then
8110                .Parent.cmdAdd.SetFocus
8120              ElseIf .Parent.cmdEdit.Enabled = True Then
8130                .Parent.cmdEdit.SetFocus
8140              ElseIf .Parent.cmdUpdate.Enabled = True Then
8150                .Parent.cmdUpdate.SetFocus
8160              Else
8170                .Parent.cmdClose.SetFocus
8180              End If
8190            End If
8200          End If
8210        End With
8220      End Select
8230    End If

        ' ** Shift keys.
8240    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8250      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8260        With Me
8270          intRetVal = 0
8280          .revcode_SORTORDER.SetFocus
8290        End With
8300      End Select
8310    End If

EXITP:
8320    KeyCode = intRetVal
8330    Exit Sub

ERRH:
8340    intRetVal = 0
8350    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8360    Case Else
8370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8380    End Select
8390    Resume EXITP

End Sub

Private Sub taxcode_type_lbl_DblClick(Cancel As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "TaxCode_Type_lbl_DblClick"

8410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
8420    Exit Sub

ERRH:
8430    Select Case ERR.Number
        Case Else
8440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8450    End Select
8460    Resume EXITP

End Sub

Private Sub TaxCode_Type_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "TaxCode_Type_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8520    intShiftDown = (Shift And acShiftMask) > 0
8530    intAltDown = (Shift And acAltMask) > 0
8540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8570        With Me
8580          intRetVal = 0
8590          lngRecsCur = RecCnt  ' ** Function: Below.
8600          If .CurrentRecord < lngRecsCur Then
8610            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
8620            .opgIncomeExpense.SetFocus
8630          Else
8640            DoCmd.SelectObject acForm, .Parent.Name, False
8650            If .Parent.cmdPrintReport.Enabled = True Then
8660              .Parent.cmdPrintReport.SetFocus
8670            ElseIf .Parent.cmdRenumber.Enabled = True Then
8680              .Parent.cmdRenumber.SetFocus
8690            ElseIf .Parent.cmdAdd.Enabled = True Then
8700              .Parent.cmdAdd.SetFocus
8710            ElseIf .Parent.cmdEdit.Enabled = True Then
8720              .Parent.cmdEdit.SetFocus
8730            ElseIf .Parent.cmdUpdate.Enabled = True Then
8740              .Parent.cmdUpdate.SetFocus
8750            Else
8760              .Parent.cmdClose.SetFocus
8770            End If
8780          End If
8790        End With
8800      End Select
8810    End If

        ' ** Shift keys.
8820    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8830      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8840        With Me
8850          intRetVal = 0
8860          .revcode_ACTIVE.SetFocus
8870        End With
8880      End Select
8890    End If

EXITP:
8900    KeyCode = intRetVal
8910    Exit Sub

ERRH:
8920    intRetVal = 0
8930    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
8940    Case Else
8950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8960    End Select
8970    Resume EXITP

End Sub

Private Sub LockDetail(blnLock As Boolean)
' ** Lock the detail fields for Unspecified Income/Expense, etc.

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "LockDetail"

9010    With Me
9020      .opgIncomeExpense.Locked = blnLock
9030      .revcode_DESC.Locked = blnLock
9040      .revcode_SORTORDER.Locked = blnLock
9050      .revcode_ACTIVE.Locked = blnLock
9060    End With

EXITP:
9070    Exit Sub

ERRH:
9080    Select Case ERR.Number
        Case Else
9090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9100    End Select
9110    Resume EXITP

End Sub

Public Sub ClosingForm(blnClose As Boolean)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "ClosingForm"

9210    gblnClosing = blnClose

EXITP:
9220    Exit Sub

ERRH:
9230    Select Case ERR.Number
        Case Else
9240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9250    End Select
9260    Resume EXITP

End Sub

Public Function SortOrdChanged() As Boolean

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "SortOrdChanged"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim blnRetVal As Boolean

9310    blnRetVal = False

9320    Set dbs = CurrentDb
9330    With dbs
          ' ** m_REVCODE, linked to tmpRevCodeEdit, just non-matching revcode_SORTORDER.
9340      Set qdf = .QueryDefs("qryRevCodes_19")
9350      Set rst = qdf.OpenRecordset
9360      With rst
9370        If .BOF = True And .EOF = True Then
              ' ** No changes were made.
9380        Else
9390          blnRetVal = True
9400        End If
9410        .Close
9420      End With
9430      .Close
9440    End With

9450    If blnRetVal = False Then
9460      lngTmp01 = DCount("*", "m_REVCODE")
9470      lngTmp02 = DCount("*", "tmpRevCodeEdit")
9480      If lngTmp01 <> lngTmp02 Then
9490        blnRetVal = True
9500      End If
9510    End If

EXITP:
9520    Set rst = Nothing
9530    Set qdf = Nothing
9540    Set dbs = Nothing
9550    SortOrdChanged = blnRetVal
9560    Exit Function

ERRH:
9570    blnRetVal = False
9580    Select Case ERR.Number
        Case Else
9590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9600    End Select
9610    Resume EXITP

End Function

Public Sub ShowTaxCodes()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowTaxCodes"

        Dim blnShow As Boolean

9710    With Me
9720      blnShow = (Not .taxcode_type.Visible)
9730      .FocusHolder.SetFocus
9740      Select Case blnShow
          Case True
9750        .taxcode_type.Visible = True
9760        .TaxCode_Type_lbl.Visible = True
9770        .TaxCode_Type_lbl2.Visible = True
9780        .TaxCode_Type_lbl_line.Visible = True
9790      Case False
9800        .taxcode_type.Visible = False
9810        .TaxCode_Type_lbl.Visible = False
9820        .TaxCode_Type_lbl2.Visible = False
9830        .TaxCode_Type_lbl_line.Visible = False
9840      End Select
9850    End With

EXITP:
9860    Exit Sub

ERRH:
9870    Select Case ERR.Number
        Case Else
9880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9890    End Select
9900    Resume EXITP

End Sub

Public Function RecCnt() As Long

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

10010   lngRetVal = 0&

10020   With Me
10030     Set rst = .RecordsetClone
10040     With rst
10050       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
10060       Else
10070         .MoveLast
10080         lngRetVal = .RecordCount
10090       End If
10100       .Close
10110     End With
10120   End With

EXITP:
10130   Set rst = Nothing
10140   RecCnt = lngRetVal
10150   Exit Function

ERRH:
10160   lngRetVal = 0&
10170   Select Case ERR.Number
        Case Else
10180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10190   End Select
10200   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

10310   Select Case intWhere
        Case acCmdRecordsGoToFirst
10320     DoCmd.RunCommand acCmdRecordsGoToFirst
10330   Case acCmdRecordsGoToLast
10340     DoCmd.RunCommand acCmdRecordsGoToLast
10350   Case acCmdRecordsGoToPrevious
10360     DoCmd.RunCommand acCmdRecordsGoToPrevious
10370   Case acCmdRecordsGoToNext
10380     DoCmd.RunCommand acCmdRecordsGoToNext
10390   Case Else
10400     If IsMissing(varID) = False Then
10410       With Me
10420         Set rst = .RecordsetClone
10430         With rst
10440           .FindFirst "[revcode_ID] = " & CStr(varID)
10450           If .NoMatch = False Then
10460             Me.Bookmark = .Bookmark
10470           End If
10480           .Close
10490         End With
10500       End With
10510     End If
10520   End Select

EXITP:
10530   Set rst = Nothing
10540   Exit Sub

ERRH:
10550   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
10560   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
10570   Case Else
10580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10590   End Select
10600   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [revcode_TYPE], [revcode_SORTORDER]"

10710   With Me
10720     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
10730       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
10740     End If
10750     .Controls(strSortLbl).Visible = False
10760     .Controls(strSortLine).Visible = False
10770     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the the right side of the form.
10780     If strProc = "Form_Load" Then
10790       strCalled = "opgIncomeExpense"
10800       strSortNow = strSortOrig
10810       lngSortLbl_Top = .Controls(strCalled & "_lbl").Top
10820       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
10830       lngSortLine_Top = .Controls(strCalled & "_lbl").Top
10840       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
10850       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
10860       .Controls(strSortLbl).Top = lngSortLbl_Top
10870       .Controls(strSortLbl).Left = lngSortLbl_Left
10880       .Controls(strSortLine).Top = lngSortLine_Top
10890       .Controls(strSortLine).Left = lngSortLine_Left
10900       .Controls(strSortLine).Width = lngSortLine_Width
10910       .Controls(strSortLbl).Caption = strArwUp
10920       .Controls(strSortLbl).ForeColor = CLR_DKBLU
10930     Else
10940       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
10950       lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
10960       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
10970       lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
10980       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
10990       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
11000       .Controls(strSortLbl).Top = lngSortLbl_Top
11010       .Controls(strSortLbl).Left = lngSortLbl_Left
11020       .Controls(strSortLine).Top = lngSortLine_Top
11030       .Controls(strSortLine).Left = lngSortLine_Left
11040       .Controls(strSortLine).Width = lngSortLine_Width
11050       If strCalled = "opgIncomeExpense" Then
11060         strSortAsc = strSortOrig
11070         If strSortNow = strSortAsc Then
11080           strSortNow = "[revcode_TYPE] DESC, [revcode_SORTORDER] DESC"
11090           .Controls(strSortLbl).Caption = strArwDn
11100           .Controls(strSortLbl).ForeColor = CLR_DKRED
11110         Else
11120           strSortNow = strSortAsc
11130           .Controls(strSortLbl).Caption = strArwUp
11140           .Controls(strSortLbl).ForeColor = CLR_DKBLU
11150         End If
11160         lngSortLbl_Top = (lngSortLbl_Top - lngTpp)
11170         lngSortLine_Top = (lngSortLine_Top - lngTpp)
11180         .Controls(strSortLbl).Top = lngSortLbl_Top
11190         .Controls(strSortLine).Top = lngSortLine_Top
11200       Else
11210         .Controls(strSortLbl).Caption = strArwUp
11220         .Controls(strSortLbl).ForeColor = CLR_DKBLU
11230         Select Case strCalled
              Case "revcode_DESC"
11240           strSortAsc = "[revcode_DESC]" & strStdAsc
11250         Case "revcode_SORTORDER"
11260           strSortAsc = "[revcode_SORTORDER], [revcode_TYPE]"
11270         Case "revcode_ACTIVE"
11280           strSortAsc = "[revcode_ACTIVE]" & strStdAsc
11290           lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
11300           .Controls(strSortLbl).Left = lngSortLbl_Left
11310         Case "TaxCode_Type"
11320           strSortAsc = "[TaxCode_Type_Description]" & strStdAsc
11330         End Select
11340         If strSortNow = strSortAsc Then
11350           intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
11360           Select Case intCnt
                Case 1
11370             strTmp01 = strSortAsc & " DESC"
11380           Case 2
11390             intPos01 = InStr(strSortAsc, ",")
11400             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11410             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
11420           Case 3
11430             intPos01 = InStr(strSortAsc, ",")
11440             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11450             strTmp02 = Mid(strSortAsc, intPos01)
11460             intPos01 = InStr(2, strTmp02, ",")
11470             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11480             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
11490           Case 4
11500             intPos01 = InStr(strSortAsc, ",")
11510             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
11520             strTmp02 = Mid(strSortAsc, intPos01)
11530             intPos01 = InStr(2, strTmp02, ",")
11540             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11550             strTmp02 = Mid(strTmp02, intPos01)
11560             intPos01 = InStr(2, strTmp02, ",")
11570             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
11580             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
11590           End Select
11600           strSortNow = strTmp01
11610           .Controls(strSortLbl).Caption = strArwDn
11620           .Controls(strSortLbl).ForeColor = CLR_DKRED
11630         Else
11640           strSortNow = strSortAsc
11650         End If
11660       End If
11670     End If
11680     .Controls(strSortLbl).Visible = True
11690     .Controls(strSortLine).Visible = True
11700     .OrderBy = strSortNow
11710     .OrderByOn = True
11720   End With

        ' ** Potential Sorts:
        ' **   [revcode_TYPE], [revcode_SORTORDER]
        ' **   [revcode_DESC], [revcode_TYPE], [revcode_SORTORDER]
        ' **   [revcode_SORTORDER], [revcode_TYPE]
        ' **   [revcode_ACTIVE], [revcode_TYPE], [revcode_SORTORDER]
        ' **   [TaxCode_Type_Description], [revcode_TYPE], [revcode_SORTORDER]

EXITP:
11730   Exit Sub

ERRH:
11740   Select Case ERR.Number
        Case Else
11750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11760   End Select
11770   Resume EXITP

End Sub

Public Function SortNow_Get() As String

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

11810   strRetVal = strSortNow

EXITP:
11820   SortNow_Get = strRetVal
11830   Exit Function

ERRH:
11840   strRetVal = vbNullString
11850   Select Case ERR.Number
        Case Else
11860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11870   End Select
11880   Resume EXITP

End Function
