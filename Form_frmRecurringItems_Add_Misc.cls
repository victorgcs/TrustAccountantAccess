VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRecurringItems_Add_Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRecurringItems_Add_Misc"

'VGC 09/20/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Description:      D {Recur_Name}
' **   Form 1099:        F {RecurringItemAux_1099}
' **   Other:            H {opgLoc_optOther}
' **   Opt Misc:         M {opgRecurType_optMisc}
' **   OK:               O {cmdOK}
' **   Opt Paid:         P {opgRecurType_optPaid}
' **   Opt Received:     R {opgRecurType_optReceived}
' **   USA:              U {opgLoc_optUSA}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private lngRecurID As Long, blnIsOpen As Boolean
Private blnEnableCountry As Boolean, intOpgLoc As Integer
Private strCallingForm As String, strMiscType As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strCallingForm = Trim(.OpenArgs)
150           intPos01 = InStr(strCallingForm, "~")
160           If intPos01 > 0 Then

170             blnIsOpen = True

180             strMiscType = Mid(strCallingForm, (intPos01 + 1))
190             strCallingForm = Left(strCallingForm, (intPos01 - 1))

200             .RecurringItemAux_1099.Enabled = False
210             .RecurringItemAux_1099.Visible = False
220             .RecurringItemAux_TaxNo.Enabled = False
230             .RecurringItemAux_TaxNo.Visible = False

240             StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

250             CLR_DISABLED_FG = CLR_DKGRY
260             CLR_DISABLED_BG = CLR_LTTEAL

270           Else
280             Cancel = -1
290             strCallingForm = vbNullString
300           End If
310         Else
320           Cancel = -1
330         End If
340       Else
350         Cancel = -1
360       End If
370     End With

380     If Cancel = -1 Then
390       Beep
400       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
410     End If

EXITP:
420     Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me

520       Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

530       opgRecurType_AfterUpdate  ' ** Procedure: Below.

540       If strMiscType = vbNullString Then
550         strMiscType = "Misc."
560         .Recur_Type.Visible = False
570         .opgRecurType.Visible = True
580         .opgRecurType_box.Visible = True
590       Else
600         .Recur_Type.Visible = True
610         .opgRecurType.Visible = False
620         .opgRecurType_box.Visible = False
630       End If

640       Select Case strMiscType
          Case "Misc."
650         .Recur_Type = "Misc"
660         .Caption = "Add New Recurring Item"
670         .Header_lbl_Item.Visible = True
680         .Header_lbl_Payment.Visible = False
690         .Header_lbl_Receipt.Visible = False
700         .Recur_lbl.Caption = "Recurring Item"
710         .Recur_Address.Visible = False
720         .Recur_City.Visible = False
730         .Recur_State.Visible = False
740         .Recur_Zip.Visible = False
750         .Recur_Country.Visible = False
760         .Recur_PostalCode.Visible = False
770         .opgLoc.Visible = False
780         .opgLoc_box.Visible = False
            '.RecurringItemAux_1099.Visible = False
            '.RecurringItemAux_TaxNo.Visible = False
790       Case "Paid"
800         .Recur_Type = "Payee"
810         .Caption = "Add New Recurring Payment"
820         .Header_lbl_Item.Visible = False
830         .Header_lbl_Payment.Visible = True
840         .Header_lbl_Receipt.Visible = False
850         .Recur_lbl.Caption = "Recurring Payment"
860         .Recur_Address.Visible = True
870         .Recur_City.Visible = True
880         .Recur_State.Visible = True
890         .Recur_Zip.Visible = True
900         .Recur_Country.Visible = True
910         .Recur_PostalCode.Visible = True
920         .opgLoc.Visible = True
930         .opgLoc_box.Visible = True
            '.RecurringItemAux_1099.Visible = True
            '.RecurringItemAux_TaxNo.Visible = True
            '.RecurringItemAux_TaxNo.Enabled = False
            '.RecurringItemAux_TaxNo.BorderColor = WIN_CLR_DISR
            '.RecurringItemAux_TaxNo.BackStyle = acBackStyleTransparent
940       Case "Received"
950         .Recur_Type = "Payor"
960         .Caption = "Add New Recurring Receipt"
970         .Header_lbl_Item.Visible = False
980         .Header_lbl_Payment.Visible = False
990         .Header_lbl_Receipt.Visible = True
1000        .Recur_lbl.Caption = "Recurring Receipt"
1010        .Recur_Address.Visible = False
1020        .Recur_City.Visible = False
1030        .Recur_State.Visible = False
1040        .Recur_Zip.Visible = False
1050        .Recur_Country.Visible = False
1060        .Recur_PostalCode.Visible = False
1070        .opgLoc.Visible = False
1080        .opgLoc_box.Visible = False
            '.RecurringItemAux_1099.Visible = False
            '.RecurringItemAux_TaxNo.Visible = False
1090      End Select

1100      If IsNull(.Recur_State_Pref) = False Then
1110        .Recur_State.DefaultValue = Chr(34) & .Recur_State_Pref & Chr(34)
1120      End If

1130      .opgLoc.Enabled = True
1140      opgLoc_AfterUpdate  ' ** Procedure: Below.

1150      lngRecurID = 0&
1160      .Recur_Name.SetFocus

1170      .Recur_Type.ForeColor = CLR_DISABLED_FG
1180      .Recur_Type.BackColor = CLR_DISABLED_BG

1190    End With

EXITP:
1200    Exit Sub

ERRH:
1210    Select Case ERR.Number
        Case -2147352567  ' ** Could be any number of things!
          ' ** Ignore.
1220    Case Else  ' ** Those huge negative error numbers are too big for the Integer in the Form_Error() parameters.
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality.
        ' **   Other:            H {opgLoc_optOther}
        ' **   USA:              U {opgLoc_optUSA}

        ' ** Shortcut Ctrl keys to other forms and functionality.
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyEscape
1370        With Me
1380          intRetVal = 0
1390          cmdCancel_Click  ' ** Procedure: Below.
1400        End With
1410      End Select
1420    End If

        ' ** Alt keys.
1430    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1440      Select Case intRetVal
          Case vbKeyH
1450        With Me
1460          intRetVal = 0
1470          If .opgLoc.Enabled = True Then
1480            .opgLoc.SetFocus
1490            .opgLoc = .opgLoc_optOther.OptionValue
1500            opgLoc_AfterUpdate  ' ** Procedure: Below.
1510          End If
1520        End With
1530      Case vbKeyU
1540        With Me
1550          intRetVal = 0
1560          If .opgLoc.Enabled = True Then
1570            .opgLoc.SetFocus
1580            .opgLoc = .opgLoc_optUSA.OptionValue
1590            opgLoc_AfterUpdate  ' ** Procedure: Below.
1600          End If
1610        End With
1620      End Select
1630    End If

        ' ** Ctl keys.
1640    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1650      Select Case intRetVal
          Case vbKeyS
1660        intRetVal = 0
1670        cmdSave_Click  ' ** Procedure: Below.
1680      End Select
1690    End If

        ' ** Ctl-Shift keys.
1700    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1710      Select Case intRetVal
          Case vbKeyF
1720        With Me
1730          intRetVal = 0
1740          .FocusHolder.SetFocus
1750        End With
1760      End Select
1770    End If

EXITP:
1780    KeyCode = intRetVal
1790    Exit Sub

ERRH:
1800    intRetVal = 0
1810    Select Case ERR.Number
        Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
1820    Case Else
1830      THAT_PROC = THIS_PROC
1840      That_Erl = Erl
1850      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1860    End Select
1870    Resume EXITP

End Sub

Private Sub Form_Timer()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1910    With Me
1920      .TimerInterval = 0&
1930      If gblnSetFocus = True Then
1940        gblnSetFocus = False
1950        strFormName = vbNullString
1960  On Error Resume Next
1970        Do Until strFormName = THIS_NAME
1980          DoCmd.SelectObject acForm, THIS_NAME, False
1990          strFormName = Screen.ActiveForm.Name
2000        Loop
2010  On Error GoTo ERRH
2020      End If
2030    End With

EXITP:
2040    Exit Sub

ERRH:
2050    DoCmd.Hourglass False
2060    Select Case ERR.Number
        Case Else
2070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2080    End Select
2090    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2110    If DataErr <> 0 Then
2120      Select Case DataErr
          Case 2448  ' ** You can't assign a value to this object.
            ' ** Oh yes I can!
2130        Response = acDataErrContinue
2140      Case 3022  ' ** The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.
2150        MsgBox "A Recurring Item by this name already exists for this type.", vbInformation + vbOKOnly, "Duplicate Recurring Item"
2160        Response = acDataErrContinue
2170      Case 3058
2180        Select Case Me.ActiveControl.Name
            Case "Recur_Name"
2190          MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
2200          Me.Recur_Name.SetFocus
2210          Response = acDataErrContinue
2220        End Select
2230      Case Else
2240        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2250        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2260      End Select
2270    End If

EXITP:
2280    Exit Sub

ERRH:
2290    Select Case ERR.Number
        Case Else
2300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2310    End Select
2320    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2410    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2420    Select Case strCallingForm
        Case "frmJournal", "frmJournal_Columns"
2430      gblnSetFocus = True
2440      gstrReturningForm = THIS_NAME
2450      DoCmd.SelectObject acForm, strCallingForm, False
2460      Forms(strCallingForm).TimerInterval = 100&
2470    End Select

EXITP:
2480    Exit Sub

ERRH:
2490    Select Case ERR.Number
        Case Else
2500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2510    End Select
2520    Resume EXITP

End Sub

Private Sub cmdSave_Click()

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

Private Sub cmdOK_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRecurName As String, strRecurType As String

2710    With Me
2720      If IsNull(.Recur_Name) Then
2730        MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
2740        .Recur_Name.SetFocus
2750      Else
2760        If Trim(.Recur_Name) = vbNullString Then
2770          MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
2780          .Recur_Name.SetFocus
2790        Else
2800          .Recur_Name.SetFocus
2810          If InStr(.Recur_Name.text, Chr(34)) > 0 Then
2820            .cmdOK.SetFocus
2830            MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
2840            DoCmd.CancelEvent
2850          Else

2860            strRecurName = .Recur_Name
2870            strRecurType = .Recur_Type

                ' ** If they've flipped around, only take the final choice.
2880            Select Case .opgLoc
                Case .opgLoc_optUSA.OptionValue
2890              If IsNull(.Recur_Country) = False Then
2900                .Recur_Country = Null
2910              End If
2920              If IsNull(.Recur_PostalCode) = False Then
2930                .Recur_PostalCode = Null
2940              End If
2950            Case .opgLoc_optOther.OptionValue
2960              If IsNull(.Recur_State) = False Then
2970                .Recur_State = Null
2980              End If
2990              If IsNull(.Recur_Zip) = False Then
3000                .Recur_Zip = Null
3010              End If
3020            End Select

                ' ** Same sentiments.
3030            If .opgRecurType.Visible = True And .opgRecurType.Enabled = True Then
3040              If .opgRecurType <> .opgRecurType_optPaid.OptionValue Then
3050                .Recur_Address = Null
3060                .Recur_City = Null
3070                .Recur_State = Null
3080                .Recur_Zip = Null
3090                .Recur_Country = Null
3100                .Recur_PostalCode = Null
3110                .RecurringItemAux_1099 = False
3120                .RecurringItemAux_TaxNo = Null
3130              End If
3140            End If

                ' ** If this is saved, it will always come up as already existing!
                'DoCmd.RunCommand acCmdSelectRecord
                'DoCmd.RunCommand acCmdSaveRecord

3150            Set dbs = CurrentDb
                ' ** RecurringItems, by specified [nam], [typ].
3160            Set qdf = dbs.QueryDefs("qryRecurringItems_04")
3170            With qdf.Parameters
3180              ![nam] = strRecurName
3190              ![typ] = strRecurType
3200            End With
3210            Set rst = qdf.OpenRecordset
3220            If rst.BOF = True And rst.EOF = True Then
                  ' ** All's well.
3230              rst.Close
3240              dbs.Close
3250              .cmdOK.SetFocus

3260              Select Case strCallingForm
                  Case "frmJournal"
3270                Forms(strCallingForm).frmJournal_Sub5_Misc.Form.cmbRecurringItems = .Recur_Name
3280                Forms(strCallingForm).frmJournal_Sub5_Misc.Form.cmbRecurringItems.Requery
3290              Case "frmJournal_Columns"
3300                gstrCrtRpt_NetLoss = .Recur_Name  ' ** Borrowing this variable from the Court Reports.
3310              End Select
3320              DoEvents

3330              DoCmd.Close acForm, THIS_NAME

3340              DoCmd.SelectObject acForm, strCallingForm, False
3350              Select Case strCallingForm
                  Case "frmJournal"
3360                Forms(strCallingForm).frmJournal_Sub5_Misc.Form.miscICash.SetFocus
3370              Case "frmJournal_Columns"
                    ' ** Let JC_Frm_RecurringItems() set the focus.
3380              End Select

3390            Else
3400              rst.Close
3410              dbs.Close
3420              MsgBox "A Recurring Item by this name already exists for this type.", vbInformation + vbOKOnly, "Duplicate Recurring Item"
3430              .Recur_Name.SetFocus
3440            End If
3450          End If
3460        End If
3470      End If
3480    End With

EXITP:
3490    Set rst = Nothing
3500    Set qdf = Nothing
3510    Set dbs = Nothing
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3620    intShiftDown = (Shift And acShiftMask) > 0
3630    intAltDown = (Shift And acAltMask) > 0
3640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3660      Select Case intRetVal
          Case vbKeyTab
3670        With Me
3680          intRetVal = 0
3690          .cmdCancel.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab
3750        With Me
3760          intRetVal = 0
3770          If .RecurringItemAux_TaxNo.Visible = True And .RecurringItemAux_TaxNo.Enabled = True Then
3780            .RecurringItemAux_TaxNo.SetFocus
3790          ElseIf .RecurringItemAux_1099.Visible = True And .RecurringItemAux_1099.Enabled = True Then
3800            .RecurringItemAux_1099.SetFocus
3810          Else
3820            If .Recur_Address.Enabled = True And .Recur_Address.Visible = True Then
3830              If .Recur_Country.Enabled = True Then
3840                .Recur_PostalCode.SetFocus
3850              Else
3860                .Recur_Zip.SetFocus
3870              End If
3880            Else
3890              .Recur_Name.SetFocus
3900            End If
3910          End If
3920        End With
3930      End Select
3940    End If

EXITP:
3950    KeyCode = intRetVal
3960    Exit Sub

ERRH:
3970    intRetVal = 0
3980    THAT_PROC = THIS_PROC
3990    That_Erl = Erl
4000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4010    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

4110    With Me
4120      If .NewRecord = True Then
4130  On Error Resume Next  ' ** If it's unavailable, e.g., new record with nothing in it, just continue to close.
4140        DoCmd.RunCommand acCmdUndo
4150  On Error GoTo ERRH
4160      Else
4170        DoCmd.RunCommand acCmdSelectRecord
4180        DoCmd.RunCommand acCmdDeleteRecord
4190      End If
4200      gblnMessage = False

4210      DoCmd.Close acForm, THIS_NAME

4220      DoCmd.SelectObject acForm, strCallingForm, False
4230      Select Case strCallingForm
          Case "frmJournal"
4240        Forms(strCallingForm).frmJournal_Sub5_Misc.SetFocus
4250        Forms(strCallingForm).frmJournal_Sub5_Misc.Form.cmbRecurringItems.SetFocus
4260      Case "frmJournal_Columns"
            ' ** Let JC_Frm_RecurringItems() set the focus.
4270      End Select
4280    End With

EXITP:
4290    Exit Sub

ERRH:
4300    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4310    Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
          ' ** Ignore.
4320    Case 3031, 3101
4330      DoCmd.RunCommand acCmdUndo
4340    Case Else
4350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4360    End Select
4370    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
          Case vbKeyTab
4470        With Me
4480          intRetVal = 0
4490          If .opgRecurType.Visible = True And .opgRecurType.Enabled = True Then
4500            .opgRecurType.SetFocus
4510          Else
4520            .Recur_Name.SetFocus
4530          End If
4540        End With
4550      End Select
4560    End If

        ' ** Shift keys.
4570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4580      Select Case intRetVal
          Case vbKeyTab
4590        With Me
4600          intRetVal = 0
4610          .cmdOK.SetFocus
4620        End With
4630      End Select
4640    End If

EXITP:
4650    KeyCode = intRetVal
4660    Exit Sub

ERRH:
4670    intRetVal = 0
4680    THAT_PROC = THIS_PROC
4690    That_Erl = Erl
4700    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4710    Resume EXITP

End Sub

Private Sub opgRecurType_AfterUpdate()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgRecurType_AfterUpdate"

4810    With Me
4820      Select Case .opgRecurType
          Case .opgRecurType_optMisc.OptionValue
4830        .opgRecurType_optMisc_lbl.FontBold = True
4840        .opgRecurType_optPaid_lbl.FontBold = False
4850        .opgRecurType_optReceived_lbl.FontBold = False
4860        .Recur_Type = "Misc"
4870        .Caption = "Add New Recurring Item"
4880        .Header_lbl_Item.Visible = True
4890        .Header_lbl_Payment.Visible = False
4900        .Header_lbl_Receipt.Visible = False
4910        .Recur_lbl.Caption = "Recurring Item"
4920        .Recur_Address.Visible = False
4930        .Recur_City.Visible = False
4940        .Recur_State.Visible = False
4950        .Recur_Zip.Visible = False
4960        .Recur_Country.Visible = False
4970        .Recur_PostalCode.Visible = False
4980        .opgLoc.Visible = False
4990        .opgLoc_box.Visible = False
5000      Case .opgRecurType_optPaid.OptionValue
5010        .opgRecurType_optMisc_lbl.FontBold = False
5020        .opgRecurType_optPaid_lbl.FontBold = True
5030        .opgRecurType_optReceived_lbl.FontBold = False
5040        .Recur_Type = "Payee"
5050        .Caption = "Add New Recurring Payment"
5060        .Header_lbl_Item.Visible = False
5070        .Header_lbl_Payment.Visible = True
5080        .Header_lbl_Receipt.Visible = False
5090        .Recur_lbl.Caption = "Recurring Payment"
5100        .Recur_Address.Visible = True
5110        .Recur_City.Visible = True
5120        .Recur_State.Visible = True
5130        .Recur_Zip.Visible = True
5140        .Recur_Country.Visible = True
5150        .Recur_PostalCode.Visible = True
5160        .opgLoc.Visible = True
5170        .opgLoc_box.Visible = True
5180      Case .opgRecurType_optReceived.OptionValue
5190        .opgRecurType_optMisc_lbl.FontBold = False
5200        .opgRecurType_optPaid_lbl.FontBold = False
5210        .opgRecurType_optReceived_lbl.FontBold = True
5220        .Recur_Type = "Payor"
5230        .Caption = "Add New Recurring Receipt"
5240        .Header_lbl_Item.Visible = False
5250        .Header_lbl_Payment.Visible = False
5260        .Header_lbl_Receipt.Visible = True
5270        .Recur_lbl.Caption = "Recurring Receipt"
5280        .Recur_Address.Visible = False
5290        .Recur_City.Visible = False
5300        .Recur_State.Visible = False
5310        .Recur_Zip.Visible = False
5320        .Recur_Country.Visible = False
5330        .Recur_PostalCode.Visible = False
5340        .opgLoc.Visible = False
5350        .opgLoc_box.Visible = False
5360      End Select
5370      .Recur_Name.SetFocus
5380    End With

EXITP:
5390    Exit Sub

ERRH:
5400    Select Case ERR.Number
        Case Else
5410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5420    End Select
5430    Resume EXITP

End Sub

Private Sub opgRecurType_optMisc_KeyDown(KeyCode As Integer, Shift As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgRecurType_optMisc_KeyDown"

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
5590          .Recur_Name.SetFocus
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5650        With Me
5660          intRetVal = 0
5670          .cmdCancel.SetFocus
5680        End With
5690      End Select
5700    End If

EXITP:
5710    KeyCode = intRetVal
5720    Exit Sub

ERRH:
5730    intRetVal = 0
5740    THAT_PROC = THIS_PROC
5750    That_Erl = Erl
5760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5770    Resume EXITP

End Sub

Private Sub opgRecurType_optPaid_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgRecurType_optPaid_KeyDown"

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
5890          .Recur_Name.SetFocus
5900        End With
5910      End Select
5920    End If

        ' ** Shift keys.
5930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5950        With Me
5960          intRetVal = 0
5970          .cmdCancel.SetFocus
5980        End With
5990      End Select
6000    End If

EXITP:
6010    KeyCode = intRetVal
6020    Exit Sub

ERRH:
6030    intRetVal = 0
6040    THAT_PROC = THIS_PROC
6050    That_Erl = Erl
6060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6070    Resume EXITP

End Sub

Private Sub opgRecurType_optReceived_KeyDown(KeyCode As Integer, Shift As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgRecurType_optReceived_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6120    intShiftDown = (Shift And acShiftMask) > 0
6130    intAltDown = (Shift And acAltMask) > 0
6140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6170        With Me
6180          intRetVal = 0
6190          .Recur_Name.SetFocus
6200        End With
6210      End Select
6220    End If

        ' ** Shift keys.
6230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6250        With Me
6260          intRetVal = 0
6270          .cmdCancel.SetFocus
6280        End With
6290      End Select
6300    End If

EXITP:
6310    KeyCode = intRetVal
6320    Exit Sub

ERRH:
6330    intRetVal = 0
6340    THAT_PROC = THIS_PROC
6350    That_Erl = Erl
6360    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6370    Resume EXITP

End Sub

Private Sub Recur_Name_AfterUpdate()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_AfterUpdate"

        Dim strTmp01 As String

6410    With Me
6420      If IsNull(.Recur_Name) = False Then
6430        strTmp01 = FixQuotes(.Recur_Name)  ' ** Module Function: modStringFuncs.
6440        If Compare_StringA_StringB(.Recur_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
6450          .Recur_Name = strTmp01
6460        End If
6470      End If
6480    End With

EXITP:
6490    Exit Sub

ERRH:
6500    Select Case ERR.Number
        Case Else
6510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6520    End Select
6530    Resume EXITP

End Sub

Private Sub Recur_Name_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6620    intShiftDown = (Shift And acShiftMask) > 0
6630    intAltDown = (Shift And acAltMask) > 0
6640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6670        With Me
6680          intRetVal = 0
6690          If .Recur_Address.Enabled = True And .Recur_Address.Visible = True Then
6700            .Recur_Address.SetFocus
6710          Else
6720            .cmdOK.SetFocus
6730          End If
6740        End With
6750      End Select
6760    End If

        ' ** Shift keys.
6770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6780      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6790        With Me
6800          intRetVal = 0
6810          If .opgRecurType.Visible = True And .opgRecurType.Enabled = True Then
6820            .opgRecurType.SetFocus
6830          Else
6840            .cmdCancel.SetFocus
6850          End If
6860        End With
6870      End Select
6880    End If

EXITP:
6890    KeyCode = intRetVal
6900    Exit Sub

ERRH:
6910    intRetVal = 0
6920    THAT_PROC = THIS_PROC
6930    That_Erl = Erl
6940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6950    Resume EXITP

End Sub

Private Sub Recur_Name_Exit(Cancel As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_Exit"

7010    With Me
7020      If IsNull(.Recur_Name) = True And .NewRecord = False Then
7030        MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
7040        Cancel = -1
7050      Else
7060        If Trim(.Recur_Name.text) = vbNullString And .NewRecord = False Then
7070          MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
7080          Cancel = -1
7090        End If
7100      End If
7110    End With

EXITP:
7120    Exit Sub

ERRH:
7130    THAT_PROC = THIS_PROC
7140    That_Erl = Erl
7150    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7160    Resume EXITP

End Sub

Private Sub Recur_Address_AfterUpdate()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Address_AfterUpdate"

        Dim strTmp01 As String

7210    With Me
7220      If IsNull(.Recur_Address) = False Then
7230        strTmp01 = FixQuotes(.Recur_Address)  ' ** Module Function: modStringFuncs.
7240        If Compare_StringA_StringB(.Recur_Address, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
7250          .Recur_Address = strTmp01
7260        End If
7270      End If
7280    End With

EXITP:
7290    Exit Sub

ERRH:
7300    Select Case ERR.Number
        Case Else
7310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7320    End Select
7330    Resume EXITP

End Sub

Private Sub Recur_City_AfterUpdate()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_City_AfterUpdate"

        Dim strTmp01 As String

7410    With Me
7420      If IsNull(.Recur_City) = False Then
7430        strTmp01 = FixQuotes(.Recur_City)  ' ** Module Function: modStringFuncs.
7440        If Compare_StringA_StringB(.Recur_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
7450          .Recur_City = strTmp01
7460        End If
7470      End If
7480    End With

EXITP:
7490    Exit Sub

ERRH:
7500    Select Case ERR.Number
        Case Else
7510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7520    End Select
7530    Resume EXITP

End Sub

Private Sub opgLoc_AfterUpdate()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_AfterUpdate"

7610    With Me
7620      Select Case .opgLoc
          Case .opgLoc_optUSA.OptionValue
7630        .Recur_State.Enabled = True
7640        .Recur_State.BorderColor = CLR_LTBLU2
7650        .Recur_State.BackStyle = acBackStyleNormal
7660        .Recur_Zip.Enabled = True
7670        .Recur_Zip.BorderColor = CLR_LTBLU2
7680        .Recur_Zip.BackStyle = acBackStyleNormal
7690        .Recur_Country.Enabled = False
7700        .Recur_Country.BorderColor = WIN_CLR_DISR
7710        .Recur_Country.BackStyle = acBackStyleTransparent
7720        .Recur_PostalCode.Enabled = False
7730        .Recur_PostalCode.BorderColor = WIN_CLR_DISR
7740        .Recur_PostalCode.BackStyle = acBackStyleTransparent
7750        Select Case blnIsOpen
            Case True
7760          blnIsOpen = False
7770        Case False
7780          .Recur_State.SetFocus
7790        End Select
7800      Case .opgLoc_optOther.OptionValue
7810        .Recur_Country.Enabled = True
7820        .Recur_Country.BorderColor = CLR_LTBLU2
7830        .Recur_Country.BackStyle = acBackStyleNormal
7840        .Recur_PostalCode.Enabled = True
7850        .Recur_PostalCode.BorderColor = CLR_LTBLU2
7860        .Recur_PostalCode.BackStyle = acBackStyleNormal
7870        .Recur_State.Enabled = False
7880        .Recur_State.BorderColor = WIN_CLR_DISR
7890        .Recur_State.BackStyle = acBackStyleTransparent
7900        .Recur_Zip.Enabled = False
7910        .Recur_Zip.BorderColor = WIN_CLR_DISR
7920        .Recur_Zip.BackStyle = acBackStyleTransparent
7930        Select Case blnIsOpen
            Case True
7940          blnIsOpen = False
7950        Case False
7960          .Recur_Country.SetFocus
7970        End Select
7980      End Select
7990    End With

EXITP:
8000    Exit Sub

ERRH:
8010    Select Case ERR.Number
        Case Else
8020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
8030    End Select
8040    Resume EXITP

End Sub

Private Sub opgLoc_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8120    intShiftDown = (Shift And acShiftMask) > 0
8130    intAltDown = (Shift And acAltMask) > 0
8140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8170        With Me
8180          intRetVal = 0
8190          If .Recur_State.Enabled = True Then
8200            .Recur_State.SetFocus
8210          Else
8220            .Recur_Country.SetFocus
8230          End If
8240        End With
8250      End Select
8260    End If

        ' ** Shift keys.
8270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8280      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8290        With Me
8300          intRetVal = 0
8310          .Recur_City.SetFocus
8320        End With
8330      End Select
8340    End If

EXITP:
8350    KeyCode = intRetVal
8360    Exit Sub

ERRH:
8370    intRetVal = 0
8380    THAT_PROC = THIS_PROC
8390    That_Erl = Erl
8400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8410    Resume EXITP

End Sub

Private Sub opgLoc_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optOther_KeyDown"

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
8590          If .Recur_State.Enabled = True Then
8600            .Recur_State.SetFocus
8610          Else
8620            .Recur_Country.SetFocus
8630          End If
8640        End With
8650      End Select
8660    End If

        ' ** Shift keys.
8670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8680      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8690        With Me
8700          intRetVal = 0
8710          .Recur_City.SetFocus
8720        End With
8730      End Select
8740    End If

EXITP:
8750    KeyCode = intRetVal
8760    Exit Sub

ERRH:
8770    intRetVal = 0
8780    THAT_PROC = THIS_PROC
8790    That_Erl = Erl
8800    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8810    Resume EXITP

End Sub

Private Sub Recur_State_Enter()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_Enter"

8910    With Me
8920      If IsNull(.Recur_State) = True Then
8930        ExpandCombo .Recur_State  ' ** Module Procedure: modUtilities.
8940      End If
8950    End With

EXITP:
8960    Exit Sub

ERRH:
8970    Select Case ERR.Number
        Case Else
8980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8990    End Select
9000    Resume EXITP

End Sub

Private Sub Recur_State_KeyUp(KeyCode As Integer, Shift As Integer)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_KeyUp"

9110    With Me
9120      If .Recur_State.SelStart <> 0 Then
9130        .Recur_State.SelStart = 1
9140        .Recur_State.SelLength = 1
9150      End If
9160    End With

EXITP:
9170    Exit Sub

ERRH:
9180    Select Case ERR.Number
        Case Else
9190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9200    End Select
9210    Resume EXITP

End Sub

Private Sub Recur_State_AfterUpdate()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_AfterUpdate"

9310    With Me
9320      If IsNull(.Recur_State) = False Then
9330        If Trim(.Recur_State) <> vbNullString Then
9340          Select Case IsNull(.Recur_State_Pref)
              Case True
9350            .Recur_State_Pref = .Recur_State
9360          Case False
9370            If .Recur_State_Pref <> .Recur_State Then
9380              .Recur_State_Pref = .Recur_State
9390            End If
9400          End Select
9410        End If
9420      End If
9430    End With

EXITP:
9440    Exit Sub

ERRH:
9450    Select Case ERR.Number
        Case Else
9460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9470    End Select
9480    Resume EXITP

End Sub

Private Sub Recur_State_NotInList(NewData As String, Response As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_NotInList"

9510    With Me
9520      .Recur_State = Null
9530      MsgBox "Must select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
9540      Response = acDataErrContinue
9550    End With

EXITP:
9560    Exit Sub

ERRH:
9570    Select Case ERR.Number
        Case Else
9580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9590    End Select
9600    Resume EXITP

End Sub

Private Sub Recur_Zip_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Zip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9720    intShiftDown = (Shift And acShiftMask) > 0
9730    intAltDown = (Shift And acAltMask) > 0
9740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9770        With Me
9780          intRetVal = 0
9790          If .RecurringItemAux_1099.Visible = True And .RecurringItemAux_1099.Enabled = True Then
9800            .RecurringItemAux_1099.SetFocus
9810          Else
9820            .cmdOK.SetFocus
9830          End If
9840        End With
9850      End Select
9860    End If

        ' ** Shift keys.
9870    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9880      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9890        With Me
9900          intRetVal = 0
9910          .Recur_State.SetFocus
9920        End With
9930      End Select
9940    End If

EXITP:
9950    KeyCode = intRetVal
9960    Exit Sub

ERRH:
9970    intRetVal = 0
9980    THAT_PROC = THIS_PROC
9990    That_Erl = Erl
10000   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10010   Resume EXITP

End Sub

Private Sub Recur_Country_AfterUpdate()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Country_AfterUpdate"

        Dim strTmp01 As String

10110   With Me
10120     If IsNull(.Recur_Country) = False Then
10130       strTmp01 = FixQuotes(.Recur_Country)  ' ** Module Function: modStringFuncs.
10140       If Compare_StringA_StringB(.Recur_Country, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
10150         .Recur_Country = strTmp01
10160       End If
10170     End If
10180   End With

EXITP:
10190   Exit Sub

ERRH:
10200   Select Case ERR.Number
        Case Else
10210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10220   End Select
10230   Resume EXITP

End Sub

Private Sub Recur_Country_KeyDown(KeyCode As Integer, Shift As Integer)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Country_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10320   intShiftDown = (Shift And acShiftMask) > 0
10330   intAltDown = (Shift And acAltMask) > 0
10340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10370       With Me
10380         intRetVal = 0
10390         .Recur_PostalCode.SetFocus
10400       End With
10410     End Select
10420   End If

        ' ** Shift keys.
10430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10450       With Me
10460         intRetVal = 0
10470         If .Recur_Zip.Enabled = True Then
10480           .Recur_Zip.SetFocus
10490         Else
10500           .Recur_City.SetFocus
10510         End If
10520       End With
10530     End Select
10540   End If

EXITP:
10550   KeyCode = intRetVal
10560   Exit Sub

ERRH:
10570   intRetVal = 0
10580   THAT_PROC = THIS_PROC
10590   That_Erl = Erl
10600   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10610   Resume EXITP

End Sub

Private Sub Recur_PostalCode_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_PostalCode_KeyDown"

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
10790         If .RecurringItemAux_1099.Visible = True And .RecurringItemAux_1099.Enabled = True Then
10800           .RecurringItemAux_1099.SetFocus
10810         Else
10820           .cmdOK.SetFocus
10830         End If
10840       End With
10850     End Select
10860   End If

        ' ** Shift keys.
10870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10880     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10890       With Me
10900         intRetVal = 0
10910         .Recur_Country.SetFocus
10920       End With
10930     End Select
10940   End If

EXITP:
10950   KeyCode = intRetVal
10960   Exit Sub

ERRH:
10970   intRetVal = 0
10980   THAT_PROC = THIS_PROC
10990   That_Erl = Erl
11000   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11010   Resume EXITP

End Sub

Private Sub RecurringItemAux_1099_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_1099_AfterUpdate"

11110   With Me
11120     Select Case .RecurringItemAux_1099
          Case True
11130       .RecurringItemAux_TaxNo.Enabled = True
11140       .RecurringItemAux_TaxNo.BorderColor = CLR_LTBLU2
11150       .RecurringItemAux_TaxNo.BackStyle = acBackStyleNormal
11160     Case False
11170       .RecurringItemAux_TaxNo.Enabled = False
11180       .RecurringItemAux_TaxNo.BorderColor = WIN_CLR_DISR
11190       .RecurringItemAux_TaxNo.BackStyle = acBackStyleTransparent
11200     End Select
11210   End With

EXITP:
11220   Exit Sub

ERRH:
11230   Select Case ERR.Number
        Case Else
11240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11250   End Select
11260   Resume EXITP

End Sub

Private Sub RecurringItemAux_1099_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_1099_KeyDown"

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
11390         Select Case .RecurringItemAux_1099
              Case True
11400           .RecurringItemAux_TaxNo.SetFocus
11410         Case False
11420           .cmdOK.SetFocus
11430         End Select
11440       End With
11450     End Select
11460   End If

        ' ** Shift keys.
11470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11490       With Me
11500         intRetVal = 0
11510         If .Recur_PostalCode.Enabled = True Then
11520           .Recur_PostalCode.SetFocus
11530         Else
11540           .Recur_Zip.SetFocus
11550         End If
11560       End With
11570     End Select
11580   End If

EXITP:
11590   KeyCode = intRetVal
11600   Exit Sub

ERRH:
11610   intRetVal = 0
11620   THAT_PROC = THIS_PROC
11630   That_Erl = Erl
11640   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11650   Resume EXITP

End Sub

Private Sub RecurringItemAux_TaxNo_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_TaxNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11720   intShiftDown = (Shift And acShiftMask) > 0
11730   intAltDown = (Shift And acAltMask) > 0
11740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11770       With Me
11780         intRetVal = 0
11790         .cmdOK.SetFocus
11800       End With
11810     End Select
11820   End If

        ' ** Shift keys.
11830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11850       With Me
11860         intRetVal = 0
11870         .RecurringItemAux_1099.SetFocus
11880       End With
11890     End Select
11900   End If

EXITP:
11910   KeyCode = intRetVal
11920   Exit Sub

ERRH:
11930   intRetVal = 0
11940   THAT_PROC = THIS_PROC
11950   That_Erl = Erl
11960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11970   Resume EXITP

End Sub
