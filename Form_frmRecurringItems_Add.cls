VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRecurringItems_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRecurringItems_Add"

'VGC 09/20/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Description:      D {Recur_Name}
' **   Form 1099:        F {RecurringItemAux_1099}
' **   Other:            H {opgLoc_optOther}
' **   OK:               O {cmdOK}
' **   USA:              U {opgLoc_optUSA}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private lngRecurID As Long, blnIsOpen As Boolean
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

170           CLR_DISABLED_FG = CLR_DKGRY
180           CLR_DISABLED_BG = CLR_LTTEAL

190           DoCmd.RunCommand acCmdRecordsGoToNew

200         Else
210           Cancel = -1
220         End If
230       Else
240         Cancel = -1
250       End If
260     End With

270     If Cancel = -1 Then
280       Beep
290       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
300     End If

EXITP:
310     Exit Sub

ERRH:
320     Select Case ERR.Number
        Case Else
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
340     End Select
350     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me

420       Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

430       .Recur_Type.ForeColor = CLR_DISABLED_FG
440       .Recur_Type.BackColor = CLR_DISABLED_BG

450       If IsNull(.Recur_State_Pref) = False Then
460         .Recur_State.DefaultValue = Chr(34) & .Recur_State_Pref & Chr(34)
470       End If

480       .opgLoc.Enabled = True
490       opgLoc_AfterUpdate  ' ** Procedure: Below.

500       lngRecurID = 0&
510       .Recur_Name.SetFocus

520     End With

EXITP:
530     Exit Sub

ERRH:
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

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

        ' ** Shortcut Alt keys
        ' **   Other:            H {opgLoc_optOther}
        ' **   USA:              U {opgLoc_optUSA}

        ' ** Shortcut Ctrl keys to other forms and functionality.
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyEscape
670         With Me
680           intRetVal = 0
690           If .NewRecord = True Then
700             .Undo
710           Else
720             DoCmd.SetWarnings False
730             DoCmd.RunCommand acCmdDeleteRecord
740             DoCmd.SetWarnings True
750           End If
760           DoCmd.OpenForm "frmRecurringItems"
770           DoCmd.Close acForm, THIS_NAME
780         End With
790       End Select
800     End If

        ' ** Alt keys.
810     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
820       Select Case intRetVal
          Case vbKeyH
830         With Me
840           intRetVal = 0
850           If .opgLoc.Enabled = True Then
860             .opgLoc.SetFocus
870             .opgLoc = .opgLoc_optOther.OptionValue
880             opgLoc_AfterUpdate  ' ** Procedure: Below.
890           End If
900         End With
910       Case vbKeyU
920         With Me
930           intRetVal = 0
940           If .opgLoc.Enabled = True Then
950             .opgLoc.SetFocus
960             .opgLoc = .opgLoc_optUSA.OptionValue
970             opgLoc_AfterUpdate  ' ** Procedure: Below.
980           End If
990         End With
1000      End Select
1010    End If

        ' ** Ctl keys.
1020    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1030      Select Case intRetVal
          Case vbKeyS
1040        intRetVal = 0
1050        cmdSave_Click  ' ** Procedure: Below.
1060      End Select
1070    End If

        ' ** Ctl-Shift keys.
1080    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1090      Select Case intRetVal
          Case vbKeyF
1100        With Me
1110          intRetVal = 0
1120          .FocusHolder.SetFocus
1130        End With
1140      End Select
1150    End If

EXITP:
1160    KeyCode = intRetVal
1170    Exit Sub

ERRH:
1180    intRetVal = 0
1190    THAT_PROC = THIS_PROC
1200    That_Erl = Erl
1210    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1220    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1310    If DataErr <> 0 Then
1320      Select Case DataErr
          Case 3022  ' ** The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.
1330        MsgBox "A Recurring Item by this name already exists for this type.", vbInformation + vbOKOnly, "Duplicate Recurring Item"
1340        Response = acDataErrContinue
1350      Case 3058  ' ** Index or primary key can't contain a Null value.
1360        Select Case Me.ActiveControl.Name
            Case "Recur_Name"
1370          If IsNull(Me.Recur_Name) Then
1380            MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
1390            Me.Recur_Name.SetFocus
1400            Response = acDataErrContinue
1410          Else
1420            Select Case Forms("frmRecurringItems").opgType.Value
                Case 2
1430              Me.Recur_Type = "Payee"
1440            Case 1
1450              Me.Recur_Type = "Payor"
1460            Case 3
1470              Me.Recur_Type = "Misc"
1480            End Select
1490            Response = acDataErrContinue
1500          End If
1510        Case Else
1520          If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1530          zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1540        End Select
1550      End Select
1560    End If

EXITP:
1570    Exit Sub

ERRH:
1580    Select Case ERR.Number
        Case Else
1590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1600    End Select
1610    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1710    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1720    Select Case strCallingForm
        Case "frmRecurringItems"
1730      gblnSetFocus = True
1740      gstrReturningForm = THIS_NAME
1750      DoCmd.SelectObject acForm, strCallingForm, False
1760      Forms(strCallingForm).TimerInterval = 100&
1770    End Select

EXITP:
1780    Exit Sub

ERRH:
1790    Select Case ERR.Number
        Case Else
1800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1810    End Select
1820    Resume EXITP

End Sub

Private Sub cmdSave_Click()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

1910    DoCmd.RunCommand acCmdSelectRecord
1920    DoCmd.RunCommand acCmdSaveRecord

EXITP:
1930    Exit Sub

ERRH:
1940    Select Case ERR.Number
        Case Else
1950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1960    End Select
1970    Resume EXITP

End Sub

Private Sub cmdOK_Click()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

2010    With Me
2020      If IsNull(.Recur_Name) = True Then
2030        MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
2040        .Recur_Name.SetFocus
2050      Else
2060        If .Recur_Name = vbNullString Then
2070          MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
2080          .Recur_Name.SetFocus
2090        Else
2100          .Recur_Name.SetFocus
2110          If InStr(.Recur_Name.text, Chr(34)) > 0 Then
2120            MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
2130            .cmdOK.SetFocus
2140            DoCmd.CancelEvent
2150          Else
                ' ** If they've flipped around, only take the final choice.
2160            Select Case .opgLoc
                Case .opgLoc_optUSA.OptionValue
2170              If IsNull(.Recur_Country) = False Then
2180                .Recur_Country = Null
2190              End If
2200              If IsNull(.Recur_PostalCode) = False Then
2210                .Recur_PostalCode = Null
2220              End If
2230            Case .opgLoc_optOther.OptionValue
2240              If IsNull(.Recur_State) = False Then
2250                .Recur_State = Null
2260              End If
2270              If IsNull(.Recur_Zip) = False Then
2280                .Recur_Zip = Null
2290              End If
2300            End Select
2310            lngRecurID = .RecurringItem_ID
2320            Set dbs = CurrentDb
                ' ** RecurringItems, by specified [nam], [typ].
2330            Set qdf = dbs.QueryDefs("qryRecurringItems_04")
2340            With qdf.Parameters
2350              ![nam] = Me.Recur_Name
2360              ![typ] = Me.Recur_Type
2370            End With
2380            Set rst = qdf.OpenRecordset
2390            If rst.BOF = True And rst.EOF = True Then
                  ' ** All's well.
2400              rst.Close
2410              dbs.Close
2420              gdblCrtRpt_CostTot = lngRecurID  ' ** Borrowing this Public variable.
2430              gblnSetFocus = True
2440              Forms("frmRecurringItems").TimerInterval = 500&
2450              DoCmd.Close acForm, THIS_NAME
2460            Else
2470              rst.Close
2480              dbs.Close
2490              MsgBox "A Recurring Item by this name already exists for this type.", vbInformation + vbOKOnly, "Duplicate Recurring Item"
2500              .Recur_Name.SetFocus
2510            End If
2520          End If
2530        End If
2540      End If
2550    End With

EXITP:
2560    Set rst = Nothing
2570    Set qdf = Nothing
2580    Set dbs = Nothing
2590    Exit Sub

ERRH:
2600    Select Case ERR.Number
        Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
2770        With Me
2780          intRetVal = 0
2790          .cmdCancel.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab
2850        With Me
2860          intRetVal = 0
2870          If .RecurringItemAux_TaxNo.Enabled = True Then
2880            .RecurringItemAux_TaxNo.SetFocus
2890          ElseIf .RecurringItemAux_1099.Enabled = True Then
2900            .RecurringItemAux_1099.SetFocus
2910          Else
2920            If .Recur_Address.Enabled = True And .Recur_Address.Visible = True Then
2930              If .Recur_Country.Enabled = True Then
2940                .Recur_PostalCode.SetFocus
2950              Else
2960                .Recur_Zip.SetFocus
2970              End If
2980            Else
2990              .Recur_Name.SetFocus
3000            End If
3010          End If
3020        End With
3030      End Select
3040    End If

EXITP:
3050    KeyCode = intRetVal
3060    Exit Sub

ERRH:
3070    intRetVal = 0
3080    THAT_PROC = THIS_PROC
3090    That_Erl = Erl
3100    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3110    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

3210    With Me
3220      If .NewRecord Then
3230        DoCmd.RunCommand acCmdUndo
3240      Else
3250        DoCmd.SetWarnings False
3260        DoCmd.RunCommand acCmdDeleteRecord
3270        DoCmd.SetWarnings True
3280      End If
3290    End With
3300    DoCmd.Close acForm, THIS_NAME

EXITP:
3310    Exit Sub

ERRH:
3320    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3330    Case 3031, 3101
3340      DoCmd.RunCommand acCmdUndo
3350    Case Else
3360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3370    End Select
3380    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
          Case vbKeyTab
3470        With Me
3480          intRetVal = 0
3490          .Recur_Name.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab
3550        With Me
3560          intRetVal = 0
3570          .cmdOK.SetFocus
3580        End With
3590      End Select
3600    End If

EXITP:
3610    KeyCode = intRetVal
3620    Exit Sub

ERRH:
3630    intRetVal = 0
3640    THAT_PROC = THIS_PROC
3650    That_Erl = Erl
3660    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3670    Resume EXITP

End Sub

Private Sub Recur_Name_AfterUpdate()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_AfterUpdate"

        Dim strTmp01 As String

3710    With Me
3720      If IsNull(.Recur_Name) = False Then
3730        strTmp01 = FixQuotes(.Recur_Name)  ' ** Module Function: modStringFuncs.
3740        If Compare_StringA_StringB(.Recur_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
3750          .Recur_Name = strTmp01
3760        End If
3770      End If
3780    End With

EXITP:
3790    Exit Sub

ERRH:
3800    Select Case ERR.Number
        Case Else
3810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3820    End Select
3830    Resume EXITP

End Sub

Private Sub Recur_Name_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_KeyDown"

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
3990          If .Recur_Address.Enabled = True And .Recur_Address.Visible = True Then
4000            .Recur_Address.SetFocus
4010          Else
4020            .cmdOK.SetFocus
4030          End If
4040        End With
4050      End Select
4060    End If

        ' ** Shift keys.
4070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4080      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4090        With Me
4100          intRetVal = 0
4110          .cmdCancel.SetFocus
4120        End With
4130      End Select
4140    End If

EXITP:
4150    KeyCode = intRetVal
4160    Exit Sub

ERRH:
4170    intRetVal = 0
4180    THAT_PROC = THIS_PROC
4190    That_Erl = Erl
4200    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4210    Resume EXITP

End Sub

Private Sub Recur_Name_Exit(Cancel As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Name_Exit"

4310    With Me
4320      If IsNull(.Recur_Name) = True And .NewRecord = False Then
4330        MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
4340        Cancel = -1
4350      Else
4360        If .Recur_Name.text = vbNullString And .NewRecord = False Then
4370          MsgBox "A name must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
4380          Cancel = -1
4390        End If
4400      End If
4410    End With

EXITP:
4420    Exit Sub

ERRH:
4430    THAT_PROC = THIS_PROC
4440    That_Erl = Erl
4450    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4460    Resume EXITP

End Sub

Private Sub Recur_Address_AfterUpdate()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Address_AfterUpdate"

        Dim strTmp01 As String

4510    With Me
4520      If IsNull(.Recur_Address) = False Then
4530        strTmp01 = FixQuotes(.Recur_Address)  ' ** Module Function: modStringFuncs.
4540        If Compare_StringA_StringB(.Recur_Address, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
4550          .Recur_Address = strTmp01
4560        End If
4570      End If
4580    End With

EXITP:
4590    Exit Sub

ERRH:
4600    Select Case ERR.Number
        Case Else
4610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4620    End Select
4630    Resume EXITP

End Sub

Private Sub Recur_City_AfterUpdate()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_City_AfterUpdate"

        Dim strTmp01 As String

4710    With Me
4720      If IsNull(.Recur_City) = False Then
4730        strTmp01 = FixQuotes(.Recur_City)  ' ** Module Function: modStringFuncs.
4740        If Compare_StringA_StringB(.Recur_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
4750          .Recur_City = strTmp01
4760        End If
4770      End If
4780    End With

EXITP:
4790    Exit Sub

ERRH:
4800    Select Case ERR.Number
        Case Else
4810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4820    End Select
4830    Resume EXITP

End Sub

Private Sub opgLoc_AfterUpdate()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_AfterUpdate"

4910    With Me
4920      Select Case .opgLoc
          Case .opgLoc_optUSA.OptionValue
4930        .Recur_State.Enabled = True
4940        .Recur_State.BorderColor = CLR_LTBLU2
4950        .Recur_State.BackStyle = acBackStyleNormal
4960        .Recur_Zip.Enabled = True
4970        .Recur_Zip.BorderColor = CLR_LTBLU2
4980        .Recur_Zip.BackStyle = acBackStyleNormal
4990        .Recur_Country.Enabled = False
5000        .Recur_Country.BorderColor = WIN_CLR_DISR
5010        .Recur_Country.BackStyle = acBackStyleTransparent
5020        .Recur_PostalCode.Enabled = False
5030        .Recur_PostalCode.BorderColor = WIN_CLR_DISR
5040        .Recur_PostalCode.BackStyle = acBackStyleTransparent
5050        Select Case blnIsOpen
            Case True
5060          blnIsOpen = False
5070        Case False
5080          .Recur_State.SetFocus
5090        End Select
5100      Case .opgLoc_optOther.OptionValue
5110        .Recur_Country.Enabled = True
5120        .Recur_Country.BorderColor = CLR_LTBLU2
5130        .Recur_Country.BackStyle = acBackStyleNormal
5140        .Recur_PostalCode.Enabled = True
5150        .Recur_PostalCode.BorderColor = CLR_LTBLU2
5160        .Recur_PostalCode.BackStyle = acBackStyleNormal
5170        .Recur_State.Enabled = False
5180        .Recur_State.BorderColor = WIN_CLR_DISR
5190        .Recur_State.BackStyle = acBackStyleTransparent
5200        .Recur_Zip.Enabled = False
5210        .Recur_Zip.BorderColor = WIN_CLR_DISR
5220        .Recur_Zip.BackStyle = acBackStyleTransparent
5230        Select Case blnIsOpen
            Case True
5240          blnIsOpen = False
5250        Case False
5260          .Recur_Country.SetFocus
5270        End Select
5280      End Select
5290    End With

EXITP:
5300    Exit Sub

ERRH:
5310    Select Case ERR.Number
        Case Else
5320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
5330    End Select
5340    Resume EXITP

End Sub

Private Sub opgLoc_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optUSA_KeyDown"

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
5490          If .Recur_State.Enabled = True Then
5500            .Recur_State.SetFocus
5510          Else
5520            .Recur_Country.SetFocus
5530          End If
5540        End With
5550      End Select
5560    End If

        ' ** Shift keys.
5570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5580      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5590        With Me
5600          intRetVal = 0
5610          .Recur_City.SetFocus
5620        End With
5630      End Select
5640    End If

EXITP:
5650    KeyCode = intRetVal
5660    Exit Sub

ERRH:
5670    intRetVal = 0
5680    THAT_PROC = THIS_PROC
5690    That_Erl = Erl
5700    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5710    Resume EXITP

End Sub

Private Sub opgLoc_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optOther_KeyDown"

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
5890          If .Recur_State.Enabled = True Then
5900            .Recur_State.SetFocus
5910          Else
5920            .Recur_Country.SetFocus
5930          End If
5940        End With
5950      End Select
5960    End If

        ' ** Shift keys.
5970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5980      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5990        With Me
6000          intRetVal = 0
6010          .Recur_City.SetFocus
6020        End With
6030      End Select
6040    End If

EXITP:
6050    KeyCode = intRetVal
6060    Exit Sub

ERRH:
6070    intRetVal = 0
6080    THAT_PROC = THIS_PROC
6090    That_Erl = Erl
6100    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6110    Resume EXITP

End Sub

Private Sub Recur_State_Enter()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_Enter"

6210    With Me
6220      If IsNull(.Recur_State) = True Then
6230        ExpandCombo .Recur_State  ' ** Module Procedure: modUtilities.
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

Private Sub Recur_State_KeyUp(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_KeyUp"

6410    With Me
6420      If .Recur_State.SelStart <> 0 Then
6430        .Recur_State.SelStart = 1
6440        .Recur_State.SelLength = 1
6450      End If
6460    End With

EXITP:
6470    Exit Sub

ERRH:
6480    Select Case ERR.Number
        Case Else
6490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6500    End Select
6510    Resume EXITP

End Sub

Private Sub Recur_State_AfterUpdate()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_AfterUpdate"

6610    With Me
6620      If IsNull(.Recur_State) = False Then
6630        If Trim(.Recur_State) <> vbNullString Then
6640          Select Case IsNull(.Recur_State_Pref)
              Case True
6650            .Recur_State_Pref = .Recur_State
6660          Case False
6670            If .Recur_State_Pref <> .Recur_State Then
6680              .Recur_State_Pref = .Recur_State
6690            End If
6700          End Select
6710        End If
6720      End If
6730    End With

EXITP:
6740    Exit Sub

ERRH:
6750    Select Case ERR.Number
        Case Else
6760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6770    End Select
6780    Resume EXITP

End Sub

Private Sub Recur_State_NotInList(NewData As String, Response As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_State_NotInList"

6810    With Me
6820      .Recur_State = Null
6830      MsgBox "Must select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
6840      Response = acDataErrContinue
6850    End With

EXITP:
6860    Exit Sub

ERRH:
6870    Select Case ERR.Number
        Case Else
6880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6890    End Select
6900    Resume EXITP

End Sub

Private Sub Recur_Zip_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Zip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7020    intShiftDown = (Shift And acShiftMask) > 0
7030    intAltDown = (Shift And acAltMask) > 0
7040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7070        With Me
7080          intRetVal = 0
7090          If .RecurringItemAux_1099.Enabled = True Then
7100            .RecurringItemAux_1099.SetFocus
7110          Else
7120            .cmdOK.SetFocus
7130          End If
7140        End With
7150      End Select
7160    End If

        ' ** Shift keys.
7170    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7180      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7190        With Me
7200          intRetVal = 0
7210          .Recur_State.SetFocus
7220        End With
7230      End Select
7240    End If

EXITP:
7250    KeyCode = intRetVal
7260    Exit Sub

ERRH:
7270    intRetVal = 0
7280    THAT_PROC = THIS_PROC
7290    That_Erl = Erl
7300    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7310    Resume EXITP

End Sub

Private Sub Recur_Country_AfterUpdate()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Country_AfterUpdate"

        Dim strTmp01 As String

7410    With Me
7420      If IsNull(.Recur_Country) = False Then
7430        strTmp01 = FixQuotes(.Recur_Country)  ' ** Module Function: modStringFuncs.
7440        If Compare_StringA_StringB(.Recur_Country, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
7450          .Recur_Country = strTmp01
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

Private Sub Recur_Country_KeyDown(KeyCode As Integer, Shift As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_Country_KeyDown"

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
7690          .Recur_PostalCode.SetFocus
7700        End With
7710      End Select
7720    End If

        ' ** Shift keys.
7730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7750        With Me
7760          intRetVal = 0
7770          If .Recur_Zip.Enabled = True Then
7780            .Recur_Zip.SetFocus
7790          Else
7800            .Recur_City.SetFocus
7810          End If
7820        End With
7830      End Select
7840    End If

EXITP:
7850    KeyCode = intRetVal
7860    Exit Sub

ERRH:
7870    intRetVal = 0
7880    THAT_PROC = THIS_PROC
7890    That_Erl = Erl
7900    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7910    Resume EXITP

End Sub

Private Sub Recur_PostalCode_KeyDown(KeyCode As Integer, Shift As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "Recur_PostalCode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8020    intShiftDown = (Shift And acShiftMask) > 0
8030    intAltDown = (Shift And acAltMask) > 0
8040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8070        With Me
8080          intRetVal = 0
8090          If .RecurringItemAux_1099.Enabled = True Then
8100            .RecurringItemAux_1099.SetFocus
8110          Else
8120            .cmdOK.SetFocus
8130          End If
8140        End With
8150      End Select
8160    End If

        ' ** Shift keys.
8170    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8180      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8190        With Me
8200          intRetVal = 0
8210          .Recur_Country.SetFocus
8220        End With
8230      End Select
8240    End If

EXITP:
8250    KeyCode = intRetVal
8260    Exit Sub

ERRH:
8270    intRetVal = 0
8280    THAT_PROC = THIS_PROC
8290    That_Erl = Erl
8300    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8310    Resume EXITP

End Sub

Private Sub RecurringItemAux_1099_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_1099_AfterUpdate"

8410    With Me
8420      Select Case .RecurringItemAux_1099
          Case True
8430        .RecurringItemAux_TaxNo.Enabled = True
8440        .RecurringItemAux_TaxNo.BorderColor = CLR_LTBLU2
8450        .RecurringItemAux_TaxNo.BackStyle = acBackStyleNormal
8460      Case False
8470        .RecurringItemAux_TaxNo.Enabled = False
8480        .RecurringItemAux_TaxNo.BorderColor = WIN_CLR_DISR
8490        .RecurringItemAux_TaxNo.BackStyle = acBackStyleTransparent
8500      End Select
8510    End With

EXITP:
8520    Exit Sub

ERRH:
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Sub

Private Sub RecurringItemAux_1099_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_1099_KeyDown"

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
8690          Select Case .RecurringItemAux_1099
              Case True
8700            .RecurringItemAux_TaxNo.SetFocus
8710          Case False
8720            .cmdOK.SetFocus
8730          End Select
8740        End With
8750      End Select
8760    End If

        ' ** Shift keys.
8770    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8780      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8790        With Me
8800          intRetVal = 0
8810          If .Recur_PostalCode.Enabled = True Then
8820            .Recur_PostalCode.SetFocus
8830          Else
8840            .Recur_Zip.SetFocus
8850          End If
8860        End With
8870      End Select
8880    End If

EXITP:
8890    KeyCode = intRetVal
8900    Exit Sub

ERRH:
8910    intRetVal = 0
8920    THAT_PROC = THIS_PROC
8930    That_Erl = Erl
8940    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8950    Resume EXITP

End Sub

Private Sub RecurringItemAux_TaxNo_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "RecurringItemAux_TaxNo_KeyDown"

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
9090          .cmdOK.SetFocus
9100        End With
9110      End Select
9120    End If

        ' ** Shift keys.
9130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9150        With Me
9160          intRetVal = 0
9170          .RecurringItemAux_1099.SetFocus
9180        End With
9190      End Select
9200    End If

EXITP:
9210    KeyCode = intRetVal
9220    Exit Sub

ERRH:
9230    intRetVal = 0
9240    THAT_PROC = THIS_PROC
9250    That_Erl = Erl
9260    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9270    Resume EXITP

End Sub
