VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmIncomeExpenseCodes_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmIncomeExpenseCodes_Add"

'VGC 11/21/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Active:           A {revcode_ACTIVE}
' **   Cancel:           C {cmdCancel}
' **   Description:      D {revcode_DESC}
' **   Expense:          E {opgIncomeExpense_optExpense}
' **   Income:           I {opgIncomeExpense_optIncome}
' **   OK:               O {cmdOK}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private lngRevcodeID As Long, lngOpenType As Long
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = .OpenArgs

150           CLR_DISABLED_FG = CLR_DKGRY
160           CLR_DISABLED_BG = CLR_LTTEAL

170           gdblCrtRpt_CostTot = 0#  ' ** Borrowing this public variable.

180           lngOpenType = Forms(strCallingForm).frmIncomeExpenseCodes_Sub.Form.revcode_TYPE
190           .opgIncomeExpense.DefaultValue = 0

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

420       .revcode_SORTORDER.ForeColor = CLR_DISABLED_FG
430       .revcode_SORTORDER.BackColor = CLR_DISABLED_BG

440       lngRevcodeID = DMax("[revcode_ID]", "tmpRevCodeEdit") + 1&

450       Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

460       If .opgIncomeExpense > 0 Then  ' ** Preference will override type they're on.
470         lngOpenType = .opgIncomeExpense
480       Else
490         .opgIncomeExpense = lngOpenType
500       End If

510       .opgIncomeExpense.DefaultValue = lngOpenType
520       Select Case .opgIncomeExpense.DefaultValue
          Case .opgIncomeExpense_optIncome.OptionValue
530         .revcode_TYPE_Description.DefaultValue = Chr(34) & "Income" & Chr(34)
540       Case .opgIncomeExpense_optExpense.OptionValue
550         .revcode_TYPE_Description.DefaultValue = Chr(34) & "Expense" & Chr(34)
560       End Select

570       opgIncomeExpense_AfterUpdate  ' ** Procedure: Below.

580       .revcode_DESC.SetFocus
590       .revcode_ID = lngRevcodeID
600       .AllowAdditions = False

610     End With

EXITP:
620     Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

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
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         With Me
780           intRetVal = 0
790           If .Dirty = True Then
800             .Undo
810           Else
820             DoCmd.SetWarnings False
830             DoCmd.RunCommand acCmdDeleteRecord
840             DoCmd.SetWarnings True
850           End If
860           DoCmd.Close acForm, THIS_NAME
870         End With
880       Case vbKeyDelete
890         intRetVal = 0
900       End Select
910     End If

        ' ** Ctrl keys.
920     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
930       Select Case intRetVal
          Case vbKeyS
940         intRetVal = 0
950         cmdSave_Click  ' ** Procedure: Below.
960       End Select
970     End If

        ' ** Ctrl-Shift keys.
980     If intCtrlDown And (Not intAltDown) And intShiftDown Then
990       Select Case intRetVal
          Case vbKeyF
1000        With Me
1010          intRetVal = 0
1020          .FocusHolder.SetFocus
1030        End With
1040      End Select
1050    End If

EXITP:
1060    KeyCode = intRetVal
1070    Exit Sub

ERRH:
1080    intRetVal = 0
1090    THAT_PROC = THIS_PROC
1100    That_Erl = Erl
1110    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1120    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1210    If DataErr <> 0 Then
1220      Select Case DataErr
          Case 3022  ' ** The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.
1230        MsgBox "An " & Me.revcode_TYPE_Description & " code by this name already exists.", _
              vbInformation + vbOKOnly, "Duplicate " & Me.revcode_TYPE_Description & " Code"
1240        Response = acDataErrContinue
1250      Case 3058  ' ** Index or primary key can't contain a Null value.
1260        Select Case Me.ActiveControl.Name
            Case "revcode_DESC"
1270          If IsNull(Me.revcode_DESC) Then
1280            MsgBox "A Description must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
1290            Me.revcode_DESC.SetFocus
1300            Response = acDataErrContinue
1310            Exit Sub
1320          End If
1330        Case Else
1340          If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1350          zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1360        End Select
1370      End Select
1380    End If

EXITP:
1390    Exit Sub

ERRH:
1400    Select Case ERR.Number
        Case Else
1410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1420    End Select
1430    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1510    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1520    gblnSetFocus = True
1530    gstrReturningForm = THIS_NAME
1540    DoCmd.SelectObject acForm, strCallingForm, False
1550    Forms(strCallingForm).TimerInterval = 100&

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

1710    With Me
1720      If IsNull(.revcode_ID) = True Then
1730        .revcode_ID = lngRevcodeID
1740      End If
1750      If IsNull(.revcode_DESC) = True Then
1760        MsgBox "A Description must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
1770        .revcode_DESC.SetFocus
1780      Else
1790        If .revcode_DESC = vbNullString Then
1800          MsgBox "A Description must be entered to continue.", vbInformation + vbOKOnly, "Entry Required"
1810          .revcode_DESC.SetFocus
1820        Else
1830          .revcode_DESC.SetFocus
1840          If InStr(.revcode_DESC.text, Chr(34)) > 0 Then
1850            MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
1860            DoCmd.CancelEvent
1870          Else
1880            lngRevcodeID = .revcode_ID
1890            Set dbs = CurrentDb
                ' ** tmpRevCodeEdit, by specified [revdsc], [revtyp].
1900            Set qdf = dbs.QueryDefs("qryRevCodes_13")
1910            With qdf.Parameters
1920              ![revdsc] = Me.revcode_DESC
1930              ![revtyp] = Me.opgIncomeExpense
1940            End With
1950            Set rst = qdf.OpenRecordset
1960            If rst.BOF = True And rst.EOF = True Then
                  ' ** All's well.
1970              rst.Close
1980              dbs.Close
1990              gdblCrtRpt_CostTot = lngRevcodeID  ' ** Borrowing this Public variable.
                  ' ** MAKE SURE cmdUpdate ENABLED, AND THE FORM KNOWS THERE'S SOMETHING NEW!
2000              DoCmd.Close acForm, THIS_NAME
2010            Else
2020              rst.Close
2030              dbs.Close
2040              MsgBox "An " & .revcode_TYPE_Description & " code by this name already exists.", _
                    vbInformation + vbOKOnly, "Duplicate " & .revcode_TYPE_Description & " Code"
2050              .revcode_DESC.SetFocus
2060            End If
2070          End If
2080        End If
2090      End If
2100    End With

EXITP:
2110    Set rst = Nothing
2120    Set qdf = Nothing
2130    Set dbs = Nothing
2140    Exit Sub

ERRH:
2150    Select Case ERR.Number
        Case Else
2160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2170    End Select
2180    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
2270        With Me
2280          intRetVal = 0
2290          .cmdCancel.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab
2350        With Me
2360          intRetVal = 0
2370          .revcode_SORTORDER.SetFocus
2380        End With
2390      End Select
2400    End If

EXITP:
2410    KeyCode = intRetVal
2420    Exit Sub

ERRH:
2430    intRetVal = 0
2440    THAT_PROC = THIS_PROC
2450    That_Erl = Erl
2460    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2470    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2510    If Me.NewRecord Then
2520      DoCmd.RunCommand acCmdUndo
2530    Else
2540      DoCmd.SetWarnings False
2550      DoCmd.RunCommand acCmdDeleteRecord
2560      DoCmd.SetWarnings True
2570    End If
2580    DoCmd.Close acForm, THIS_NAME

EXITP:
2590    Exit Sub

ERRH:
2600    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2610    Case 3031, 3101
2620      DoCmd.RunCommand acCmdUndo
2630    Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
2790          .opgIncomeExpense.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab
2850        With Me
2860          intRetVal = 0
2870          .cmdOK.SetFocus
2880        End With
2890      End Select
2900    End If

EXITP:
2910    KeyCode = intRetVal
2920    Exit Sub

ERRH:
2930    intRetVal = 0
2940    THAT_PROC = THIS_PROC
2950    That_Erl = Erl
2960    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2970    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3010    With Me
3020      If IsNull(.revcode_ID) = True Then
3030        .revcode_ID = lngRevcodeID
3040      End If
3050      DoCmd.RunCommand acCmdSelectRecord
3060      DoCmd.RunCommand acCmdSaveRecord
3070    End With

EXITP:
3080    Exit Sub

ERRH:
3090    Select Case ERR.Number
        Case Else
3100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3110    End Select
3120    Resume EXITP

End Sub

Private Sub opgIncomeExpense_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_AfterUpdate"

        Dim lngSortOrd As Long

3210    With Me
3220      DoCmd.Hourglass True
3230      DoEvents
3240      Select Case .opgIncomeExpense
          Case .opgIncomeExpense_optIncome.OptionValue
3250        .opgIncomeExpense_optIncome_lbl.FontBold = True
3260        .opgIncomeExpense_optIncome_box.Visible = True
3270        .opgIncomeExpense_optExpense_lbl.FontBold = False
3280        .opgIncomeExpense_optExpense_box.Visible = False
3290        .revcode_TYPE_Description = "Income"
3300        .Caption = "Add New Income Code"
3310        .Header_lbl_Income.Visible = True
3320        .Header_lbl_Expense.Visible = False
3330        lngSortOrd = AvailSortOrd(REVTYP_INC)  ' ** Function: Below.
3340        .revcode_SORTORDER_I = lngSortOrd
3350        .revcode_SORTORDER_E = Null
3360      Case .opgIncomeExpense_optExpense.OptionValue
3370        .opgIncomeExpense_optIncome_lbl.FontBold = False
3380        .opgIncomeExpense_optIncome_box.Visible = False
3390        .opgIncomeExpense_optExpense_lbl.FontBold = True
3400        .opgIncomeExpense_optExpense_box.Visible = True
3410        .revcode_TYPE_Description = "Expense"
3420        .Caption = "Add New Expense Code"
3430        .Header_lbl_Income.Visible = False
3440        .Header_lbl_Expense.Visible = True
3450        lngSortOrd = AvailSortOrd(REVTYP_EXP)  ' ** Function: Below.
3460        .revcode_SORTORDER_I = Null
3470        .revcode_SORTORDER_E = lngSortOrd
3480      End Select
3490      .revcode_SORTORDER = lngSortOrd
3500      DoEvents
3510      DoCmd.Hourglass False
3520    End With

EXITP:
3530    Exit Sub

ERRH:
3540    DoCmd.Hourglass False
3550    THAT_PROC = THIS_PROC
3560    That_Erl = Erl
3570    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3580    Resume EXITP

End Sub

Private Sub opgIncomeExpense_optExpense_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_optExpense_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
3670        With Me
3680          intRetVal = 0
3690          .revcode_DESC.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3750        With Me
3760          intRetVal = 0
3770          .cmdCancel.SetFocus
3780        End With
3790      End Select
3800    End If

EXITP:
3810    KeyCode = intRetVal
3820    Exit Sub

ERRH:
3830    intRetVal = 0
3840    THAT_PROC = THIS_PROC
3850    That_Erl = Erl
3860    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3870    Resume EXITP

End Sub

Private Sub opgIncomeExpense_optIncome_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgIncomeExpense_optIncome_KeyDown"

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
3990          .revcode_DESC.SetFocus
4000        End With
4010      End Select
4020    End If

        ' ** Shift keys.
4030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4050        With Me
4060          intRetVal = 0
4070          .cmdCancel.SetFocus
4080        End With
4090      End Select
4100    End If

EXITP:
4110    KeyCode = intRetVal
4120    Exit Sub

ERRH:
4130    intRetVal = 0
4140    THAT_PROC = THIS_PROC
4150    That_Erl = Erl
4160    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4170    Resume EXITP

End Sub

Private Sub revcode_SORTORDER_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "revcode_SORTORDER_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4220    intShiftDown = (Shift And acShiftMask) > 0
4230    intAltDown = (Shift And acAltMask) > 0
4240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4270        With Me
4280          intRetVal = 0
4290          .cmdOK.SetFocus
4300        End With
4310      End Select
4320    End If

        ' ** Shift keys.
4330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4350        With Me
4360          intRetVal = 0
4370          .revcode_ACTIVE.SetFocus
4380        End With
4390      End Select
4400    End If

EXITP:
4410    KeyCode = intRetVal
4420    Exit Sub

ERRH:
4430    intRetVal = 0
4440    THAT_PROC = THIS_PROC
4450    That_Erl = Erl
4460    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4470    Resume EXITP

End Sub

Private Function AvailSortOrd(lngRevcodeType As Long) As Long

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "AvailSortOrd"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngLastSortOrd As Long, lngMaxSortOrd As Long
        Dim lngRecs As Long
        Dim lngX As Long
        Dim lngRetVal As Long

4510    lngRetVal = 0&

4520    Set dbs = CurrentDb
4530    With dbs
          ' ** tmpRevCodeEdit, by specified [revtyp].
4540      Set qdf = .QueryDefs("qryRevCodes_14")
4550      With qdf.Parameters
4560        ![revtyp] = lngRevcodeType
4570      End With
4580      Set rst = qdf.OpenRecordset
4590      With rst
4600        If .BOF = True And .EOF = True Then
              ' ** Something's terribly wrong!
4610          lngRetVal = 98&
4620        Else
4630          .MoveLast
4640          lngRecs = .RecordCount
4650          .MoveFirst
4660          lngLastSortOrd = 0&: lngMaxSortOrd = 0&
4670          For lngX = 1& To lngRecs  ' ** They all should be in revcode_SORTORDER!
4680            If ![revcode_SORTORDER] > lngMaxSortOrd Then lngMaxSortOrd = ![revcode_SORTORDER]
4690            If (![revcode_SORTORDER] > (lngLastSortOrd + 1&)) And (lngRetVal = 0&) Then
                  ' ** Take the first available hole, but keep going to get the max.
4700              lngRetVal = lngLastSortOrd + 1&
4710            End If
4720            lngLastSortOrd = ![revcode_SORTORDER]
4730            If lngX < lngRecs Then .MoveNext
4740          Next
4750          If lngRetVal > 0& Then
                ' ** There should never be a hole down here, but just in case.
4760            If lngRetVal <= 2& Then lngRetVal = 0&
4770          End If
4780          If lngRetVal = 0& Then
4790            If lngMaxSortOrd >= 2& Then
4800              lngRetVal = lngMaxSortOrd + 1&
4810            Else
4820              lngRetVal = 3&
4830            End If
4840          End If
4850        End If
4860        .Close
4870      End With
4880      .Close
4890    End With

EXITP:
4900    Set rst = Nothing
4910    Set qdf = Nothing
4920    Set dbs = Nothing
4930    AvailSortOrd = lngRetVal
4940    Exit Function

ERRH:
4950    THAT_PROC = THIS_PROC
4960    That_Erl = Erl
4970    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4980    Resume EXITP

End Function
