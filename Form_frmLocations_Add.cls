VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLocations_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLocations_Add"

'VGC 09/20/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Other:            H {opgLoc_optOther}
' **   Name:             N {Loc_Name}
' **   OK:               O {cmdOK}
' **   USA:              U {opgLoc_optUSA}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private lngLocID As Long, blnIsOpen As Boolean
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

170           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

180           If IsNull(.Loc_State_Pref) = False Then
190             .Loc_State.DefaultValue = Chr(34) & .Loc_State_Pref & Chr(34)
200           End If

210           .opgLoc.Enabled = True
220           opgLoc_AfterUpdate  ' ** Procedure: Below.

230           lngLocID = 0&
240           .Loc_Name.SetFocus

250         Else
260           Cancel = -1
270         End If
280       Else
290         Cancel = -1
300       End If
310     End With

320     If Cancel = -1 Then
330       Beep
340       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
350     End If

EXITP:
360     Exit Sub

ERRH:
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Other:            H {opgLoc_optOther}
        ' **   USA:              U {opgLoc_optUSA}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         With Me
580           intRetVal = 0
590           If .Dirty = True Then
600             .Undo
610           End If
620         End With
630         DoCmd.Close acForm, THIS_NAME
640       End Select
650     End If

        ' ** Alt keys.
660     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
670       Select Case intRetVal
          Case vbKeyH
680         With Me
690           intRetVal = 0
700           If .opgLoc.Enabled = True Then
710             .opgLoc.SetFocus
720             .opgLoc = .opgLoc_optOther.OptionValue
730             opgLoc_AfterUpdate  ' ** Procedure: Below.
740           End If
750         End With
760       Case vbKeyU
770         With Me
780           intRetVal = 0
790           If .opgLoc.Enabled = True Then
800             .opgLoc.SetFocus
810             .opgLoc = .opgLoc_optUSA.OptionValue
820             opgLoc_AfterUpdate  ' ** Procedure: Below.
830           End If
840         End With
850       End Select
860     End If

        ' ** Ctrl keys.
870     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
880       Select Case intRetVal
          Case vbKeyS
890         intRetVal = 0
900         cmdSave_Click  ' ** Procedure: Below.
910       End Select
920     End If

        ' ** Ctl-Shift keys.
930     If intCtrlDown And (Not intAltDown) And intShiftDown Then
940       Select Case intRetVal
          Case vbKeyF
950         With Me
960           intRetVal = 0
970           .FocusHolder.SetFocus
980         End With
990       End Select
1000    End If

EXITP:
1010    KeyCode = intRetVal
1020    Exit Sub

ERRH:
1030    intRetVal = 0
1040    THAT_PROC = THIS_PROC
1050    That_Erl = Erl
1060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1070    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1110    If DataErr <> 0 Then
1120      Select Case DataErr
          Case 2046, 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do Nothing.
1130      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
1140        Select Case Me.ActiveControl.Name
            Case "Loc_State"
1150          MsgBox "Must enter a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
1160        Case "Loc_Zip"
1170          MsgBox "Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
1180        Case "Loc_Phone"
1190          MsgBox "Must enter a valid phone number.", vbInformation + vbOKOnly, "Invalid Entry"
1200        End Select
1210        Response = acDataErrContinue
1220      Case 3162
1230        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
1240        DoCmd.CancelEvent
1250      Case Else
1260        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1270        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1280      End Select
1290    End If

EXITP:
1300    Exit Sub

ERRH:
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1410    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1420    Select Case strCallingForm
        Case "frmLocations"
1430      gblnSetFocus = True
1440      gstrReturningForm = THIS_NAME
1450      DoCmd.SelectObject acForm, strCallingForm, False
1460      Forms(strCallingForm).TimerInterval = 100&
1470    End Select

EXITP:
1480    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

1610    With Me
1620      If IsNull(.Loc_Name) = True Then
1630        MsgBox "Location name must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1640      Else
1650        If .Loc_Name = vbNullString Then
1660          MsgBox "Location name must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1670        Else
1680          .Loc_Name.SetFocus
1690          If InStr(.Loc_Name.text, Chr(34)) > 0 Then
1700            MsgBox "Cannot have quote marks in Name.", vbInformation + vbOKOnly, "Invalid Characters"
1710            .cmdOK.SetFocus
1720            DoCmd.CancelEvent
1730          Else
                ' ** If they've flipped around, only take the final choice.
1740            Select Case .opgLoc
                Case .opgLoc_optUSA.OptionValue
1750              If IsNull(.Loc_Country) = False Then
1760                .Loc_Country = Null
1770              End If
1780              If IsNull(.Loc_PostalCode) = False Then
1790                .Loc_PostalCode = Null
1800              End If
1810            Case .opgLoc_optOther.OptionValue
1820              If IsNull(.Loc_State) = False Then
1830                .Loc_State = Null
1840              End If
1850              If IsNull(.Loc_Zip) = False Then
1860                .Loc_Zip = Null
1870              End If
1880            End Select
1890            lngLocID = .[Location_ID]
1900            Set dbs = CurrentDb
                ' ** Location, by specified [nam].
1910            Set qdf = dbs.QueryDefs("qryLocation_05")
1920            With qdf.Parameters
1930              ![nam] = Me.Loc_Name
1940            End With
1950            Set rst = qdf.OpenRecordset
1960            If rst.BOF = True And rst.EOF = True Then
                  ' ** All's well.
1970              rst.Close
1980              dbs.Close
1990              gdblCrtRpt_CostTot = lngLocID  ' ** Borrowing this Public variable.
2000              gblnSetFocus = True
2010              Forms("frmLocations").TimerInterval = 500&
2020              DoCmd.Close acForm, THIS_NAME
2030            Else
2040              rst.Close
2050              dbs.Close
2060              MsgBox "A Location by this name already exists.", vbInformation + vbOKOnly, "Duplicate Location"
2070              .Loc_Name.SetFocus
2080            End If
2090          End If
2100        End If
2110      End If

2120    End With

EXITP:
2130    Set rst = Nothing
2140    Set qdf = Nothing
2150    Set dbs = Nothing
2160    Exit Sub

ERRH:
2170    THAT_PROC = THIS_PROC
2180    That_Erl = Erl
2190    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2200    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2320    intShiftDown = (Shift And acShiftMask) > 0
2330    intAltDown = (Shift And acAltMask) > 0
2340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyTab
2370        With Me
2380          intRetVal = 0
2390          .cmdCancel.SetFocus
2400        End With
2410      End Select
2420    End If

        ' ** Shift keys.
2430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2440      Select Case intRetVal
          Case vbKeyTab
2450        With Me
2460          intRetVal = 0
2470          .Loc_Phone.SetFocus
2480        End With
2490      End Select
2500    End If

EXITP:
2510    KeyCode = intRetVal
2520    Exit Sub

ERRH:
2530    intRetVal = 0
2540    THAT_PROC = THIS_PROC
2550    That_Erl = Erl
2560    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2570    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2610    With Me
2620      If .NewRecord Then
2630  On Error Resume Next  ' ** If it's unavailable, e.g., new record with nothing in it, just continue to close.
2640        DoCmd.RunCommand acCmdUndo
2650  On Error GoTo ERRH
2660      Else
2670        DoCmd.RunCommand acCmdSelectRecord
2680        DoCmd.RunCommand acCmdDeleteRecord
2690      End If
2700      DoCmd.Close acForm, THIS_NAME
2710    End With

EXITP:
2720    Exit Sub

ERRH:
2730    THAT_PROC = THIS_PROC
2740    That_Erl = Erl
2750    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2760    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab
2870        With Me
2880          intRetVal = 0
2890          .Loc_Name.SetFocus
2900        End With
2910      End Select
2920    End If

        ' ** Shift keys.
2930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2940      Select Case intRetVal
          Case vbKeyTab
2950        With Me
2960          intRetVal = 0
2970          .cmdOK.SetFocus
2980        End With
2990      End Select
3000    End If

EXITP:
3010    KeyCode = intRetVal
3020    Exit Sub

ERRH:
3030    intRetVal = 0
3040    THAT_PROC = THIS_PROC
3050    That_Erl = Erl
3060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3070    Resume EXITP

End Sub

Private Sub cmdSave_Click()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3110    DoCmd.RunCommand acCmdSelectRecord
3120    DoCmd.RunCommand acCmdSaveRecord

EXITP:
3130    Exit Sub

ERRH:
3140    Select Case ERR.Number
        Case Else
3150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3160    End Select
3170    Resume EXITP

End Sub

Private Sub Loc_Name_AfterUpdate()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_AfterUpdate"

        Dim strTmp01 As String

3210    With Me
3220      If IsNull(.Loc_Name) = False Then
3230        strTmp01 = FixQuotes(.Loc_Name)  ' ** Module Function: modStringFuncs.
3240        If Compare_StringA_StringB(.Loc_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
3250          .Loc_Name = strTmp01
3260        End If
3270      End If
3280    End With

EXITP:
3290    Exit Sub

ERRH:
3300    Select Case ERR.Number
        Case Else
3310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3320    End Select
3330    Resume EXITP

End Sub

Private Sub Loc_Name_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
3470        With Me
3480          intRetVal = 0
3490          .Loc_Address1.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3550        With Me
3560          intRetVal = 0
3570          .cmdCancel.SetFocus
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

Private Sub Loc_Address1_AfterUpdate()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Address1_AfterUpdate"

        Dim strTmp01 As String

3710    With Me
3720      If IsNull(.Loc_Address1) = False Then
3730        strTmp01 = FixQuotes(.Loc_Address1)  ' ** Module Function: modStringFuncs.
3740        If Compare_StringA_StringB(.Loc_Address1, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
3750          .Loc_Address1 = strTmp01
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

Private Sub Loc_Address2_AfterUpdate()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Address2_AfterUpdate"

        Dim strTmp01 As String

3910    With Me
3920      If IsNull(.Loc_Address2) = False Then
3930        strTmp01 = FixQuotes(.Loc_Address2)  ' ** Module Function: modStringFuncs.
3940        If Compare_StringA_StringB(.Loc_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
3950          .Loc_Address2 = strTmp01
3960        End If
3970      End If
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

Private Sub Loc_City_AfterUpdate()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_City_AfterUpdate"

        Dim strTmp01 As String

4110    With Me
4120      If IsNull(.Loc_City) = False Then
4130        strTmp01 = FixQuotes(.Loc_City)  ' ** Module Function: modStringFuncs.
4140        If Compare_StringA_StringB(.Loc_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
4150          .Loc_City = strTmp01
4160        End If
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

Private Sub opgLoc_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_AfterUpdate"

4310    With Me
4320      Select Case .opgLoc
          Case .opgLoc_optUSA.OptionValue
4330        .Loc_State.Enabled = True
4340        .Loc_State.BackStyle = acBackStyleNormal
4350        .Loc_State.BorderColor = CLR_LTBLU2
4360        .Loc_Zip.Enabled = True
4370        .Loc_Zip.BackStyle = acBackStyleNormal
4380        .Loc_Zip.BorderColor = CLR_LTBLU2
4390        .Loc_Country.Enabled = False
4400        .Loc_Country.BackStyle = acBackStyleTransparent
4410        .Loc_Country.BorderColor = WIN_CLR_DISR
4420        .Loc_PostalCode.Enabled = False
4430        .Loc_PostalCode.BackStyle = acBackStyleTransparent
4440        .Loc_PostalCode.BorderColor = WIN_CLR_DISR
4450        Select Case blnIsOpen
            Case True
4460          blnIsOpen = False
4470        Case False
4480          .Loc_State.SetFocus
4490        End Select
4500      Case .opgLoc_optOther.OptionValue
4510        .Loc_Country.Enabled = True
4520        .Loc_Country.BackStyle = acBackStyleNormal
4530        .Loc_Country.BorderColor = CLR_LTBLU2
4540        .Loc_PostalCode.Enabled = True
4550        .Loc_PostalCode.BackStyle = acBackStyleNormal
4560        .Loc_PostalCode.BorderColor = CLR_LTBLU2
4570        .Loc_State.Enabled = False
4580        .Loc_State.BackStyle = acBackStyleTransparent
4590        .Loc_State.BorderColor = WIN_CLR_DISR
4600        .Loc_Zip.Enabled = False
4610        .Loc_Zip.BackStyle = acBackStyleTransparent
4620        .Loc_Zip.BorderColor = WIN_CLR_DISR
4630        Select Case blnIsOpen
            Case True
4640          blnIsOpen = False
4650        Case False
4660          .Loc_Country.SetFocus
4670        End Select
4680      End Select
4690    End With

EXITP:
4700    Exit Sub

ERRH:
4710    Select Case ERR.Number
        Case Else
4720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4730    End Select
4740    Resume EXITP

End Sub

Private Sub opgLoc_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optUSA_KeyDown"

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
4890          If .Loc_State.Enabled = True Then
4900            .Loc_State.SetFocus
4910          Else
4920            .Loc_Country.SetFocus
4930          End If
4940        End With
4950      End Select
4960    End If

        ' ** Shift keys.
4970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4980      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4990        With Me
5000          intRetVal = 0
5010          .Loc_City.SetFocus
5020        End With
5030      End Select
5040    End If

EXITP:
5050    KeyCode = intRetVal
5060    Exit Sub

ERRH:
5070    intRetVal = 0
5080    THAT_PROC = THIS_PROC
5090    That_Erl = Erl
5100    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5110    Resume EXITP

End Sub

Private Sub opgLoc_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5270        With Me
5280          intRetVal = 0
5290          If .Loc_State.Enabled = True Then
5300            .Loc_State.SetFocus
5310          Else
5320            .Loc_Country.SetFocus
5330          End If
5340        End With
5350      End Select
5360    End If

        ' ** Shift keys.
5370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5380      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5390        With Me
5400          intRetVal = 0
5410          .Loc_City.SetFocus
5420        End With
5430      End Select
5440    End If

EXITP:
5450    KeyCode = intRetVal
5460    Exit Sub

ERRH:
5470    intRetVal = 0
5480    THAT_PROC = THIS_PROC
5490    That_Erl = Erl
5500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5510    Resume EXITP

End Sub

Private Sub Loc_State_Enter()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_Enter"

5610    With Me
5620      If IsNull(.Loc_State) = True Then
5630        ExpandCombo .Loc_State  ' ** Module Procedure: modUtilities.
5640      End If
5650    End With

EXITP:
5660    Exit Sub

ERRH:
5670    Select Case ERR.Number
        Case Else
5680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5690    End Select
5700    Resume EXITP

End Sub

Private Sub Loc_State_AfterUpdate()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_AfterUpdate"

5810    With Me
5820      If IsNull(.Loc_State) = False Then
5830        If Trim(.Loc_State) <> vbNullString Then
5840          Select Case IsNull(.Loc_State_Pref)
              Case True
5850            .Loc_State_Pref = .Loc_State
5860          Case False
5870            If .Loc_State_Pref <> .Loc_State Then
5880              .Loc_State_Pref = .Loc_State
5890            End If
5900          End Select
5910        End If
5920      End If
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

Private Sub Loc_State_NotInList(NewData As String, Response As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_NotInList"

6010    With Me
6020      .Loc_State = Null
6030      MsgBox "Must select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
6040      Response = acDataErrContinue
6050    End With

EXITP:
6060    Exit Sub

ERRH:
6070    Select Case ERR.Number
        Case Else
6080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6090    End Select
6100    Resume EXITP

End Sub

Private Sub Loc_Zip_Exit(Cancel As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Zip_Exit"

6210    With Me
6220      If Not IsNull(.Loc_Zip) Then
6230        If Right(.Loc_Zip.text, 4) <> "____" And Right(.Loc_Zip.text, 4) <> "    " And _
                (Not IsNumeric(Mid(.Loc_Zip.text, 7, 1)) Or Not IsNumeric(Mid(.Loc_Zip.text, 8, 1)) Or _
                Not IsNumeric(Mid(.Loc_Zip.text, 9, 1)) Or Not IsNumeric(Mid(.Loc_Zip.text, 10, 1))) Then
6240          MsgBox "Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
6250          DoCmd.CancelEvent
6260        End If
6270      End If
6280    End With

EXITP:
6290    Exit Sub

ERRH:
6300    Select Case ERR.Number
        Case Else
6310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6320    End Select
6330    Resume EXITP

End Sub

Private Sub Loc_Phone_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Phone_KeyDown"

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
6490          .cmdOK.SetFocus
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6550        With Me
6560          intRetVal = 0
6570          Select Case .opgLoc
              Case .opgLoc_optUSA.OptionValue
6580            .Loc_Zip.SetFocus
6590          Case .opgLoc_optOther.OptionValue
6600            .Loc_PostalCode.SetFocus
6610          End Select
6620        End With
6630      End Select
6640    End If

EXITP:
6650    KeyCode = intRetVal
6660    Exit Sub

ERRH:
6670    intRetVal = 0
6680    THAT_PROC = THIS_PROC
6690    That_Erl = Erl
6700    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6710    Resume EXITP

End Sub
