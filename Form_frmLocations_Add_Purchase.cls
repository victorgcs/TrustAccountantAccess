VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLocations_Add_Purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLocations_Add_Purchase"

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
Private blnEnableCountry As Boolean, intOpgLoc As Integer
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
570         intRetVal = 0
580         cmdCancel_Click  ' ** Procedure: Below.
590       End Select
600     End If

        ' ** Alt keys.
610     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
620       Select Case intRetVal
          Case vbKeyH
630         With Me
640           intRetVal = 0
650           If .opgLoc.Enabled = True Then
660             .opgLoc.SetFocus
670             .opgLoc = .opgLoc_optOther.OptionValue
680             opgLoc_AfterUpdate  ' ** Procedure: Below.
690           End If
700         End With
710       Case vbKeyU
720         With Me
730           intRetVal = 0
740           If .opgLoc.Enabled = True Then
750             .opgLoc.SetFocus
760             .opgLoc = .opgLoc_optUSA.OptionValue
770             opgLoc_AfterUpdate  ' ** Procedure: Below.
780           End If
790         End With
800       End Select
810     End If

        ' ** Ctrl keys.
820     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
830       Select Case intRetVal
          Case vbKeyS
840         intRetVal = 0
850         cmdSave_Click  ' ** Procedure: Below.
860       End Select
870     End If

        ' ** Ctl-Shift keys.
880     If intCtrlDown And (Not intAltDown) And intShiftDown Then
890       Select Case intRetVal
          Case vbKeyF
900         With Me
910           intRetVal = 0
920           .FocusHolder.SetFocus
930         End With
940       End Select
950     End If

EXITP:
960     KeyCode = intRetVal
970     Exit Sub

ERRH:
980     intRetVal = 0
990     THAT_PROC = THIS_PROC
1000    That_Erl = Erl
1010    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1020    Resume EXITP

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
        Case "frmJournal", "frmJournal_Columns"
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

        Dim lngLocationID As Long

1610    With Me

1620      If IsNull(.Loc_Name) = True Then
1630        MsgBox "Location name must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1640      Else
1650        If .Loc_Name = vbNullString Then
1660          MsgBox "Location name must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1670        Else

              ' ** If they've flipped around, only take the final choice.
1680          Select Case .opgLoc
              Case .opgLoc_optUSA.OptionValue
1690            If IsNull(.Loc_Country) = False Then
1700              .Loc_Country = Null
1710            End If
1720            If IsNull(.Loc_PostalCode) = False Then
1730              .Loc_PostalCode = Null
1740            End If
1750          Case .opgLoc_optOther.OptionValue
1760            If IsNull(.Loc_State) = False Then
1770              .Loc_State = Null
1780            End If
1790            If IsNull(.Loc_Zip) = False Then
1800              .Loc_Zip = Null
1810            End If
1820          End Select

1830          DoCmd.RunCommand acCmdSelectRecord
1840          DoCmd.RunCommand acCmdSaveRecord
1850          lngLocationID = .[Location_ID]

1860          Select Case strCallingForm
              Case "frmJournal"
1870            Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.cmbLocations.Requery
1880            Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.cmbLocations = lngLocationID
1890            Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.[Location_ID] = lngLocationID
1900          Case "frmJournal_Columns"
1910            gdblCrtRpt_CostTot = lngLocationID  ' ** Borrowing this variable from the Court Reports.
1920            gstrCrtRpt_NetLoss = .Loc_Name  ' ** Borrowing this variable from the Court Reports.
1930          End Select
1940          DoEvents

1950          DoCmd.Close acForm, THIS_NAME

1960          DoCmd.SelectObject acForm, strCallingForm, False
1970          Select Case strCallingForm
              Case "frmJournal"
1980            If Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchaseICash.Enabled = True Then
1990              Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchaseICash.SetFocus
2000            Else
2010              If Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchasePCash.Enabled = True Then
2020                Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchasePCash.SetFocus
2030              Else
2040                Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchaseCost.SetFocus
2050              End If
2060            End If
2070          Case "frmJournal_Columns"
                ' ** Let JC_Frm_Locations() set the focus.
2080          End Select

2090        End If
2100      End If

2110    End With

EXITP:
2120    Exit Sub

ERRH:
2130    THAT_PROC = THIS_PROC
2140    That_Erl = Erl
2150    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2160    Resume EXITP

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
2370          .Loc_Phone.SetFocus
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

2510    With Me
2520      If .NewRecord = True Then
2530  On Error Resume Next  ' ** If it's unavailable, e.g., new record with nothing in it, just continue to close.
2540        DoCmd.RunCommand acCmdUndo
2550  On Error GoTo ERRH
2560      Else
2570        DoCmd.RunCommand acCmdSelectRecord
2580        DoCmd.RunCommand acCmdDeleteRecord
2590      End If
2600      gblnMessage = False

2610      DoCmd.Close acForm, THIS_NAME

2620      DoCmd.SelectObject acForm, strCallingForm, False
2630      Select Case strCallingForm
          Case "frmJournal"
2640        Forms(strCallingForm).frmJournal_Sub3_Purchase.SetFocus
2650        If Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchaseICash.Enabled = True Then
2660          Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchaseICash.SetFocus
2670        Else
2680          If Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchasePCash.Enabled = True Then
2690            Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchasePCash.SetFocus
2700          Else
2710            Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchaseCost.SetFocus
2720          End If
2730        End If
2740      Case "frmJournal_Columns"
            ' ** Let JC_Frm_Locations() set the focus.
2750      End Select
2760    End With

EXITP:
2770    Exit Sub

ERRH:
2780    THAT_PROC = THIS_PROC
2790    That_Erl = Erl
2800    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2810    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2920    intShiftDown = (Shift And acShiftMask) > 0
2930    intAltDown = (Shift And acAltMask) > 0
2940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2960      Select Case intRetVal
          Case vbKeyTab
2970        With Me
2980          intRetVal = 0
2990          .Loc_Name.SetFocus
3000        End With
3010      End Select
3020    End If

        ' ** Shift keys.
3030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3040      Select Case intRetVal
          Case vbKeyTab
3050        With Me
3060          intRetVal = 0
3070          .cmdOK.SetFocus
3080        End With
3090      End Select
3100    End If

EXITP:
3110    KeyCode = intRetVal
3120    Exit Sub

ERRH:
3130    intRetVal = 0
3140    THAT_PROC = THIS_PROC
3150    That_Erl = Erl
3160    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3170    Resume EXITP

End Sub

Private Sub cmdSave_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3210    DoCmd.RunCommand acCmdSelectRecord
3220    DoCmd.RunCommand acCmdSaveRecord

EXITP:
3230    Exit Sub

ERRH:
3240    Select Case ERR.Number
        Case Else
3250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3260    End Select
3270    Resume EXITP

End Sub

Private Sub Loc_Name_AfterUpdate()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_AfterUpdate"

        Dim strTmp01 As String

3310    With Me
3320      If IsNull(.Loc_Name) = False Then
3330        strTmp01 = FixQuotes(.Loc_Name)  ' ** Module Function: modStringFuncs.
3340        If Compare_StringA_StringB(.Loc_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
3350          .Loc_Name = strTmp01
3360        End If
3370      End If
3380    End With

EXITP:
3390    Exit Sub

ERRH:
3400    Select Case ERR.Number
        Case Else
3410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3420    End Select
3430    Resume EXITP

End Sub

Private Sub Loc_Name_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3520    intShiftDown = (Shift And acShiftMask) > 0
3530    intAltDown = (Shift And acAltMask) > 0
3540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3570        With Me
3580          intRetVal = 0
3590          .Loc_Address1.SetFocus
3600        End With
3610      End Select
3620    End If

        ' ** Shift keys.
3630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3650        With Me
3660          intRetVal = 0
3670          .cmdCancel.SetFocus
3680        End With
3690      End Select
3700    End If

EXITP:
3710    KeyCode = intRetVal
3720    Exit Sub

ERRH:
3730    intRetVal = 0
3740    THAT_PROC = THIS_PROC
3750    That_Erl = Erl
3760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3770    Resume EXITP

End Sub

Private Sub Loc_Address1_AfterUpdate()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Address1_AfterUpdate"

        Dim strTmp01 As String

3810    With Me
3820      If IsNull(.Loc_Address1) = False Then
3830        strTmp01 = FixQuotes(.Loc_Address1)  ' ** Module Function: modStringFuncs.
3840        If Compare_StringA_StringB(.Loc_Address1, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
3850          .Loc_Address1 = strTmp01
3860        End If
3870      End If
3880    End With

EXITP:
3890    Exit Sub

ERRH:
3900    Select Case ERR.Number
        Case Else
3910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3920    End Select
3930    Resume EXITP

End Sub

Private Sub Loc_Address2_AfterUpdate()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Address2_AfterUpdate"

        Dim strTmp01 As String

4010    With Me
4020      If IsNull(.Loc_Address2) = False Then
4030        strTmp01 = FixQuotes(.Loc_Address2)  ' ** Module Function: modStringFuncs.
4040        If Compare_StringA_StringB(.Loc_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
4050          .Loc_Address2 = strTmp01
4060        End If
4070      End If
4080    End With

EXITP:
4090    Exit Sub

ERRH:
4100    Select Case ERR.Number
        Case Else
4110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4120    End Select
4130    Resume EXITP

End Sub

Private Sub Loc_City_AfterUpdate()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_City_AfterUpdate"

        Dim strTmp01 As String

4210    With Me
4220      If IsNull(.Loc_City) = False Then
4230        strTmp01 = FixQuotes(.Loc_City)  ' ** Module Function: modStringFuncs.
4240        If Compare_StringA_StringB(.Loc_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
4250          .Loc_City = strTmp01
4260        End If
4270      End If
4280    End With

EXITP:
4290    Exit Sub

ERRH:
4300    Select Case ERR.Number
        Case Else
4310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4320    End Select
4330    Resume EXITP

End Sub

Private Sub opgLoc_AfterUpdate()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_AfterUpdate"

4410    With Me
4420      Select Case .opgLoc
          Case .opgLoc_optUSA.OptionValue
4430        .Loc_State.Enabled = True
4440        .Loc_State.BackStyle = acBackStyleNormal
4450        .Loc_State.BorderColor = CLR_LTBLU2
4460        .Loc_Zip.Enabled = True
4470        .Loc_Zip.BackStyle = acBackStyleNormal
4480        .Loc_Zip.BorderColor = CLR_LTBLU2
4490        .Loc_Country.Enabled = False
4500        .Loc_Country.BackStyle = acBackStyleTransparent
4510        .Loc_Country.BorderColor = WIN_CLR_DISR
4520        .Loc_PostalCode.Enabled = False
4530        .Loc_PostalCode.BackStyle = acBackStyleTransparent
4540        .Loc_PostalCode.BorderColor = WIN_CLR_DISR
4550        Select Case blnIsOpen
            Case True
4560          blnIsOpen = False
4570        Case False
4580          .Loc_State.SetFocus
4590        End Select
4600      Case .opgLoc_optOther.OptionValue
4610        .Loc_Country.Enabled = True
4620        .Loc_Country.BackStyle = acBackStyleNormal
4630        .Loc_Country.BorderColor = CLR_LTBLU2
4640        .Loc_PostalCode.Enabled = True
4650        .Loc_PostalCode.BackStyle = acBackStyleNormal
4660        .Loc_PostalCode.BorderColor = CLR_LTBLU2
4670        .Loc_State.Enabled = False
4680        .Loc_State.BackStyle = acBackStyleTransparent
4690        .Loc_State.BorderColor = WIN_CLR_DISR
4700        .Loc_Zip.Enabled = False
4710        .Loc_Zip.BackStyle = acBackStyleTransparent
4720        .Loc_Zip.BorderColor = WIN_CLR_DISR
4730        Select Case blnIsOpen
            Case True
4740          blnIsOpen = False
4750        Case False
4760          .Loc_Country.SetFocus
4770        End Select
4780      End Select
4790    End With

EXITP:
4800    Exit Sub

ERRH:
4810    Select Case ERR.Number
        Case Else
4820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4830    End Select
4840    Resume EXITP

End Sub

Private Sub opgLoc_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optUSA_KeyDown"

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
4990          If .Loc_State.Enabled = True Then
5000            .Loc_State.SetFocus
5010          Else
5020            .Loc_Country.SetFocus
5030          End If
5040        End With
5050      End Select
5060    End If

        ' ** Shift keys.
5070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5080      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5090        With Me
5100          intRetVal = 0
5110          .Loc_City.SetFocus
5120        End With
5130      End Select
5140    End If

EXITP:
5150    KeyCode = intRetVal
5160    Exit Sub

ERRH:
5170    intRetVal = 0
5180    THAT_PROC = THIS_PROC
5190    That_Erl = Erl
5200    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5210    Resume EXITP

End Sub

Private Sub opgLoc_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5320    intShiftDown = (Shift And acShiftMask) > 0
5330    intAltDown = (Shift And acAltMask) > 0
5340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5370        With Me
5380          intRetVal = 0
5390          If .Loc_State.Enabled = True Then
5400            .Loc_State.SetFocus
5410          Else
5420            .Loc_Country.SetFocus
5430          End If
5440        End With
5450      End Select
5460    End If

        ' ** Shift keys.
5470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5480      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5490        With Me
5500          intRetVal = 0
5510          .Loc_City.SetFocus
5520        End With
5530      End Select
5540    End If

EXITP:
5550    KeyCode = intRetVal
5560    Exit Sub

ERRH:
5570    intRetVal = 0
5580    THAT_PROC = THIS_PROC
5590    That_Erl = Erl
5600    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5610    Resume EXITP

End Sub

Private Sub Loc_State_Enter()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_Enter"

5710    With Me
5720      If IsNull(.Loc_State) Then
5730        ExpandCombo .Loc_State  ' ** Module Procedure: modUtilities.
5740      End If
5750    End With

EXITP:
5760    Exit Sub

ERRH:
5770    Select Case ERR.Number
        Case Else
5780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5790    End Select
5800    Resume EXITP

End Sub

Private Sub Loc_State_AfterUpdate()

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_AfterUpdate"

5910    With Me
5920      If IsNull(.Loc_State) = False Then
5930        If Trim(.Loc_State) <> vbNullString Then
5940          Select Case IsNull(.Loc_State_Pref)
              Case True
5950            .Loc_State_Pref = .Loc_State
5960          Case False
5970            If .Loc_State_Pref <> .Loc_State Then
5980              .Loc_State_Pref = .Loc_State
5990            End If
6000          End Select
6010        End If
6020      End If
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

Private Sub Loc_State_NotInList(NewData As String, Response As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_NotInList"

6110    With Me
6120      .Loc_State = Null
6130      MsgBox "Must select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
6140      Response = acDataErrContinue
6150    End With

EXITP:
6160    Exit Sub

ERRH:
6170    Select Case ERR.Number
        Case Else
6180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6190    End Select
6200    Resume EXITP

End Sub

Private Sub Loc_Zip_Exit(Cancel As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Zip_Exit"

6310    With Me
6320      If Not IsNull(.Loc_Zip) Then
6330        If Right(.Loc_Zip.text, 4) <> "____" And Right(.Loc_Zip.text, 4) <> "    " And _
                (Not IsNumeric(Mid(.Loc_Zip.text, 7, 1)) Or Not IsNumeric(Mid(.Loc_Zip.text, 8, 1)) Or _
                Not IsNumeric(Mid(.Loc_Zip.text, 9, 1)) Or Not IsNumeric(Mid(.Loc_Zip.text, 10, 1))) Then
6340          MsgBox "Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
6350          DoCmd.CancelEvent
6360        End If
6370      End If
6380    End With

EXITP:
6390    Exit Sub

ERRH:
6400    Select Case ERR.Number
        Case Else
6410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6420    End Select
6430    Resume EXITP

End Sub

Private Sub Loc_Phone_KeyDown(KeyCode As Integer, Shift As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Phone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6520    intShiftDown = (Shift And acShiftMask) > 0
6530    intAltDown = (Shift And acAltMask) > 0
6540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6570        With Me
6580          intRetVal = 0
6590          .cmdOK.SetFocus
6600        End With
6610      End Select
6620    End If

        ' ** Shift keys.
6630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6650        With Me
6660          intRetVal = 0
6670          Select Case .opgLoc
              Case .opgLoc_optUSA.OptionValue
6680            .Loc_Zip.SetFocus
6690          Case .opgLoc_optOther.OptionValue
6700            .Loc_PostalCode.SetFocus
6710          End Select
6720        End With
6730      End Select
6740    End If

EXITP:
6750    KeyCode = intRetVal
6760    Exit Sub

ERRH:
6770    intRetVal = 0
6780    THAT_PROC = THIS_PROC
6790    That_Erl = Erl
6800    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6810    Resume EXITP

End Sub
