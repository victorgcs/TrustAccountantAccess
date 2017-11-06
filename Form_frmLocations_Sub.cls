VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLocations_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLocations_Sub"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd on frmLocations}
' **   Close:            C {cmdClose on frmLocations}
' **   Delete:           D {cmdDelete on frmLocations}
' **   Enable Country:   E {chkEnableCountry on frmLocations}
' **   Name:             N {Loc_Name}
' **   Print Report:     P {cmdPrintReport on frmLocations}
' **   Move Assets:      V {cmdMoveAssets on frmLocations}
' **   Exit:             X {cmdClose on frmLocations}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[Loc_Name], [Loc_Address1]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private lngForm_Width As Long, lngCountry_Left As Long, lngPostalCode_Left As Long
Private blnKeyLeft As Boolean, blnKeyRight As Boolean
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long, That_ErrDesc As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmLocations" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
190         blnKeyLeft = False: blnKeyRight = False

200         .form_width_line.Left = 0&  ' ** Just get it out of the way.

210         lngSortLbl_Width = .Sort_lbl.Width
220         lngForm_Width = .Width
230         lngCountry_Left = .Loc_Country.Left
240         lngPostalCode_Left = .Loc_PostalCode.Left

250       Else
260         Cancel = -1
270       End If
280     End With

290     If Cancel = -1 Then
300       Beep
310       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
320     End If

EXITP:
330     Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       .Location_ID.ForeColor = CLR_DISABLED_FG
430       .Location_ID.BackColor = CLR_DISABLED_BG
440       SortNow THIS_PROC  ' ** Procedure: Below
450     End With

EXITP:
460     Exit Sub

ERRH:
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub Form_Current()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

610     With Me
620       If .Location_ID = 1& And .Loc_Name = "{Unassigned}" Then
630         .Loc_Name.Locked = True
640         .Loc_Address1.Locked = True
650         .Loc_Address2.Locked = True
660         .Loc_City.Locked = True
670         .Loc_State.Locked = True
680         .Loc_Zip.Locked = True
690         .Loc_Phone.Locked = True
700       Else
710         .Loc_Name.Locked = False
720         .Loc_Address1.Locked = False
730         .Loc_Address2.Locked = False
740         .Loc_City.Locked = False
750         .Loc_State.Locked = False
760         .Loc_Zip.Locked = False
770         .Loc_Phone.Locked = False
780       End If
790     End With

EXITP:
800     Exit Sub

ERRH:
810     THAT_PROC = THIS_PROC
820     That_Erl = Erl
830     That_ErrDesc = ERR.description
840     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
850     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strControl As String
        Dim intRetVal As Integer

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960   On Error Resume Next
970       strControl = Screen.ActiveControl.Name
980   On Error GoTo ERRH
990       Select Case intRetVal
          Case vbKeyEscape
1000        With Me
1010          intRetVal = 0
1020          blnKeyLeft = False
1030          blnKeyRight = False
1040          .Parent.cmdClose_Click  ' ** Form Procedure: frmLocations.
1050        End With
1060      Case vbKeyUp
1070        With Me
1080          intRetVal = 0
1090          blnKeyLeft = False
1100          blnKeyRight = False
1110          Select Case strControl
              Case "Loc_Address1"
1120            If .CurrentRecord > 1 Then
1130              MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1140              .Loc_Address2.SetFocus
1150            End If
1160          Case "Loc_Address2"
1170            .Loc_Address1.SetFocus
1180          Case "Loc_City", "Loc_State"
1190            If .CurrentRecord > 1 Then
1200              MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1210              .Loc_Phone.SetFocus
1220            End If
1230          Case "Loc Phone"
1240            .Loc_City.SetFocus
1250          Case Else
1260            If .CurrentRecord > 1 Then
1270              MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1280            End If
1290          End Select
1300        End With
1310      Case vbKeyDown
1320        With Me
1330          intRetVal = 0
1340          blnKeyLeft = False
1350          blnKeyRight = False
1360          lngRecsCur = RecCnt  ' ** Function: Below.
1370          Select Case strControl
              Case "Loc_Address1"
1380            .Loc_Address2.SetFocus
1390          Case "Loc_Address2"
1400            If .CurrentRecord < lngRecsCur Then
1410              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1420              .Loc_Address1.SetFocus
1430            End If
1440          Case "Loc_City", "Loc_State"
1450            .Loc_Phone.SetFocus
1460          Case "Loc_Phone"
1470            If .CurrentRecord < lngRecsCur Then
1480              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1490              .Loc_City.SetFocus
1500            End If
1510          Case Else
1520            If .CurrentRecord < lngRecsCur Then
1530              MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1540            End If
1550          End Select
1560        End With
1570      Case vbKeyLeft
1580        Select Case strControl  ' ** Don't change intRetVal.
            Case "Loc_Phone", "Loc_City"
1590          blnKeyLeft = True
1600          blnKeyRight = False
1610        Case Else
1620          blnKeyLeft = False
1630          blnKeyRight = False
1640        End Select
1650      Case vbKeyRight
1660        Select Case strControl  ' ** Don't change intRetVal.
            Case "Loc_Address1", "Loc_Address2"
1670          blnKeyRight = True
1680          blnKeyLeft = False
1690        Case Else
1700          blnKeyRight = False
1710          blnKeyLeft = False
1720        End Select
1730      Case Else
1740        blnKeyLeft = False
1750        blnKeyRight = False
1760      End Select
1770    End If

        ' ** Alt keys.
1780    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1790      Select Case intRetVal
          Case vbKeyN
1800        With Me
1810          intRetVal = 0         ' ** Needed because labels aren't connected to
1820          .Loc_Name.SetFocus  ' ** text boxes in order to permit DblClicking.
1830        End With
1840      Case vbKeyX
1850        With Me
1860          intRetVal = 0
1870          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountAssets.
1880        End With
1890      End Select
1900    End If

        ' ** Ctrl keys.
1910    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1920      Select Case intRetVal
          Case vbKeyS
1930        intRetVal = 0
1940        cmdSave_Click  ' ** Procedure: Below.
1950      Case vbKeyTab
1960        With Me
1970          intRetVal = 0
1980          DoCmd.SelectObject acForm, .Parent.Name, False
1990          .Parent.chkEnableCountry.SetFocus
2000        End With
2010      Case vbKeyPageUp, vbKeyHome
2020        intRetVal = 0
2030        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2040      Case vbKeyPageDown, vbKeyEnd
2050        intRetVal = 0
2060        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2070      End Select
2080    End If

        ' ** Ctrl-Shift keys.
2090    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2100      Select Case intRetVal
          Case vbKeyF
2110        With Me
2120          intRetVal = 0
2130          DoCmd.SelectObject acForm, .Parent.Name, False
2140          .Parent.FocusHolder.SetFocus
2150        End With
2160      Case vbKeyTab
2170        With Me
2180          intRetVal = 0
2190          DoCmd.SelectObject acForm, .Parent.Name, False
2200          .Parent.cmdClose.SetFocus
2210        End With
2220      End Select
2230    End If

EXITP:
2240    KeyCode = intRetVal
2250    Exit Sub

ERRH:
2260    intRetVal = 0
2270    THAT_PROC = THIS_PROC
2280    That_Erl = Erl
2290    That_ErrDesc = ERR.description
2300    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2310    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

2410    With Me
2420      .Loc_Username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2430      .Loc_DateModified = Now()
2440    End With

EXITP:
2450    Exit Sub

ERRH:
2460    THAT_PROC = THIS_PROC
2470    That_Erl = Erl
2480    That_ErrDesc = ERR.description
2490    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2500    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2610    If DataErr <> 0 Then
2620      Select Case DataErr
          Case 2046, 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
2630      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2640        Select Case Me.ActiveControl.Name
            Case "Loc_State"
2650          MsgBox "Must enter a valid state code.", vbInformation + vbOKOnly, "Entry Required"
2660        Case "Loc_Zip"
2670          MsgBox "Must enter a valid zip code.", vbInformation + vbOKOnly, "Entry Required"
2680        Case "Loc_Phone"
2690          MsgBox "Must enter a valid phone number.", vbInformation + vbOKOnly, "Entry Required"
2700        End Select
2710        Response = acDataErrContinue
2720      Case 3162
2730        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
2740        DoCmd.CancelEvent
2750      Case Else
2760        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2770        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl
2780      End Select
2790    End If

EXITP:
2800    Exit Sub

ERRH:
2810    Select Case ERR.Number
        Case Else
2820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2830    End Select
2840    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2910    DoCmd.RunCommand acCmdSelectRecord
2920    DoCmd.RunCommand acCmdSaveRecord

EXITP:
2930    Exit Sub

ERRH:
2940    Select Case ERR.Number
        Case Else
2950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2960    End Select
2970    Resume EXITP

End Sub

Private Sub Loc_Name_lbl_DblClick(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_lbl_DblClick"

3010    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub Loc_Name_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3170        With Me
3180          intRetVal = 0
3190          .Loc_Address1.SetFocus
3200        End With
3210      End Select
3220    End If

        ' ** Shift keys.
3230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3250        With Me
3260          intRetVal = 0
3270          If .CurrentRecord > 1 Then
3280            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3290            .Loc_Phone.SetFocus
3300          Else
3310            DoCmd.SelectObject acForm, .Parent.Name, False
3320            .Parent.cmdClose.SetFocus
3330          End If
3340        End With
3350      End Select
3360    End If

EXITP:
3370    KeyCode = intRetVal
3380    Exit Sub

ERRH:
3390    intRetVal = 0
3400    THAT_PROC = THIS_PROC
3410    That_Erl = Erl
3420    That_ErrDesc = ERR.description
3430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3440    Resume EXITP

End Sub

Private Sub Loc_Name_GotFocus()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Name_GotFocus"

3510    With Me
3520      .Loc_Name.SelLength = 0
3530      .Loc_Name.SelStart = 99
3540    End With

EXITP:
3550    Exit Sub

ERRH:
3560    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
3570    Case Else
3580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3590    End Select
3600    Resume EXITP

End Sub

Private Sub Loc_Address1_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Address1_lbl_DblClick"

3710    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub Loc_Address1_Exit(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Address1_Exit"

3810    With Me
3820      If blnKeyRight = True Then
3830        blnKeyRight = False
3840        .Loc_City.SetFocus
3850      End If
3860    End With

EXITP:
3870    Exit Sub

ERRH:
3880    Select Case ERR.Number
        Case Else
3890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3900    End Select
3910    Resume EXITP

End Sub

Private Sub Loc_Address2_Exit(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Address2_Exit"

4010    With Me
4020      If blnKeyRight = True Then
4030        blnKeyRight = False
4040        .Loc_Phone.SetFocus
4050      End If
4060    End With

EXITP:
4070    Exit Sub

ERRH:
4080    Select Case ERR.Number
        Case Else
4090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4100    End Select
4110    Resume EXITP

End Sub

Private Sub Loc_City_lbl_DblClick(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_City_lbl_DblClick"

4210    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub Loc_City_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_City_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4320    intShiftDown = (Shift And acShiftMask) > 0
4330    intAltDown = (Shift And acAltMask) > 0
4340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4370        With Me
4380          intRetVal = 0
4390          Select Case .Parent.chkEnableCountry
              Case True
4400            Select Case IsNull(.Loc_Country)
                Case True
4410              .Loc_State.SetFocus
4420            Case False
4430              .Loc_Country.SetFocus
4440            End Select
4450          Case False
4460            .Loc_State.SetFocus
4470          End Select
4480        End With
4490      End Select
4500    End If

        ' ** Shift keys.
4510    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4520      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4530        With Me
4540          intRetVal = 0
4550          .Loc_Address1.SetFocus
4560        End With
4570      End Select
4580    End If

EXITP:
4590    KeyCode = intRetVal
4600    Exit Sub

ERRH:
4610    intRetVal = 0
4620    THAT_PROC = THIS_PROC
4630    That_Erl = Erl
4640    That_ErrDesc = ERR.description
4650    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4660    Resume EXITP

End Sub

Private Sub Loc_City_Exit(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_City_Exit"

4710    With Me
4720      If blnKeyLeft = True Then
4730        blnKeyLeft = False
4740        .Loc_Address1.SetFocus
4750      End If
4760    End With

EXITP:
4770    Exit Sub

ERRH:
4780    Select Case ERR.Number
        Case Else
4790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4800    End Select
4810    Resume EXITP

End Sub

Private Sub Loc_State_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_lbl_DblClick"

4910    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub Loc_State_Enter()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_Enter"

5010    With Me
5020      If IsNull(.Loc_State) = True Then
5030        ExpandCombo .Loc_State  ' ** Module Procedure: modUtilities.
5040      End If
5050    End With

EXITP:
5060    Exit Sub

ERRH:
5070    THAT_PROC = THIS_PROC
5080    That_Erl = Erl
5090    That_ErrDesc = ERR.description
5100    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5110    Resume EXITP

End Sub

Private Sub Loc_State_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_AfterUpdate"

5210    With Me
5220      If IsNull(.Loc_State) = False Then
5230        If Trim(.Loc_State) <> vbNullString Then
5240          Select Case IsNull(.Parent.Loc_State_Pref)
              Case True
5250            .Parent.Loc_State_Pref = .Loc_State
5260            .Loc_State.DefaultValue = Chr(34) & .Loc_State & Chr(34)
5270          Case False
5280            If .Parent.Loc_State_Pref <> .Loc_State Then
5290              .Parent.Loc_State_Pref = .Loc_State
5300              .Loc_State.DefaultValue = Chr(34) & .Loc_State & Chr(34)
5310            End If
5320          End Select
5330        End If
5340      End If
5350    End With

EXITP:
5360    Exit Sub

ERRH:
5370    Select Case ERR.Number
        Case Else
5380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5390    End Select
5400    Resume EXITP

End Sub

Private Sub Loc_State_KeyDown(KeyCode As Integer, Shift As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_KeyDown"

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
5590          .Loc_Zip.SetFocus
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5650        With Me
5660          intRetVal = 0
5670          .Loc_City.SetFocus
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
5760    That_ErrDesc = ERR.description
5770    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5780    Resume EXITP

End Sub

Private Sub Loc_State_NotInList(NewData As String, Response As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_State_NotInList"

5810    With Me
5820      .Loc_State = Null
5830      MsgBox "Must select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
5840      Response = acDataErrContinue
5850    End With

EXITP:
5860    Exit Sub

ERRH:
5870    THAT_PROC = THIS_PROC
5880    That_Erl = Erl
5890    That_ErrDesc = ERR.description
5900    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5910    Resume EXITP

End Sub

Private Sub Loc_Zip_lbl_DblClick(Cancel As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Zip_lbl_DblClick"

6010    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
6020    Exit Sub

ERRH:
6030    Select Case ERR.Number
        Case Else
6040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6050    End Select
6060    Resume EXITP

End Sub

Private Sub Loc_Zip_KeyDown(KeyCode As Integer, Shift As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Zip_KeyDown"

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
6190          Select Case .Parent.chkEnableCountry
              Case True
6200            Select Case IsNull(.Loc_State)
                Case True
6210              .Loc_Country.SetFocus
6220            Case False
6230              .Loc_Phone.SetFocus
6240            End Select
6250          Case False
6260            .Loc_Phone.SetFocus
6270          End Select
6280        End With
6290      End Select
6300    End If

        ' ** Shift keys.
6310    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6330        With Me
6340          intRetVal = 0
6350          .Loc_State.SetFocus
6360        End With
6370      End Select
6380    End If

EXITP:
6390    KeyCode = intRetVal
6400    Exit Sub

ERRH:
6410    intRetVal = 0
6420    THAT_PROC = THIS_PROC
6430    That_Erl = Erl
6440    That_ErrDesc = ERR.description
6450    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6460    Resume EXITP

End Sub

Private Sub Loc_Zip_Exit(Cancel As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Zip_Exit"

6510    With Me
6520      If Not IsNull(.Loc_Zip) Then
6530        If Right(.Loc_Zip.text, 4) <> "____" And Right(.Loc_Zip.text, 4) <> "    " And _
                (Not IsNumeric(Mid(.Loc_Zip.text, 7, 1)) Or Not IsNumeric(Mid(.Loc_Zip.text, 8, 1)) Or _
                Not IsNumeric(Mid(.Loc_Zip.text, 9, 1)) Or Not IsNumeric(Mid(.Loc_Zip.text, 10, 1))) Then
6540          MsgBox "Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
6550          DoCmd.CancelEvent
6560        End If
6570      End If
6580    End With

EXITP:
6590    Exit Sub

ERRH:
6600    Select Case ERR.Number
        Case Else
6610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6620    End Select
6630    Resume EXITP

End Sub

Private Sub Loc_Country_lbl_DblClick(Cancel As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Country_lbl_DblClick"

6710    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
6720    Exit Sub

ERRH:
6730    Select Case ERR.Number
        Case Else
6740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6750    End Select
6760    Resume EXITP

End Sub

Private Sub Loc_Country_KeyDown(KeyCode As Integer, Shift As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Country_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6820    intShiftDown = (Shift And acShiftMask) > 0
6830    intAltDown = (Shift And acAltMask) > 0
6840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6870        With Me
6880          intRetVal = 0
6890          .Loc_PostalCode.SetFocus
6900        End With
6910      End Select
6920    End If

        ' ** Shift keys.
6930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6950        With Me
6960          intRetVal = 0
6970          Select Case IsNull(.Loc_Country)
              Case True
6980            .Loc_Zip.SetFocus
6990          Case False
7000            .Loc_City.SetFocus
7010          End Select
7020        End With
7030      End Select
7040    End If

EXITP:
7050    KeyCode = intRetVal
7060    Exit Sub

ERRH:
7070    intRetVal = 0
7080    THAT_PROC = THIS_PROC
7090    That_Erl = Erl
7100    That_ErrDesc = ERR.description
7110    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7120    Resume EXITP

End Sub

Private Sub Loc_PostalCode_lbl_DblClick(Cancel As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_PostalCode_lbl_DblClick"

7210    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
7220    Exit Sub

ERRH:
7230    Select Case ERR.Number
        Case Else
7240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7250    End Select
7260    Resume EXITP

End Sub

Private Sub Loc_PostalCode_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_PostalCode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7320    intShiftDown = (Shift And acShiftMask) > 0
7330    intAltDown = (Shift And acAltMask) > 0
7340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7370        With Me
7380          intRetVal = 0
7390          .Loc_Phone.SetFocus
7400        End With
7410      End Select
7420    End If

        ' ** Shift keys.
7430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7450        With Me
7460          intRetVal = 0
7470          .Loc_Country.SetFocus
7480        End With
7490      End Select
7500    End If

EXITP:
7510    KeyCode = intRetVal
7520    Exit Sub

ERRH:
7530    intRetVal = 0
7540    THAT_PROC = THIS_PROC
7550    That_Erl = Erl
7560    That_ErrDesc = ERR.description
7570    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7580    Resume EXITP

End Sub

Private Sub Loc_Phone_lbl_DblClick(Cancel As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Phone_lbl_DblClick"

7610    SortNow THIS_PROC  ' ** Procedure: Below

EXITP:
7620    Exit Sub

ERRH:
7630    Select Case ERR.Number
        Case Else
7640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7650    End Select
7660    Resume EXITP

End Sub

Private Sub Loc_Phone_KeyDown(KeyCode As Integer, Shift As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Phone_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7720    intShiftDown = (Shift And acShiftMask) > 0
7730    intAltDown = (Shift And acAltMask) > 0
7740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7770        With Me
7780          intRetVal = 0
7790          lngRecsCur = RecCnt  ' ** Function: Below.
7800          If .CurrentRecord < lngRecsCur Then
7810            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
7820            .Loc_Name.SetFocus
7830          Else
7840            DoCmd.SelectObject acForm, .Parent.Name, False
7850            .Parent.chkEnableCountry.SetFocus
7860          End If
7870        End With
7880      End Select
7890    End If

        ' ** Shift keys.
7900    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7910      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7920        With Me
7930          intRetVal = 0
7940          Select Case .Parent.chkEnableCountry
              Case True
7950            Select Case IsNull(.Loc_State)
                Case True
7960              .Loc_PostalCode.SetFocus
7970            Case False
7980              .Loc_Zip.SetFocus
7990            End Select
8000          Case False
8010            .Loc_Zip.SetFocus
8020          End Select
8030        End With
8040      End Select
8050    End If

EXITP:
8060    KeyCode = intRetVal
8070    Exit Sub

ERRH:
8080    intRetVal = 0
8090    THAT_PROC = THIS_PROC
8100    That_Erl = Erl
8110    That_ErrDesc = ERR.description
8120    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8130    Resume EXITP

End Sub

Private Sub Loc_Phone_Exit(Cancel As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "Loc_Phone_Exit"

8210    With Me
8220      If blnKeyLeft = True Then
8230        blnKeyLeft = False
8240        .Loc_Address2.SetFocus
8250      End If
8260    End With

EXITP:
8270    Exit Sub

ERRH:
8280    Select Case ERR.Number
        Case Else
8290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8300    End Select
8310    Resume EXITP

End Sub

Public Sub DeleteRec(lngLocID As Long)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "DeleteRec"

8410    With Me
8420      DoCmd.SelectObject acForm, .Parent.Name, False
8430      Forms(.Parent.Name).frmLocations_Sub.SetFocus
8440      If .Location_ID = lngLocID Then
8450        If .NewRecord Then
8460          MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Above.
8470        Else
8480          .AllowDeletions = True
8490          DoCmd.RunCommand acCmdDeleteRecord
8500          .AllowDeletions = False
8510        End If
8520      End If
8530    End With

EXITP:
8540    Exit Sub

ERRH:
8550    Select Case ERR.Number
        Case Else
8560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8570    End Select
8580    Resume EXITP

End Sub

Public Sub EnableCountry(blnEnable As Boolean)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "EnableCountry"

8610    With Me
8620      Select Case blnEnable
          Case True
8630        .Width = lngForm_Width
8640        .Loc_Country.Left = lngCountry_Left
8650        .Loc_Country_lbl.Left = lngCountry_Left
8660        .Loc_Country_lbl_line.Left = lngCountry_Left
8670        .Loc_Country.Enabled = True
8680        .Loc_Country.Visible = True
8690        .Loc_Country_lbl.Visible = True
8700        .Loc_Country_lbl_line.Visible = True
8710        .Loc_PostalCode.Left = lngPostalCode_Left
8720        .Loc_PostalCode_lbl.Left = lngPostalCode_Left
8730        .Loc_PostalCode_lbl_line.Left = lngPostalCode_Left
8740        .Loc_PostalCode.Enabled = True
8750        .Loc_PostalCode.Visible = True
8760        .Loc_PostalCode_lbl.Visible = True
8770        .Loc_PostalCode_lbl_line.Visible = True
8780      Case False
8790        .FocusHolder2.SetFocus
8800        .Loc_PostalCode.Visible = False
8810        .Loc_PostalCode.Enabled = False
8820        .Loc_PostalCode.Left = 0&
8830        .Loc_PostalCode_lbl.Visible = False
8840        .Loc_PostalCode_lbl.Left = 0&
8850        .Loc_PostalCode_lbl_line.Visible = False
8860        .Loc_PostalCode_lbl_line.Left = 0&
8870        .Loc_Country.Visible = False
8880        .Loc_Country.Enabled = False
8890        .Loc_Country.Left = 0&
8900        .Loc_Country_lbl.Visible = False
8910        .Loc_Country_lbl.Left = 0&
8920        .Loc_Country_lbl_line.Visible = False
8930        .Loc_Country_lbl_line.Left = 0&
8940        DoEvents
8950        .Width = lngCountry_Left
8960      End Select
8970      If InStr(strSortNow, "Country") > 0 Then
8980        Select Case blnEnable
            Case True
8990          .Sort_line.Left = .Loc_Country_lbl.Left
9000          .Sort_lbl.Left = ((.Loc_Country_lbl.Left + .Loc_Country_lbl.Width) - lngSortLbl_Width)
9010        Case False
9020          SortNow "Form_Load"  ' ** Procedure: Below.
9030        End Select
9040      ElseIf InStr(strSortNow, "PostalCode") > 0 Then
9050        Select Case blnEnable
            Case True
9060          .Sort_line.Left = .Loc_PostalCode_lbl.Left
9070          .Sort_lbl.Left = ((.Loc_PostalCode_lbl.Left + .Loc_PostalCode_lbl.Width) - lngSortLbl_Width)
9080        Case False
9090          SortNow "Form_Load"  ' ** Procedure: Below.
9100        End Select
9110      End If
9120      DoEvents
9130    End With

EXITP:
9140    Exit Sub

ERRH:
9150    Select Case ERR.Number
        Case Else
9160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9170    End Select
9180    Resume EXITP

End Sub

Public Function RecCnt() As Long

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

9210    lngRetVal = 0&

9220    With Me
9230      Set rst = .RecordsetClone
9240      With rst
9250        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
9260        Else
9270          .MoveLast
9280          lngRetVal = .RecordCount
9290        End If
9300        .Close
9310      End With
9320    End With

EXITP:
9330    Set rst = Nothing
9340    RecCnt = lngRetVal
9350    Exit Function

ERRH:
9360    lngRetVal = 0&
9370    Select Case ERR.Number
        Case Else
9380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9390    End Select
9400    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

9510    Select Case intWhere
        Case acCmdRecordsGoToFirst
9520      DoCmd.RunCommand acCmdRecordsGoToFirst
9530    Case acCmdRecordsGoToLast
9540      DoCmd.RunCommand acCmdRecordsGoToLast
9550    Case acCmdRecordsGoToPrevious
9560      DoCmd.RunCommand acCmdRecordsGoToPrevious
9570    Case acCmdRecordsGoToNext
9580      DoCmd.RunCommand acCmdRecordsGoToNext
9590    Case Else
9600      If IsMissing(varID) = False Then
9610        With Me
9620          Set rst = .RecordsetClone
9630          With rst
9640            .FindFirst "[Location_ID] = " & CStr(varID)
9650            If .NoMatch = False Then
9660              Me.Bookmark = .Bookmark
9670            End If
9680            .Close
9690          End With
9700        End With
9710      End If
9720    End Select

EXITP:
9730    Set rst = Nothing
9740    Exit Sub

ERRH:
9750    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
9760    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
9770    Case Else
9780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9790    End Select
9800    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [Loc_Name], [Loc_Address1]"

9910    With Me
9920      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
9930        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
9940      End If
9950      .Controls(strSortLbl).Visible = False
9960      .Controls(strSortLine).Visible = False
9970      .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
9980      .Controls("Sort_lbl2").Visible = False
9990      If strProc = "Form_Load" Then
10000       strCalled = "Loc_Name"
10010       strSortNow = strSortOrig
10020       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
10030       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
10040       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
10050       .Controls(strSortLbl).Left = lngSortLbl_Left
10060       .Controls(strSortLbl).Caption = strArwUp
10070       .Controls(strSortLbl).ForeColor = CLR_DKBLU
10080       .Controls(strSortLine).Left = lngSortLine_Left
10090       .Controls(strSortLine).Width = lngSortLine_Width
10100     Else
10110       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
10120       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
10130       lngSortLine_Left = .Controls(strCalled & "_lbl").Left
10140       lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
10150       .Controls(strSortLbl).Left = lngSortLbl_Left
10160       .Controls(strSortLine).Left = lngSortLine_Left
10170       .Controls(strSortLine).Width = lngSortLine_Width
10180       If strCalled = "Loc_Name" Then
10190         strSortAsc = "[Loc_Name], [Loc_Address1]"
10200         If strSortNow = strSortAsc Then
10210           strSortNow = "[Loc_Name] DESC, [Loc_Address1] DESC"
10220           .Controls(strSortLbl).Caption = strArwDn
10230           .Controls(strSortLbl).ForeColor = CLR_DKRED
10240         Else
10250           strSortNow = strSortOrig
10260           .Controls(strSortLbl).Caption = strArwUp
10270           .Controls(strSortLbl).ForeColor = CLR_DKBLU
10280         End If
10290       Else
10300         .Controls(strSortLbl).Caption = strArwUp
10310         .Controls(strSortLbl).ForeColor = CLR_DKBLU
10320         .Controls("Sort_lbl2").Caption = strArwUp
10330         .Controls("Sort_lbl2").ForeColor = CLR_DKBLU
10340         Select Case strCalled
              Case "Loc_Address1"
10350           strSortAsc = "[Loc_Address1], [Loc_Name]"
10360         Case "Loc_City"
10370           strSortAsc = "[Loc_City]" & strStdAsc
10380         Case "Loc_State"
10390           strSortAsc = "[Loc_State]" & strStdAsc
10400         Case "Loc_Zip"
10410           strSortAsc = "[Loc_Zip]" & strStdAsc
10420         Case "Loc_Country"
10430           strSortAsc = "[Loc_Country]" & strStdAsc
10440         Case "Loc_PostalCode"
10450           strSortAsc = "[Loc_PostalCode]" & strStdAsc
10460         Case "Loc_Phone"
10470           strSortAsc = "[Loc_Phone]" & strStdAsc
10480         End Select
10490         If strSortNow = strSortAsc Then
10500           intCnt = (CharCnt(strSortAsc, ",") + 1&)  ' ** Module Function: modStringFuncs.
10510           Select Case intCnt
                Case 1
10520             strTmp01 = strSortAsc & " DESC"
10530           Case 2
10540             intPos01 = InStr(strSortAsc, ",")
10550             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10560             strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
10570           Case 3
10580             intPos01 = InStr(strSortAsc, ",")
10590             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10600             strTmp02 = Mid(strSortAsc, intPos01)
10610             intPos01 = InStr(2, strTmp02, ",")
10620             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1&)) & " DESC"
10630             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10640           Case 4
10650             intPos01 = InStr(strSortAsc, ",")
10660             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
10670             strTmp02 = Mid(strSortAsc, intPos01)
10680             intPos01 = InStr(2, strTmp02, ",")
10690             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10700             strTmp02 = Mid(strTmp02, intPos01)
10710             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
10720             strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
10730           End Select
10740           strSortNow = strTmp01
10750           .Controls(strSortLbl).Caption = strArwDn
10760           .Controls(strSortLbl).ForeColor = CLR_DKRED
10770           .Controls("Sort_lbl2").Caption = strArwDn
10780           .Controls("Sort_lbl2").ForeColor = CLR_DKRED
10790         Else
10800           strSortNow = strSortAsc
10810         End If
10820       End If
10830     End If
10840     If strCalled <> "Loc_Phone" Then
10850       .Controls(strSortLbl).Visible = True
10860       .Controls("Sort_lbl2").Visible = False
10870     Else
10880       .Controls("Sort_lbl2").Visible = True
10890     End If
10900     If strCalled <> "Loc_Phone" Then
10910       .Controls(strSortLine).Visible = True
10920       .Controls("Loc_Phone_sort_line").Visible = False
10930     Else
10940       .Controls("Loc_Phone_sort_line").Visible = True
10950     End If
10960     .OrderBy = strSortNow
10970     .OrderByOn = True
10980   End With

        ' ** Potential Sorts:
        ' **   [Loc_Name], [Loc_Address1]
        ' **   [Loc_Address1], [Loc_Name]
        ' **   [Loc_City], [Loc_Name], [Loc_Address1]
        ' **   [Loc_State], [Loc_Name], [Loc_Address1]
        ' **   [Loc_Zip], [Loc_Name], [Loc_Address1]
        ' **   [Loc_Country], [Loc_Name], [Loc_Address1]
        ' **   [Loc_PostalCode], [Loc_Name], [Loc_Address1]
        ' **   [Loc_Phone], [Loc_Name], [Loc_Address1]

EXITP:
10990   Exit Sub

ERRH:
11000   Select Case ERR.Number
        Case Else
11010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11020   End Select
11030   Resume EXITP

End Sub

Public Function SortNow_Get() As String

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

11110   strRetVal = strSortNow

EXITP:
11120   SortNow_Get = strRetVal
11130   Exit Function

ERRH:
11140   strRetVal = vbNullString
11150   Select Case ERR.Number
        Case Else
11160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11170   End Select
11180   Resume EXITP

End Function
