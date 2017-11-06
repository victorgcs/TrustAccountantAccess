VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCompanyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCompanyInfo"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Cancel:           E {cmdCancel}
' **   Other:            H {opgLoc_optOther}
' **   Name:             N {CoInfo_Name}
' **   OK:               O {cmdOK}
' **   USA:              U {opgLoc_optUSA}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private intLocOpt As Integer
Private blnIsOpen As Boolean
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)
150           blnIsOpen = True

160           StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

170           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

180           If IsNull(.CoInfo_State_Pref) = False Then
190             .CoInfo_State.DefaultValue = Chr(34) & .CoInfo_State_Pref & Chr(34)
200           End If

210           .cmdOK.Enabled = False
220           .cmdCancel.Enabled = False

230           DoCmd.Hourglass False

240         Else
250           Cancel = -1
260         End If
270       Else
280         Cancel = -1
290       End If
300     End With

310     If Cancel = -1 Then
320       Beep
330       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
340     End If

EXITP:
350     Exit Sub

ERRH:
360     Select Case ERR.Number
        Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
380     End Select
390     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       Select Case IsNull(.CoInfo_Country)
          Case True
430         .opgLoc = .opgLoc_optUSA.OptionValue
440       Case False
450         .opgLoc = .opgLoc_optOther.OptionValue
460       End Select
470       opgLoc_AfterUpdate  ' ** Procedure: Below.
480     End With

EXITP:
490     Exit Sub

ERRH:
500     Select Case ERR.Number
        Case Else
510       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
520     End Select
530     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality.
        ' **   Other:            H {opgLoc_optOther}
        ' **   USA:              U {opgLoc_optUSA}

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyEscape
670         intRetVal = 0
680         cmdClose_Click  ' ** Procedure: Below.
690       End Select
700     End If

        ' ** Alt keys.
710     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
720       Select Case intRetVal
          Case vbKeyH
730         With Me
740           intRetVal = 0
750           .opgLoc = .opgLoc_optOther.OptionValue
760           .opgLoc.SetFocus
770           opgLoc_AfterUpdate  ' ** Procedure: Below.
780         End With
790       Case vbKeyU
800         With Me
810           intRetVal = 0
820           .opgLoc = .opgLoc_optUSA.OptionValue
830           .opgLoc.SetFocus
840           opgLoc_AfterUpdate  ' ** Procedure: Below.
850         End With
860       Case vbKeyX
870         intRetVal = 0
880         cmdClose_Click  ' ** Procedure: Below.
890       End Select
900     End If

        ' ** Ctrl keys.
910     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
920       Select Case intRetVal
          Case vbKeyS
930         intRetVal = 0
940         cmdSave_Click  ' ** Procedure: Below.
950       End Select
960     End If

EXITP:
970     KeyCode = intRetVal
980     Exit Sub

ERRH:
990     intRetVal = 0
1000    THAT_PROC = THIS_PROC
1010    That_Erl = Erl
1020    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1030    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

1110    With Me
1120      .Username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1130      .tmpcoinfo02_datemodified = Now()
1140    End With

EXITP:
1150    Exit Sub

ERRH:
1160    Select Case ERR.Number
        Case Else
1170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1180    End Select
1190    Resume EXITP

End Sub

Private Sub Form_Timer()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1210    With Me
1220      .TimerInterval = 0&
          'CmdBars_Hide True  ' ** Module Function: modWindowFuctions.
1230    End With

EXITP:
1240    Exit Sub

ERRH:
1250    Select Case ERR.Number
        Case Else
1260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
1270    End Select
1280    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1310    If DataErr <> 0 Then
1320      Select Case DataErr
          Case 2046, 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
1330      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
1340        Select Case Me.ActiveControl.Name
            Case "CoInfo_State"
1350          MsgBox "Must enter a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
1360        Case "CoInfo_Zip"
1370          MsgBox "Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
1380        Case "CoInfo_Phone"
1390          MsgBox "Must enter a valid phone number.", vbInformation + vbOKOnly, "Invalid Entry"
1400        End Select
1410        Response = acDataErrContinue
1420      Case 3162
1430        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
1440        DoCmd.CancelEvent
1450      Case Else
1460        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1470        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1480      End Select
1490    End If

EXITP:
1500    Exit Sub

ERRH:
1510    Select Case ERR.Number
        Case Else
1520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1610    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1620    Select Case strCallingForm
        Case "frmMenu_Utility"
1630      gblnSetFocus = True
1640      gstrReturningForm = THIS_NAME
1650      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
1660    Case "frmMenu_Main"
          ' ** If frmMenu_Main opens this form, frmCompanyInfo,
          ' ** it closes it as well, and has been having to, then,
          ' ** close frmMenu_Utility, which it doesn't need.
1670    End Select

EXITP:
1680    Exit Sub

ERRH:
1690    THAT_PROC = THIS_PROC
1700    That_Erl = Erl
1710    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1720    Resume EXITP

End Sub

Private Sub cmdClose_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1810    DoCmd.Close acForm, THIS_NAME

EXITP:
1820    Exit Sub

ERRH:
1830    THAT_PROC = THIS_PROC
1840    That_Erl = Erl
1850    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1860    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1920    intShiftDown = (Shift And acShiftMask) > 0
1930    intAltDown = (Shift And acAltMask) > 0
1940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyTab
1970        With Me
1980          intRetVal = 0
1990          .CoInfo_Name.SetFocus
2000        End With
2010      End Select
2020    End If

        ' ** Shift keys.
2030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2040      Select Case intRetVal
          Case vbKeyTab
2050        With Me
2060          intRetVal = 0
2070          .CoInfo_Phone.SetFocus
2080        End With
2090      End Select
2100    End If

EXITP:
2110    KeyCode = intRetVal
2120    Exit Sub

ERRH:
2130    intRetVal = 0
2140    THAT_PROC = THIS_PROC
2150    That_Erl = Erl
2160    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2170    Resume EXITP

End Sub

Private Sub cmdOK_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strOpt As String
        Dim blnContinue As Boolean, blnEdit As Boolean

2210    With Me
2220      DoCmd.Hourglass True
2230      DoEvents
2240      blnContinue = True
2250      cmdSave_Click  ' ** Procedure: Below.
2260      If Validate = True Then  ' ** Function: Below.

2270        If intLocOpt = .opgLoc_optUSA.OptionValue Then strOpt = "USA" Else strOpt = "Other"

2280        Set dbs = CurrentDb
2290        With dbs
2300          Set rst1 = .OpenRecordset("tmpCompanyInformation2", dbOpenDynaset, dbReadOnly)
2310          If rst1.BOF = True And rst1.EOF = True Then

2320            blnContinue = False
2330            rst1.Close
2340            Set rst1 = Nothing
2350          Else
2360            rst1.MoveFirst
2370            Set rst2 = .OpenRecordset("CompanyInformation", dbOpenDynaset, dbConsistent)
2380            With rst2
2390              .MoveFirst
2400              blnEdit = False
2410              Select Case IsNull(![CoInfo_Name])
                  Case True
2420                .Edit
2430                ![CoInfo_Name] = rst1![CoInfo_Name]
2440                blnEdit = True
2450              Case False
2460                If ![CoInfo_Name] <> rst1![CoInfo_Name] Then
2470                  .Edit
2480                  ![CoInfo_Name] = rst1![CoInfo_Name]
2490                  blnEdit = True
2500                End If
2510              End Select
2520              Select Case IsNull(![CoInfo_Address1])
                  Case True
2530                .Edit
2540                ![CoInfo_Address1] = rst1![CoInfo_Address1]
2550                blnEdit = True
2560              Case False
2570                If ![CoInfo_Address1] <> rst1![CoInfo_Address1] Then
2580                  .Edit
2590                  ![CoInfo_Address1] = rst1![CoInfo_Address1]
2600                  blnEdit = True
2610                End If
2620              End Select
2630              Select Case IsNull(![CoInfo_Address2])
                  Case True
2640                If IsNull(rst1![CoInfo_Address2]) = False Then
2650                  .Edit
2660                  ![CoInfo_Address2] = rst1![CoInfo_Address2]
2670                  blnEdit = True
2680                End If
2690              Case False
2700                Select Case IsNull(rst1![CoInfo_Address2])
                    Case True
2710                  .Edit
2720                  ![CoInfo_Address2] = Null
2730                  blnEdit = True
2740                Case False
2750                  If ![CoInfo_Address2] <> rst1![CoInfo_Address2] Then
2760                    .Edit
2770                    ![CoInfo_Address2] = rst1![CoInfo_Address2]
2780                    blnEdit = True
2790                  End If
2800                End Select
2810              End Select
2820              Select Case IsNull(![CoInfo_City])
                  Case True
2830                .Edit
2840                ![CoInfo_City] = rst1![CoInfo_City]
2850                blnEdit = True
2860              Case False
2870                If ![CoInfo_City] <> rst1![CoInfo_City] Then
2880                  .Edit
2890                  ![CoInfo_City] = rst1![CoInfo_City]
2900                  blnEdit = True
2910                End If
2920              End Select
2930              Select Case IsNull(![CoInfo_State])
                  Case True
2940                If strOpt = "USA" Then
2950                  .Edit
2960                  ![CoInfo_State] = rst1![CoInfo_State]
2970                  blnEdit = True
2980                End If
2990              Case False
3000                Select Case strOpt
                    Case "USA"
3010                  If ![CoInfo_State] <> rst1![CoInfo_State] Then
3020                    .Edit
3030                    ![CoInfo_State] = rst1![CoInfo_State]
3040                    blnEdit = True
3050                  End If
3060                Case "Other"
3070                  .Edit
3080                  ![CoInfo_State] = Null
3090                  blnEdit = True
3100                End Select
3110              End Select
3120              Select Case IsNull(![CoInfo_Zip])
                  Case True
3130                If strOpt = "USA" Then
3140                  .Edit
3150                  ![CoInfo_Zip] = rst1![CoInfo_Zip]
3160                  blnEdit = True
3170                End If
3180              Case False
3190                Select Case strOpt
                    Case "USA"
3200                  If ![CoInfo_Zip] <> rst1![CoInfo_Zip] Then
3210                    .Edit
3220                    ![CoInfo_Zip] = rst1![CoInfo_Zip]
3230                    blnEdit = True
3240                  End If
3250                Case "Other"
3260                  .Edit
3270                  ![CoInfo_Zip] = Null
3280                  blnEdit = True
3290                End Select
3300              End Select
3310              Select Case IsNull(![CoInfo_Country])
                  Case True
3320                If strOpt = "Other" Then
3330                  .Edit
3340                  ![CoInfo_Country] = rst1![CoInfo_Country]
3350                  blnEdit = True
3360                End If
3370              Case False
3380                Select Case strOpt
                    Case "USA"
3390                  .Edit
3400                  ![CoInfo_Country] = Null
3410                  blnEdit = True
3420                Case "Other"
3430                  If ![CoInfo_Country] <> rst1![CoInfo_Country] Then
3440                    .Edit
3450                    ![CoInfo_Country] = rst1![CoInfo_Country]
3460                    blnEdit = True
3470                  End If
3480                End Select
3490              End Select
3500              Select Case IsNull(![CoInfo_PostalCode])
                  Case True
3510                If strOpt = "Other" Then
3520                  If IsNull(rst1![CoInfo_PostalCode]) = False Then
3530                    .Edit
3540                    ![CoInfo_PostalCode] = rst1![CoInfo_PostalCode]
3550                    blnEdit = True
3560                  End If
3570                End If
3580              Case False
3590                Select Case strOpt
                    Case "USA"
3600                  .Edit
3610                  ![CoInfo_PostalCode] = Null
3620                  blnEdit = True
3630                Case "Other"
3640                  Select Case IsNull(rst1![CoInfo_PostalCode])
                      Case True
3650                    .Edit
3660                    ![CoInfo_PostalCode] = Null
3670                    blnEdit = True
3680                  Case False
3690                    If ![CoInfo_PostalCode] <> rst1![CoInfo_PostalCode] Then
3700                      .Edit
3710                      ![CoInfo_PostalCode] = rst1![CoInfo_PostalCode]
3720                      blnEdit = True
3730                    End If
3740                  End Select
3750                End Select
3760              End Select
3770              Select Case IsNull(![CoInfo_Phone])
                  Case True
3780                If IsNull(rst1![CoInfo_Phone]) = False Then
3790                  .Edit
3800                  ![CoInfo_Phone] = rst1![CoInfo_Phone]
3810                  blnEdit = True
3820                End If
3830              Case False
3840                Select Case IsNull(rst1![CoInfo_Phone])
                    Case True
3850                  .Edit
3860                  ![CoInfo_Phone] = Null
3870                  blnEdit = True
3880                Case False
3890                  If ![CoInfo_Phone] <> rst1![CoInfo_Phone] Then
3900                    .Edit
3910                    ![CoInfo_Phone] = rst1![CoInfo_Phone]
3920                    blnEdit = True
3930                  End If
3940                End Select
3950              End Select
3960              If blnEdit = True Then
3970                ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
3980                ![CoInfo_DateModified] = Now()
3990                .Update
4000              End If
4010              .Close
4020            End With
4030            rst1.Close
4040            Set rst1 = Nothing
4050            Set rst2 = Nothing
4060          End If
4070          .Close
4080        End With
4090        Set dbs = Nothing
4100      Else
4110        blnContinue = False
4120      End If  ' ** Validate().
4130      If blnContinue = True Then
4140        DoCmd.Close acForm, THIS_NAME
4150      End If
4160    End With

EXITP:
4170    Set rst1 = Nothing
4180    Set rst2 = Nothing
4190    Set dbs = Nothing
4200    Exit Sub

ERRH:
4210    DoCmd.Hourglass False
4220    Select Case ERR.Number
        Case Else
4230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4240    End Select
4250    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
4370        With Me
4380          intRetVal = 0
4390          .cmdCancel.SetFocus
4400        End With
4410      End Select
4420    End If

        ' ** Shift keys.
4430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4440      Select Case intRetVal
          Case vbKeyTab
4450        With Me
4460          intRetVal = 0
4470          .CoInfo_Phone.SetFocus
4480        End With
4490      End Select
4500    End If

EXITP:
4510    KeyCode = intRetVal
4520    Exit Sub

ERRH:
4530    intRetVal = 0
4540    THAT_PROC = THIS_PROC
4550    That_Erl = Erl
4560    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4570    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

4610    DoCmd.Close acForm, THIS_NAME

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4720    intShiftDown = (Shift And acShiftMask) > 0
4730    intAltDown = (Shift And acAltMask) > 0
4740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4760      Select Case intRetVal
          Case vbKeyTab
4770        With Me
4780          intRetVal = 0
4790          .CoInfo_Name.SetFocus
4800        End With
4810      End Select
4820    End If

        ' ** Shift keys.
4830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4840      Select Case intRetVal
          Case vbKeyTab
4850        With Me
4860          intRetVal = 0
4870          .cmdOK.SetFocus
4880        End With
4890      End Select
4900    End If

EXITP:
4910    KeyCode = intRetVal
4920    Exit Sub

ERRH:
4930    intRetVal = 0
4940    THAT_PROC = THIS_PROC
4950    That_Erl = Erl
4960    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4970    Resume EXITP

End Sub

Private Sub cmdSave_Click()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

5010    DoCmd.RunCommand acCmdSaveRecord
5020    DoEvents

EXITP:
5030    Exit Sub

ERRH:
5040    Select Case ERR.Number
        Case Else
5050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5060    End Select
5070    Resume EXITP

End Sub

Private Sub CoInfo_Name_GotFocus()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Name_GotFocus"

5110    With Me
5120      .CoInfo_Name.SelLength = 0
5130      .CoInfo_Name.SelStart = 99
5140    End With

EXITP:
5150    Exit Sub

ERRH:
5160    THAT_PROC = THIS_PROC
5170    That_Erl = Erl
5180    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5190    Resume EXITP

End Sub

Private Sub CoInfo_Name_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Name_AfterUpdate"

        Dim strTmp01 As String

5210    With Me
5220      If IsNull(.CoInfo_Name) = False Then
5230        strTmp01 = FixQuotes(.CoInfo_Name)  ' ** Module Function: modStringFuncs.
5240        If Compare_StringA_StringB(.CoInfo_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
5250          .CoInfo_Name = strTmp01
5260          cmdSave_Click  ' ** Procedure: Above.
5270        End If
5280      End If
5290    End With

EXITP:
5300    Exit Sub

ERRH:
5310    THAT_PROC = THIS_PROC
5320    That_Erl = Erl
5330    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5340    Resume EXITP

End Sub

Private Sub CoInfo_Name_Change()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Name_Change"

5410    CoInfoChanged  ' ** Procedure: Below.

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
5450    End Select
5460    Resume EXITP

End Sub

Private Sub CoInfo_Name_KeyDown(KeyCode As Integer, Shift As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Name_KeyDown"

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
5590          .CoInfo_Address1.SetFocus
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5650        With Me
5660          intRetVal = 0
5670          .cmdClose.SetFocus
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

Private Sub CoInfo_Address1_AfterUpdate()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Address1_AfterUpdate"

        Dim strTmp01 As String

5810    With Me
5820      If IsNull(.CoInfo_Address1) = False Then
5830        strTmp01 = FixQuotes(.CoInfo_Address1)  ' ** Module Function: modStringFuncs.
5840        If Compare_StringA_StringB(.CoInfo_Address1, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
5850          .CoInfo_Address1 = strTmp01
5860          cmdSave_Click  ' ** Procedure: Above.
5870        End If
5880      End If
5890    End With

EXITP:
5900    Exit Sub

ERRH:
5910    THAT_PROC = THIS_PROC
5920    That_Erl = Erl
5930    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5940    Resume EXITP

End Sub

Private Sub CoInfo_Address1_Change()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Address1_Change"

6010    CoInfoChanged  ' ** Procedure: Below.

EXITP:
6020    Exit Sub

ERRH:
6030    Select Case ERR.Number
        Case Else
6040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
6050    End Select
6060    Resume EXITP

End Sub

Private Sub CoInfo_Address2_AfterUpdate()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Address2_AfterUpdate"

        Dim strTmp01 As String

6110    With Me
6120      If IsNull(.CoInfo_Address2) = False Then
6130        strTmp01 = FixQuotes(.CoInfo_Address2)  ' ** Module Function: modStringFuncs.
6140        If Compare_StringA_StringB(.CoInfo_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
6150          .CoInfo_Address2 = strTmp01
6160          cmdSave_Click  ' ** Procedure: Above.
6170        End If
6180      End If
6190    End With

EXITP:
6200    Exit Sub

ERRH:
6210    THAT_PROC = THIS_PROC
6220    That_Erl = Erl
6230    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6240    Resume EXITP

End Sub

Private Sub CoInfo_Address2_Change()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Address2_Change"

6310    CoInfoChanged  ' ** Procedure: Below.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
6350    End Select
6360    Resume EXITP

End Sub

Private Sub CoInfo_City_AfterUpdate()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_City_AfterUpdate"

        Dim strTmp01 As String

6410    With Me
6420      If IsNull(.CoInfo_City) = False Then
6430        strTmp01 = FixQuotes(.CoInfo_City)  ' ** Module Function: modStringFuncs.
6440        If Compare_StringA_StringB(.CoInfo_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
              ' ** Coeur d'Alene.
6450          .CoInfo_City = strTmp01
6460          cmdSave_Click  ' ** Procedure: Above.
6470        End If
6480      End If
6490    End With

EXITP:
6500    Exit Sub

ERRH:
6510    THAT_PROC = THIS_PROC
6520    That_Erl = Erl
6530    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6540    Resume EXITP

End Sub

Private Sub CoInfo_City_Change()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_City_Change"

6610    CoInfoChanged  ' ** Procedure: Below.

EXITP:
6620    Exit Sub

ERRH:
6630    Select Case ERR.Number
        Case Else
6640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
6650    End Select
6660    Resume EXITP

End Sub

Private Sub opgLoc_AfterUpdate()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_AfterUpdate"

6710    With Me
6720      Select Case .opgLoc
          Case .opgLoc_optUSA.OptionValue
6730        .CoInfo_State.Enabled = True
6740        .CoInfo_State.BorderColor = CLR_LTBLU2
6750        .CoInfo_State.BackStyle = acBackStyleNormal
6760        .CoInfo_State_lbl2.ForeColor = CLR_RED
6770        .CoInfo_State_lbl2_dim_hi.Visible = False
6780        .CoInfo_Zip.Enabled = True
6790        .CoInfo_Zip.BorderColor = CLR_LTBLU2
6800        .CoInfo_Zip.BackStyle = acBackStyleNormal
6810        .CoInfo_Zip_lbl2.ForeColor = CLR_RED
6820        .CoInfo_Zip_lbl2_dim_hi.Visible = False
6830        .CoInfo_Country.Enabled = False
6840        .CoInfo_Country.BorderColor = WIN_CLR_DISR
6850        .CoInfo_Country.BackStyle = acBackStyleTransparent
6860        .CoInfo_Country = Null
6870        .CoInfo_Country_lbl2.ForeColor = WIN_CLR_DISF
6880        .CoInfo_Country_lbl2_dim_hi.Visible = True
6890        .CoInfo_PostalCode.Enabled = False
6900        .CoInfo_PostalCode.BorderColor = WIN_CLR_DISR
6910        .CoInfo_PostalCode.BackStyle = acBackStyleTransparent
6920        .CoInfo_PostalCode = Null
            '.CoInfo_PostalCode_lbl2.ForeColor = WIN_CLR_DISF  ' ** Not required for non-USA.
            '.CoInfo_PostalCode_lbl2_dim_hi.Visible = True
6930        Select Case blnIsOpen
            Case True
6940          blnIsOpen = False
6950        Case False
6960          .CoInfo_State.SetFocus
6970        End Select
6980      Case .opgLoc_optOther.OptionValue
6990        .CoInfo_State.Enabled = False
7000        .CoInfo_State.BorderColor = WIN_CLR_DISR
7010        .CoInfo_State.BackStyle = acBackStyleTransparent
7020        .CoInfo_State = Null
7030        .CoInfo_Zip.Enabled = False
7040        .CoInfo_Zip.BorderColor = WIN_CLR_DISR
7050        .CoInfo_Zip.BackStyle = acBackStyleTransparent
7060        .CoInfo_Zip = Null
7070        .CoInfo_Country.Enabled = True
7080        .CoInfo_Country.BorderColor = CLR_LTBLU2
7090        .CoInfo_Country.BackStyle = acBackStyleNormal
7100        .CoInfo_PostalCode.Enabled = True
7110        .CoInfo_PostalCode.BorderColor = CLR_LTBLU2
7120        .CoInfo_PostalCode.BackStyle = acBackStyleNormal
7130        .CoInfo_State_lbl2.ForeColor = WIN_CLR_DISF
7140        .CoInfo_State_lbl2_dim_hi.Visible = True
7150        .CoInfo_Zip_lbl2.ForeColor = WIN_CLR_DISF
7160        .CoInfo_Zip_lbl2_dim_hi.Visible = True
7170        .CoInfo_Country_lbl2.ForeColor = CLR_RED
7180        .CoInfo_Country_lbl2_dim_hi.Visible = False
            '.CoInfo_PostalCode_lbl2.ForeColor = CLR_RED  ' ** Not required for non-USA.
            '.CoInfo_PostalCode_lbl2_dim_hi.Visible = False
7190        Select Case blnIsOpen
            Case True
7200          blnIsOpen = False
7210        Case False
7220          .CoInfo_Country.SetFocus
7230        End Select
7240      End Select
7250    End With

EXITP:
7260    Exit Sub

ERRH:
7270    Select Case ERR.Number
        Case Else
7280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
7290    End Select
7300    Resume EXITP

End Sub

Private Sub opgLoc_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7420    intShiftDown = (Shift And acShiftMask) > 0
7430    intAltDown = (Shift And acAltMask) > 0
7440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7470        With Me
7480          intRetVal = 0
7490          If .CoInfo_State.Enabled = True Then
7500            .CoInfo_State.SetFocus
7510          Else
7520            .CoInfo_Country.SetFocus
7530          End If
7540        End With
7550      End Select
7560    End If

        ' ** Shift keys.
7570    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7580      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7590        With Me
7600          intRetVal = 0
7610          .CoInfo_City.SetFocus
7620        End With
7630      End Select
7640    End If

EXITP:
7650    KeyCode = intRetVal
7660    Exit Sub

ERRH:
7670    intRetVal = 0
7680    THAT_PROC = THIS_PROC
7690    That_Erl = Erl
7700    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7710    Resume EXITP

End Sub

Private Sub opgLoc_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optOther_KeyDown"

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
7890          If .CoInfo_State.Enabled = True Then
7900            .CoInfo_State.SetFocus
7910          Else
7920            .CoInfo_Country.SetFocus
7930          End If
7940        End With
7950      End Select
7960    End If

        ' ** Shift keys.
7970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7980      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7990        With Me
8000          intRetVal = 0
8010          .CoInfo_City.SetFocus
8020        End With
8030      End Select
8040    End If

EXITP:
8050    KeyCode = intRetVal
8060    Exit Sub

ERRH:
8070    intRetVal = 0
8080    THAT_PROC = THIS_PROC
8090    That_Erl = Erl
8100    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8110    Resume EXITP

End Sub

Private Sub CoInfo_State_Enter()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_State_Enter"

8210    With Me
8220      If IsNull(.CoInfo_State) Then
8230        ExpandCombo .CoInfo_State  ' ** Module Procedure: modUtilities.
8240      End If
8250    End With

EXITP:
8260    Exit Sub

ERRH:
8270    THAT_PROC = THIS_PROC
8280    That_Erl = Erl
8290    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8300    Resume EXITP

End Sub

Private Sub CoInfo_State_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_State_AfterUpdate"

8410    With Me
8420      If IsNull(.CoInfo_State) = False Then
8430        If Trim(.CoInfo_State) <> vbNullString Then
8440          Select Case IsNull(.CoInfo_State_Pref)
              Case True
8450            .CoInfo_State_Pref = .CoInfo_State
8460          Case False
8470            If .CoInfo_State_Pref <> .CoInfo_State Then
8480              .CoInfo_State_Pref = .CoInfo_State
8490            End If
8500          End Select
8510        End If
8520      End If
8530    End With

EXITP:
8540    Exit Sub

ERRH:
8550    Select Case ERR.Number
        Case Else
8560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
8570    End Select
8580    Resume EXITP

End Sub

Private Sub CoInfo_State_Change()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_State_Change"

8610    CoInfoChanged  ' ** Procedure: Below.

EXITP:
8620    Exit Sub

ERRH:
8630    Select Case ERR.Number
        Case Else
8640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
8650    End Select
8660    Resume EXITP

End Sub

Private Sub CoInfo_State_NotInList(NewData As String, Response As Integer)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_State_NotInList"

8710    With Me
8720      .CoInfo_State = Null
8730      MsgBox "You Must select a valid state code.", vbInformation + vbOKOnly, "Entry Required"
8740      Response = acDataErrContinue
8750    End With

EXITP:
8760    Exit Sub

ERRH:
8770    THAT_PROC = THIS_PROC
8780    That_Erl = Erl
8790    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8800    Resume EXITP

End Sub

Private Sub CoInfo_State_KeyUp(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_State_KeyUp"

8910    With Me
8920      If .CoInfo_State.SelStart <> 0 Then
8930        .CoInfo_State.SelStart = 1
8940        .CoInfo_State.SelLength = 0
8950      End If
8960    End With

EXITP:
8970    Exit Sub

ERRH:
8980    THAT_PROC = THIS_PROC
8990    That_Erl = Erl
9000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9010    Resume EXITP

End Sub

Private Sub CoInfo_Zip_Change()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Zip_Change"

9110    CoInfoChanged  ' ** Procedure: Below.

EXITP:
9120    Exit Sub

ERRH:
9130    Select Case ERR.Number
        Case Else
9140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
9150    End Select
9160    Resume EXITP

End Sub

Private Sub CoInfo_Zip_Exit(Cancel As Integer)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Zip_Exit"

9210    With Me
9220      If Not IsNull(.CoInfo_Zip) Then
9230        If Right(.CoInfo_Zip.text, 4) <> "____" And Right(.CoInfo_Zip.text, 4) <> "    " And _
                (Not IsNumeric(Mid(.CoInfo_Zip.text, 7, 1)) Or Not IsNumeric(Mid(.CoInfo_Zip.text, 8, 1)) Or _
                Not IsNumeric(Mid(.CoInfo_Zip.text, 9, 1)) Or Not IsNumeric(Mid(.CoInfo_Zip.text, 10, 1))) Then
9240          MsgBox "Must enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
9250          DoCmd.CancelEvent
9260        End If
9270      End If
9280    End With

EXITP:
9290    Exit Sub

ERRH:
9300    THAT_PROC = THIS_PROC
9310    That_Erl = Erl
9320    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9330    Resume EXITP

End Sub

Private Sub CoInfo_Country_Change()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Country_Change"

9410    CoInfoChanged  ' ** Procedure: Below.

EXITP:
9420    Exit Sub

ERRH:
9430    Select Case ERR.Number
        Case Else
9440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
9450    End Select
9460    Resume EXITP

End Sub

Private Sub CoInfo_PostalCode_Change()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_PostalCode_Change"

9510    CoInfoChanged  ' ** Procedure: Below.

EXITP:
9520    Exit Sub

ERRH:
9530    Select Case ERR.Number
        Case Else
9540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
9550    End Select
9560    Resume EXITP

End Sub

Private Sub CoInfo_Phone_Change()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Phone_Change"

9610    CoInfoChanged  ' ** Procedure: Below.

EXITP:
9620    Exit Sub

ERRH:
9630    Select Case ERR.Number
        Case Else
9640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
9650    End Select
9660    Resume EXITP

End Sub

Private Sub CoInfo_Phone_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Phone_KeyDown"

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
9790          .cmdClose.SetFocus
9800        End With
9810      End Select
9820    End If

        ' ** Shift keys.
9830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9850        With Me
9860          intRetVal = 0
9870          If .CoInfo_Zip.Enabled = True Then
9880            .CoInfo_Zip.SetFocus
9890          Else
9900            .CoInfo_PostalCode.SetFocus
9910          End If
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

Private Function Validate() As Boolean

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "Validate"

        Dim blnContinue As Boolean
        Dim blnRetVal As Boolean

10110   With Me

10120     blnRetVal = False
10130     blnContinue = True

10140     intLocOpt = .opgLoc

10150     Select Case IsNull(.CoInfo_Name)
          Case True
10160       blnContinue = False
10170     Case False
10180       If Trim(.CoInfo_Name) = vbNullString Then
10190         blnContinue = False
10200       End If
10210     End Select
10220     If blnContinue = False Then
10230       DoCmd.Hourglass False
10240       MsgBox "Please enter a Name to continue.", vbInformation + vbOKOnly, "Entry Required"
10250       .CoInfo_Name.SetFocus
10260     Else
10270       Select Case IsNull(.CoInfo_Address1)
            Case True
10280         blnContinue = False
10290       Case False
10300         If Trim(.CoInfo_Address1) = vbNullString Then
10310           blnContinue = False
10320         End If
10330       End Select
10340       If blnContinue = False Then
10350         DoCmd.Hourglass False
10360         MsgBox "Please enter an Address to continue.", vbInformation + vbOKOnly, "Entry Required"
10370         .CoInfo_Address1.SetFocus
10380       Else
10390         Select Case IsNull(.CoInfo_City)
              Case True
10400           blnContinue = False
10410         Case False
10420           If Trim(.CoInfo_City) = vbNullString Then
10430             blnContinue = False
10440           End If
10450         End Select
10460         If blnContinue = False Then
10470           DoCmd.Hourglass False
10480           MsgBox "Please enter a City to continue.", vbInformation + vbOKOnly, "Entry Required"
10490           .CoInfo_City.SetFocus
10500         Else
10510           If intLocOpt = .opgLoc_optUSA.OptionValue And IsNull(.CoInfo_State) = True Then
10520             blnContinue = False
10530           Else
10540             If intLocOpt = .opgLoc_optUSA.OptionValue And Trim(.CoInfo_State) = vbNullString Then
10550               blnContinue = False
10560             End If
10570           End If
10580           If blnContinue = False Then
10590             DoCmd.Hourglass False
10600             MsgBox "Please enter a valid State to continue.", vbInformation + vbOKOnly, "Entry Required"
10610             .CoInfo_State.SetFocus
10620           Else
10630             If intLocOpt = .opgLoc_optUSA.OptionValue And IsNull(.CoInfo_Zip) = True Then
10640               blnContinue = False
10650             Else
10660               If intLocOpt = .opgLoc_optUSA.OptionValue And Trim(.CoInfo_Zip) = vbNullString Then
10670                 blnContinue = False
10680               End If
10690             End If
10700             If blnContinue = False Then
10710               DoCmd.Hourglass False
10720               MsgBox "Please enter a valid Zip Code to continue.", vbInformation + vbOKOnly, "Entry Required"
10730               .CoInfo_Zip.SetFocus
10740             Else
10750               If intLocOpt = .opgLoc_optOther.OptionValue And IsNull(.CoInfo_Country) = True Then
10760                 blnContinue = False
10770               Else
10780                 If intLocOpt = .opgLoc_optOther.OptionValue And Trim(.CoInfo_Country) = vbNullString Then
10790                   blnContinue = False
10800                 End If
10810               End If
10820               If blnContinue = False Then
10830                 DoCmd.Hourglass False
10840                 MsgBox "Please enter a Country to continue," & vbCrLf & _
                        "or change the option to USA and enter State and Zip.", vbInformation + vbOKOnly, "Entry Required"
10850                 .CoInfo_Country.SetFocus
10860               End If  ' ** blnContinue.
10870             End If  ' ** blnContinue.
10880           End If  ' ** blnContinue.
10890         End If  ' ** blnContinue.
10900       End If  ' ** blnContinue.
10910     End If  ' ** blnContinue.

10920     blnRetVal = blnContinue

10930   End With

EXITP:
10940   Validate = blnRetVal
10950   Exit Function

ERRH:
10960   blnRetVal = False
10970   DoCmd.Hourglass False
10980   Select Case ERR.Number
        Case Else
10990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
11000   End Select
11010   Resume EXITP

End Function

Private Sub CoInfoChanged()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfoChanged"

11110   With Me
11120     If .cmdClose.Enabled = True Or .cmdCancel.Enabled = False Then
11130       .cmdOK.Enabled = True
11140       .cmdCancel.Enabled = True
11150       .cmdClose.Enabled = False
11160     End If
11170   End With

EXITP:
11180   Exit Sub

ERRH:
11190   Select Case ERR.Number
        Case Else
11200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
11210   End Select
11220   Resume EXITP

End Sub
