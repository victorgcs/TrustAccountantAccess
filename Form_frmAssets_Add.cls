VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssets_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssets_Add"

'VGC 12/31/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Asset Type:       A {assettype}
' **   Cancel:           C {cmdCancel}
' **   Description:      D {description_masterasset}
' **   OK:               O {cmdOK}
' **   CUSIP:            U {cusip}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnIsCancel As Boolean, strGoodCusip As String
Private strCallingForm As String, strCallersCallingForm As String
Private lngAssetNo As Long, blnSetToRate As Boolean
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim varTmp00 As Variant

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)
150           intPos01 = InStr(strCallingForm, "~")
160           strCallersCallingForm = Mid(strCallingForm, (intPos01 + 1))
170           strCallingForm = Left(strCallingForm, (intPos01 - 1))

180           CLR_DISABLED_FG = CLR_DKGRY
190           CLR_DISABLED_BG = CLR_LTTEAL

200           strGoodCusip = vbNullString
210           blnSetToRate = False: blnIsCancel = False

220           If strCallersCallingForm <> "frmJournal_Columns" Then
230             .Modal = True
240           End If

250           varTmp00 = DMax("[assetno]", "tmpEdit09")
260           If IsNull(varTmp00) = False Then
270             lngAssetNo = CLng(varTmp00) + 1&
280           End If

290           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

300         Else
310           Cancel = -1
320         End If
330       Else
340         Cancel = -1
350       End If
360     End With

370     If Cancel = -1 Then
380       Beep
390       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
400     End If

EXITP:
410     Exit Sub

ERRH:
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me
520       .assettype_description.ForeColor = CLR_DISABLED_FG
530       .assettype_description.BackColor = CLR_DISABLED_BG
540       .curr_id.Visible = True
550       .curr_name.Visible = True
560       .curr_name.ForeColor = CLR_DISABLED_FG
570       .curr_name.BackColor = CLR_DISABLED_BG
580     End With

EXITP:
590     Exit Sub

ERRH:
600     Select Case ERR.Number
        Case Else
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End Select
630     Resume EXITP

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

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         intRetVal = 0
780         cmdCancel_Click  ' ** Procedure: Below.
790       End Select
800     End If

        ' ** Ctrl keys.
810     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
820       Select Case intRetVal
          Case vbKeyS
830         intRetVal = 0
840         cmdSave_Click  ' ** Procedure: Below.
850       End Select
860     End If

EXITP:
870     KeyCode = intRetVal
880     Exit Sub

ERRH:
890     intRetVal = 0
900     THAT_PROC = THIS_PROC
910     That_Erl = Erl
920     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
930     Resume EXITP

End Sub

Private Sub Form_Timer()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1010    With Me
1020      .TimerInterval = 0
1030      If blnSetToRate = True Then
1040        blnSetToRate = False
1050        .rate_entry.SetFocus
1060      End If
1070    End With

EXITP:
1080    Exit Sub

ERRH:
1090    Select Case ERR.Number
        Case Else
1100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1110    End Select
1120    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim strControl As String

1210    If DataErr <> 0 Then
1220      Select Case DataErr
          Case 13, 2108, 2116, 2135, 2169, 2474, 3020, 7753, 8519, 10503, 10508
1230        Response = acDataErrContinue
1240      Case 2110  ' ** Microsoft Access can't move the focus to the control '|'.
1250        Response = acDataErrContinue
1260      Case 2113, 2279
            ' ** 2113: The value you entered isn't valid for this field.
            ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
1270        Select Case Me.ActiveControl.Name
            Case "due"
1280          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Entry Required"
1290        Case "rate"
1300          MsgBox "Please enter a valid number.", vbInformation + vbOKOnly, "Entry Required"
1310        Case "assettype"
1320          MsgBox "Please select a valid asset code.", vbInformation + vbOKOnly, "Entry Required"
1330        Case "cusip"
1340          MsgBox "Please enter a valid CUSIP.", vbInformation + vbOKOnly, "Entry Required"
1350        End Select
1360        Response = acDataErrContinue
1370      Case 2221  ' ** The text is too long to be editied.
1380        strControl = Screen.ActiveControl.Name
1390        Select Case strControl
            Case "cusip" ' 9
1400          MsgBox "The CUSIP is limited to 9 characters.", vbInformation + vbOKOnly, "Invalid Entry"
1410        Case "description_masterasset" ' 75
1420          MsgBox "The Description is limited to 75 characters.", vbInformation + vbOKOnly, "Invalid Entry"
1430        Case Else
1440          MsgBox "Your entry is too long.", vbInformation + vbOKOnly, "Invalid Entry"
1450        End Select
1460        Response = acDataErrContinue
1470      Case 2237, 3101
            ' ** 2237: The text you entered isn't an item in the list.
            ' ** 3101: The Microsoft Jet database engine cannot find a record in the table '|' with key matching field(s) '|'.
1480        If IsNull(Me.assettype) Or Trim(Me.assettype) = vbNullString Then
1490          MsgBox "Please enter a valid asset code.", vbInformation + vbOKOnly, "Invalid Entry"
1500        End If
1510        Me.assettype.SetFocus
1520        Response = acDataErrContinue
1530      Case 3162  ' ** 3162: You tried to assign the Null value to a variable that is not a Variant data type.
1540        Response = acDataErrContinue
1550      Case 3314  ' ** The field '|' cannot contain a Null value because the Required property for this field is set to True.
1560        Me.Undo
1570        Me.Undo
1580        Response = acDataErrContinue
1590        cmdCancel_Click  ' ** Procedure: Below.
1600      Case Else
1610        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1620        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1630        Response = acDataErrContinue
1640      End Select
1650    End If

EXITP:
1660    Exit Sub

ERRH:
1670    Select Case ERR.Number
        Case Else
1680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1690    End Select
1700    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1810    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOk_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strCusip As String, strAssetDesc As String
        Dim lngCurrID As Long
        Dim varTmp00 As Variant
        Dim blnContinue As Boolean

1910    blnContinue = True

1920    With Me
1930      If lngAssetNo = 0& Then
1940        varTmp00 = DMax("[assetno]", "tmpEdit09")
1950        If IsNull(varTmp00) = False Then
1960          lngAssetNo = CLng(varTmp00) + 1&
1970        End If
1980      End If
1990      If IsNull(.assetno) = True Then
2000        .assetno = lngAssetNo
2010      End If
2020      If IsNull(.cusip) = True Then
2030        blnContinue = False
2040        MsgBox "CUSIP must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
2050        .cusip.SetFocus
2060      Else
2070        If Trim(.cusip) = vbNullString Then
2080          blnContinue = False
2090          MsgBox "CUSIP must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
2100          .cusip.SetFocus
2110        Else
2120          If Trim(.cusip) <> strGoodCusip Then
2130            Set dbs = CurrentDb
                ' ** MasterAsset, by specified [cusp].
2140            Set qdf = dbs.QueryDefs("qryAssets_09")
2150            With qdf.Parameters
2160              ![cusp] = Trim(Me.cusip)
2170            End With
2180            Set rst = qdf.OpenRecordset
2190            With rst
2200              If .BOF = True And .EOF = True Then
                    ' ** OK, nothing found.
2210                strGoodCusip = Trim(Me.cusip)
2220              Else
2230                blnContinue = False
2240                strCusip = ![cusip]
2250                strAssetDesc = ![description_masterasset]
2260                strGoodCusip = vbNullString
2270              End If
2280              .Close
2290            End With  ' ** rst.
2300            dbs.Close
2310            Set rst = Nothing
2320            Set qdf = Nothing
2330            Set dbs = Nothing
2340            If blnContinue = False Then
2350              MsgBox "Cusip already exists." & Space(40) & vbCrLf & vbCrLf & strCusip & vbCrLf & _
                    strAssetDesc, vbInformation + vbOKOnly, "CUSIP Already Present"
2360              .cusip.SetFocus
2370            End If
2380          End If  ' ** strGoodCusip.
2390        End If
2400      End If

2410      If blnContinue = True Then
2420        If IsNull(.description_masterasset) = True Then
2430          blnContinue = False
2440          MsgBox "Description must be added to continue.", vbInformation + vbOKOnly, "Entry Required"
2450          .description_masterasset.SetFocus
2460        Else
2470          If Trim(.description_masterasset) = vbNullString Then
2480            blnContinue = False
2490            MsgBox "Description must be added to continue.", vbInformation + vbOKOnly, "Entry Required"
2500            .description_masterasset.SetFocus
2510          Else
2520            If IsNull(.assettype_description) = True Then
2530              blnContinue = False
2540              MsgBox "Please enter a valid asset code.", vbInformation + vbOKOnly, "Entry Required"
2550              .assettype.SetFocus
2560            Else
2570              If Trim(.assettype_description) = vbNullString Then
2580                blnContinue = False
2590                MsgBox "Please enter a valid asset code.", vbInformation + vbOKOnly, "Entry Required"
2600                .assettype.SetFocus
2610              End If
2620            End If
2630          End If
2640        End If
2650      End If  ' ** blnContinue.

2660      If blnContinue = True Then

2670        If IsNull(.rate) = True Then
2680          If IsNull(.rate_entry) = True Then
2690            .rate_entry = 0#
2700            .rate = 0#
2710          Else
2720            If .rate_entry <> 0& Then
2730              .rate = .rate_entry / 100#  ' ** Saved as a percentage, i.e., must be less than 1!
2740            Else
2750              .rate = 0#
2760            End If
2770          End If
2780        End If

2790        lngCurrID = .curr_id

2800        If IsNull(.shareface) = True Then .shareface = 0#
2810        .RecChanged = 2

2820        cmdSave_Click  ' ** Procedure: Below.

2830        With Forms(strCallingForm)
2840          .frmAssets_Sub.Form.Requery
2850          .frmAssets_Sub.Form.MoveRec 0, lngAssetNo  ' ** Form Procedure: frmAssets_Sub.
2860          .cmdEdit.Enabled = True
2870          .cmdEdit_Click  ' ** Form Procedure: frmAssets.
2880          .chkUpdated = True
2890          .cmdUpdate.Enabled = True
2900          If lngCurrID <> 150& Then  ' ** USD.
2910            If .chkIncludeCurrency.Enabled = False Then
2920              .chkIncludeCurrency.Enabled = True
2930            End If
2940            If .chkIncludeCurrency = False Then
2950              .chkIncludeCurrency = True
2960              .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmAssets.
2970            End If
2980          End If
2990          .NewAssets = .NewAssets + 1&
3000          .RecalcTots  ' ** Form Procedure: frmAssets.
3010        End With

            ' ** This form goes into tmpEdit09.
3020        DoCmd.Close acForm, THIS_NAME

3030      End If  ' ** blnContinue.

3040    End With  ' ** Me.

EXITP:
3050    Set rst = Nothing
3060    Set qdf = Nothing
3070    Set dbs = Nothing
3080    Exit Sub

ERRH:
3090    Select Case ERR.Number
        Case 2046, 3031  ' ** 2046: The command or action isn't available now (first or last record).
3100      MsgBox "The form must be complete to continue.", vbInformation + vbOKOnly, "Entry Required"  ' ** 3031: Not a valid password.
3110    Case Else
3120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3130    End Select
3140    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3220    intShiftDown = (Shift And acShiftMask) > 0
3230    intAltDown = (Shift And acAltMask) > 0
3240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyTab
3270        With Me
3280          intRetVal = 0
3290          .cmdCancel.SetFocus
3300        End With
3310      End Select
3320    End If

        ' ** Shift keys.
3330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3340      Select Case intRetVal
          Case vbKeyTab
3350        With Me
3360          intRetVal = 0
3370          .curr_id.SetFocus
3380        End With
3390      End Select
3400    End If

EXITP:
3410    KeyCode = intRetVal
3420    Exit Sub

ERRH:
3430    intRetVal = 0
3440    THAT_PROC = THIS_PROC
3450    That_Erl = Erl
3460    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3470    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

3510    With Me
3520      blnIsCancel = True
3530      .Undo
3540      Forms(strCallingForm).frmAssets_Sub.Requery
3550      DoCmd.Close acForm, THIS_NAME
3560    End With

EXITP:
3570    Exit Sub

ERRH:
3580    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
3590      Resume Next
3600    Case 3031, 3101  ' ** 3031: Not a valid password. ??  ' ** 3101: The Microsoft Jet database engine cannot find a record in the table '|2' with key matching field(s) '|1'.
3610      Me.Undo
3620      Resume Next
3630    Case 3058  ' ** Index or primary key cannot contain a Null value.
3640      Me.Undo
3650      Resume Next
3660    Case Else
3670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3680    End Select
3690    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3720    intShiftDown = (Shift And acShiftMask) > 0
3730    intAltDown = (Shift And acAltMask) > 0
3740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3760      Select Case intRetVal
          Case vbKeyTab
3770        With Me
3780          intRetVal = 0
3790          .cusip.SetFocus
3800        End With
3810      End Select
3820    End If

        ' ** Shift keys.
3830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3840      Select Case intRetVal
          Case vbKeyTab
3850        With Me
3860          intRetVal = 0
3870          .cmdOK.SetFocus
3880        End With
3890      End Select
3900    End If

EXITP:
3910    KeyCode = intRetVal
3920    Exit Sub

ERRH:
3930    intRetVal = 0
3940    THAT_PROC = THIS_PROC
3950    That_Erl = Erl
3960    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3970    Resume EXITP

End Sub

Private Sub cmdSave_Click()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

        Dim varTmp00 As Variant

4010    With Me
4020      If lngAssetNo = 0& Then
4030        varTmp00 = DMax("[assetno]", "tmpEdit09")
4040        If IsNull(varTmp00) = False Then
4050          lngAssetNo = CLng(varTmp00) + 1&
4060        End If
4070      End If
4080      If IsNull(.assetno) = True Then
4090        .assetno = lngAssetNo
4100      End If
4110    End With

4120    DoCmd.RunCommand acCmdSelectRecord
4130    DoCmd.RunCommand acCmdSaveRecord

EXITP:
4140    Exit Sub

ERRH:
4150    Select Case ERR.Number
        Case Else
4160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4170    End Select
4180    Resume EXITP

End Sub

Private Sub cusip_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_KeyDown"

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
4290          .description_masterasset.SetFocus
4300        End With
4310      End Select
4320    End If

        ' ** Shift keys.
4330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4350        With Me
4360          intRetVal = 0
4370          .cmdCancel.SetFocus
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

Private Sub cusip_Exit(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_Exit"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strCusip As String, strAssetDesc As String
        Dim blnContinue As Boolean

        Static intInstance As Integer

4510    blnContinue = True
4520    strCusip = vbNullString: strAssetDesc = vbNullString

4530    With Me
4540      If IsNull(.cusip) = False Then
4550        If Trim(.cusip) <> vbNullString Then

4560          Set dbs = CurrentDb
4570          With dbs
                ' ** MasterAsset, by specified [cusp].
4580            Set qdf = .QueryDefs("qryAssets_09")
4590            With qdf.Parameters
4600              ![cusp] = Trim(Me.cusip)
4610            End With
4620            Set rst = qdf.OpenRecordset
4630            With rst
4640              If .BOF = True And .EOF = True Then
                    ' ** OK, nothing found.
4650                strGoodCusip = Trim(Me.cusip)
4660              Else
4670                intInstance = intInstance + 1
4680                .MoveFirst
4690                Select Case blnIsCancel
                    Case True
                      ' ** Skip the message.
4700                Case False
4710                  If intInstance <= 1 Then
4720                    blnContinue = False
4730                    strCusip = ![cusip]
4740                    strAssetDesc = ![description_masterasset]
4750                  Else
4760                    Me.cusip.Undo
4770                  End If
4780                End Select
4790                strGoodCusip = vbNullString
4800              End If
4810              .Close
4820            End With  ' ** .rst
4830            .Close
4840          End With  ' ** dbs.

4850          If blnContinue = False Then
4860            MsgBox "Cusip already exists." & Space(40) & vbCrLf & vbCrLf & strCusip & vbCrLf & _
                  strAssetDesc, vbInformation + vbOKOnly, "CUSIP Already Present"
4870            Cancel = -1
4880            .cusip.SetFocus
4890          End If

4900        End If
4910      End If
4920    End With

EXITP:
4930    Set rst = Nothing
4940    Set qdf = Nothing
4950    Set dbs = Nothing
4960    Exit Sub

ERRH:
4970    THAT_PROC = THIS_PROC
4980    That_Erl = Erl
4990    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5000    Resume EXITP

End Sub

Private Sub description_masterasset_AfterUpdate()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "description_masterasset_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String

5110    With Me
5120      If IsNull(.description_masterasset) = False Then
5130        If Trim(.description_masterasset) <> vbNullString Then
5140          strTmp01 = Trim(.description_masterasset)
5150          intPos01 = InStr(strTmp01, Chr(34))  ' ** No double-quotes allowed, so just replace with single-quote, apostrophe.
5160          If intPos01 > 0 Then
5170            Do While intPos01 > 0
5180              strTmp01 = Left(strTmp01, (intPos01 - 1)) & "'" & Mid(strTmp01, (intPos01 + 1))
5190              strTmp01 = Trim(strTmp01)
5200              intPos01 = InStr(strTmp01, Chr(34))
5210            Loop
5220            .description_masterasset = strTmp01
5230            Beep
5240          End If
5250        End If
5260      End If
5270    End With

EXITP:
5280    Exit Sub

ERRH:
5290    Select Case ERR.Number
        Case Else
5300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5310    End Select
5320    Resume EXITP

End Sub

Private Sub rate_entry_AfterUpdate()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "rate_entry_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String, dblTmp02 As Double

5410    With Me
5420      If IsNull(.rate_entry) = True Then
5430        .rate_entry = 0#
5440        .rate = 0#
5450      End If
5460      strTmp01 = CStr(.rate_entry)
5470      strTmp01 = Trim(strTmp01)
5480      If Right(strTmp01, 1) = "%" Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
5490      If IsNumeric(strTmp01) = True Then
5500        dblTmp02 = CDbl(strTmp01)
5510        If dblTmp02 >= 100# Then
5520          MsgBox "Rate must be less than 100 percent.", vbInformation + vbOKOnly, "Invalid Entry"
5530          .rate_entry = Null
5540          .rate = Null
5550          .rate_entry.SetFocus
5560          blnSetToRate = True
5570          .TimerInterval = 100&
5580        Else
5590          intPos01 = InStr(strTmp01, ".")
5600          If intPos01 > 0 Then
                ' ** The user entered a decimal number.
5610            If Len(Mid(strTmp01, (intPos01 + 1))) > 3 Then
5620              MsgBox "You can only enter a rate percentage with up to 3 decimal places.", vbInformation + vbOKOnly, "Invalid Entry"
5630              .rate_entry = Null
5640              .rate = Null
5650              .rate_entry.SetFocus
5660              blnSetToRate = True
5670              .TimerInterval = 100&
5680            Else
5690              .rate = dblTmp02 / 100#  ' ** Saved as a percentage, i.e., must be less than 1!
5700              .due.SetFocus
5710            End If
5720          Else
5730            If dblTmp02 <> 0# Then
5740              .rate = dblTmp02 / 100#  ' ** Saved as a percentage, i.e., must be less than 1!
5750            Else
5760              .rate = 0#
5770            End If
5780            .due.SetFocus
5790          End If
5800        End If
5810      Else
5820        MsgBox "Please enter a valid rate percentage.", vbInformation + vbOKOnly, "Invalid Entry"
5830        .rate_entry = Null
5840        .rate = Null
5850        .rate_entry.SetFocus
5860        blnSetToRate = True
5870        .TimerInterval = 100&
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

Private Sub assettype_Enter()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_Enter"

6010    With Me
6020      If IsNull(.assettype) Then
6030        ExpandCombo  ' ** Module Procedure: modUtilities.
6040      End If
6050    End With

EXITP:
6060    Exit Sub

ERRH:
6070    THAT_PROC = THIS_PROC
6080    That_Erl = Erl
6090    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6100    Resume EXITP

End Sub

Private Sub assettype_AfterUpdate()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAssetType As String

6210    With Me
6220      If IsNull(.assettype) = True Then
6230        MsgBox "Please select a valid asset code.", vbInformation + vbOKOnly, "Entry Required"
6240        .assettype_description = Null
6250        .assettype.SetFocus
6260      Else
6270        strAssetType = .assettype
6280        Set dbs = CurrentDb
            ' ** AssetType, by specified [acod].
6290        Set qdf = dbs.QueryDefs("qryAssets_10")
6300        With qdf.Parameters
6310          ![acod] = strAssetType
6320        End With
6330        Set rst = qdf.OpenRecordset
6340        If rst.BOF = True And rst.EOF = True Then
              ' ** Not found?!
6350          MsgBox "Please select a valid asset code.", vbInformation + vbOKOnly, "Invalid Entry"
6360          .assettype_description = Null
6370          .assettype.SetFocus
6380        Else
6390          If strAssetType = "81" Then
6400            .assettype_description = "Interest Bearing"
6410          Else
6420            .assettype_description = rst![assettype_description]
6430          End If
6440        End If
6450        rst.Close
6460        dbs.Close
6470      End If
6480    End With

EXITP:
6490    Set rst = Nothing
6500    Set qdf = Nothing
6510    Set dbs = Nothing
6520    Exit Sub

ERRH:
6530    Select Case ERR.Number
        Case Else
6540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6550    End Select
6560    Resume EXITP

End Sub

Private Sub assettype_NotInList(NewData As String, Response As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_NotInList"

6610    With Me
6620      MsgBox "Please select a valid asset code.", vbInformation + vbOKOnly, "Invalid Entry"
6630      .assettype.SetFocus
6640      Response = acDataErrContinue
6650    End With

EXITP:
6660    Exit Sub

ERRH:
6670    THAT_PROC = THIS_PROC
6680    That_Erl = Erl
6690    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6700    Resume EXITP

End Sub

Private Sub assettype_KeyDown(KeyCode As Integer, Shift As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_KeyDown"

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
6890          .curr_id.SetFocus
6900        End With
6910      End Select
6920    End If

        ' ** Shift keys.
6930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6950        With Me
6960          intRetVal = 0
6970          .due.SetFocus
6980        End With
6990      End Select
7000    End If

EXITP:
7010    KeyCode = intRetVal
7020    Exit Sub

ERRH:
7030    intRetVal = 0
7040    THAT_PROC = THIS_PROC
7050    That_Erl = Erl
7060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7070    Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim lngCurrID As Long

7110    With Me
7120      If IsNull(.curr_id) = True Then
7130        MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Entry Required"
7140        .curr_name = Null
7150        .curr_id.SetFocus
7160      Else
7170        lngCurrID = .curr_id
7180        .curr_name = .curr_id.Column(2)

7190      End If
7200    End With

EXITP:
7210    Exit Sub

ERRH:
7220    Select Case ERR.Number
        Case Else
7230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7240    End Select
7250    Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

7310    With Me
7320      MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
7330      .curr_id.SetFocus
7340      Response = acDataErrContinue
7350    End With

EXITP:
7360    Exit Sub

ERRH:
7370    THAT_PROC = THIS_PROC
7380    That_Erl = Erl
7390    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7400    Resume EXITP

End Sub

Private Sub curr_id_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7520    intShiftDown = (Shift And acShiftMask) > 0
7530    intAltDown = (Shift And acAltMask) > 0
7540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7570        With Me
7580          intRetVal = 0
7590          .cmdOK.SetFocus
7600        End With
7610      End Select
7620    End If

        ' ** Shift keys.
7630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7650        With Me
7660          intRetVal = 0
7670          .assettype.SetFocus
7680        End With
7690      End Select
7700    End If

EXITP:
7710    KeyCode = intRetVal
7720    Exit Sub

ERRH:
7730    intRetVal = 0
7740    THAT_PROC = THIS_PROC
7750    That_Erl = Erl
7760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7770    Resume EXITP

End Sub
