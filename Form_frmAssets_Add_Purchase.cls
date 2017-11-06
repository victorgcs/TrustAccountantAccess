VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssets_Add_Purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssets_Add_Purchase"

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
Private strCallingForm As String
Private lngAssetNo As Long, blnSetToRate As Boolean
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
170             lngAssetNo = Val(Mid(strCallingForm, (intPos01 + 1)))
180             strCallingForm = Left(strCallingForm, (intPos01 - 1))
190           End If

200           CLR_DISABLED_FG = CLR_DKGRY
210           CLR_DISABLED_BG = CLR_LTTEAL

220           strGoodCusip = vbNullString
230           blnSetToRate = False: blnIsCancel = False

240           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

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
1020      .TimerInterval = 0&
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
1490          MsgBox "Please enter a valid asset code.", vbInformation + vbOKOnly, "Entry Required"
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
1630      End Select
1640    End If

EXITP:
1650    Exit Sub

ERRH:
1660    Select Case ERR.Number
        Case Else
1670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1680    End Select
1690    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1710    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOk_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strCusip As String, strAssetDesc As String
        Dim lngCurrID As Long
        Dim blnContinue As Boolean

1810    blnContinue = True

1820    With Me
1830      If IsNull(.cusip) = True Then
1840        blnContinue = False
1850        MsgBox "CUSIP must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1860        .cusip.SetFocus
1870      Else
1880        If Trim(.cusip) = vbNullString Then
1890          blnContinue = False
1900          MsgBox "CUSIP must be filled to continue.", vbInformation + vbOKOnly, "Entry Required"
1910          .cusip.SetFocus
1920        Else
1930          If Trim(.cusip) <> strGoodCusip Then
1940            Set dbs = CurrentDb
                ' ** MasterAsset, by specified [cusp].
1950            Set qdf = dbs.QueryDefs("qryAssets_09")
1960            With qdf.Parameters
1970              ![cusp] = Trim(Me.cusip)
1980            End With
1990            Set rst = qdf.OpenRecordset
2000            With rst
2010              If .BOF = True And .EOF = True Then
                    ' ** OK, nothing found.
2020                strGoodCusip = Trim(Me.cusip)
2030              Else
2040                blnContinue = False
2050                strCusip = ![cusip]
2060                strAssetDesc = ![description_masterasset]
2070                strGoodCusip = vbNullString
2080              End If
2090              .Close
2100            End With  ' ** rst.
2110            dbs.Close
2120            Set rst = Nothing
2130            Set qdf = Nothing
2140            Set dbs = Nothing
2150            If blnContinue = False Then
2160              MsgBox "Cusip already exists." & Space(40) & vbCrLf & vbCrLf & strCusip & vbCrLf & _
                    strAssetDesc, vbInformation + vbOKOnly, "CUSIP Already Present"
2170              .cusip.SetFocus
2180            End If
2190          End If  ' ** strGoodCusip.
2200        End If
2210      End If

2220      If blnContinue = True Then
2230        If IsNull(.description_masterasset) = True Then
2240          blnContinue = False
2250          MsgBox "Description must be added to continue.", vbInformation + vbOKOnly, "Entry Required"
2260          .description_masterasset.SetFocus
2270        Else
2280          If Trim(.description_masterasset) = vbNullString Then
2290            blnContinue = False
2300            MsgBox "Description must be added to continue.", vbInformation + vbOKOnly, "Entry Required"
2310            .description_masterasset.SetFocus
2320          Else
2330            If IsNull(.assettype_description) = True Then
2340              blnContinue = False
2350              MsgBox "Please enter a valid asset code.", vbInformation + vbOKOnly, "Entry Required"
2360              .assettype.SetFocus
2370            Else
2380              If Trim(.assettype_description) = vbNullString Then
2390                blnContinue = False
2400                MsgBox "Please enter a valid asset code.", vbInformation + vbOKOnly, "Entry Required"
2410                .assettype.SetFocus
2420              End If
2430            End If
2440          End If
2450        End If
2460      End If  ' ** blnContinue.

2470      If blnContinue = True Then

2480        If IsNull(.rate) = True Then
2490          If IsNull(.rate_entry) = True Then
2500            .rate_entry = 0#
2510            .rate = 0#
2520          Else
2530            If .rate_entry <> 0& Then
2540              .rate = .rate_entry / 100#
2550            Else
2560              .rate = 0#
2570            End If
2580          End If
2590        End If

2600        lngCurrID = .curr_id

2610        If IsNull(.shareface) = True Then .shareface = 0#

2620        cmdSave_Click  ' ** Procedure: Below.

2630        lngAssetNo = .assetno

2640        Select Case strCallingForm
            Case "frmJournal"
2650          With Forms(strCallingForm).frmJournal_Sub3_Purchase.Form
2660            .purchaseAssetNo.Requery
2670            DoEvents
2680            .purchaseAssetNo = lngAssetNo
2690            .purchaseCurr_ID = lngCurrID
2700            If lngCurrID <> 150& Then
2710              If .purchaseCurr_ID.Visible = False Then
2720                Set dbs = CurrentDb
2730                With dbs
                      ' ** tblPreference_User, for 'chkIncludeCurrency' on 'frmPostingDate', just dbs_id = 1, by specified [usr].
2740                  Set qdf = .QueryDefs("qryPreferences_07_01")
2750                  With qdf.Parameters
2760                    ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2770                  End With
2780                  Set rst = qdf.OpenRecordset
2790                  If rst.BOF = True And rst.EOF = True Then
                        ' ** No preference?
2800                    rst.Close
2810                    Set rst = Nothing
2820                    Set qdf = Nothing
                        ' ** Append 'chkIncludeCurrency' record to tblPreference_User,
                        ' ** just dbs_id = 1, by specified [usr], [pbln].
2830                    Set qdf = .QueryDefs("qryPreferences_07_02")
2840                    With qdf.Parameters
2850                      ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2860                      ![pbln] = True
2870                    End With
2880                    qdf.Execute
2890                    Set qdf = Nothing
2900                  Else
2910                    rst.MoveFirst
2920                    Select Case rst![prefuser_boolean]
                        Case True
                          ' ** Not sure how it could be True and yet not be showing currency.
2930                      rst.Close
2940                      Set rst = Nothing
2950                      Set qdf = Nothing
2960                    Case False
                          ' ** Was False, so flip it.
2970                      rst.Close
2980                      Set rst = Nothing
2990                      Set qdf = Nothing
                          ' ** Update tblPreference_User, just 'chkIncludeCurrency'
                          ' ** on 'frmPostingDate', by specified [usr], [pbln].
3000                      Set qdf = .QueryDefs("qryPreferences_07_03")
3010                      With qdf.Parameters
3020                        ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
3030                        ![pbln] = True
3040                      End With
3050                      qdf.Execute
3060                      Set qdf = Nothing
3070                    End Select
3080                  End If
3090                  .Close
3100                End With
3110                Set dbs = Nothing
3120                .Parent.IncludeCurrency  ' ** Form Procedure: frmJournal.
3130              End If
3140            End If
3150          End With
3160        Case "frmJournal_Columns"
3170          gdblCrtRpt_CostTot = .assetno  ' ** Borrowing this variable from the Court Reports.
3180          gstrPurchaseAsset = .description_masterasset
3190        End Select
3200        DoEvents

            ' ** This form goes directly into MasterAsset.
3210        DoCmd.Close acForm, THIS_NAME

3220        DoCmd.SelectObject acForm, strCallingForm, False
3230        Select Case strCallingForm
            Case "frmJournal"
3240          With Forms(strCallingForm).frmJournal_Sub3_Purchase.Form
3250            .purchaseAssetNo.SetFocus
3260            gstrPurchaseAsset = .purchaseAssetNo.text
3270            .purchaseShareFace.SetFocus
3280          End With
3290        Case "frmJournal_Columns"
              ' ** Let JC_Frm_Assets() set the focus.
3300        End Select

3310      End If  ' ** blnContinue.
3320    End With  ' ** Me.

EXITP:
3330    Set rst = Nothing
3340    Set qdf = Nothing
3350    Set dbs = Nothing
3360    Exit Sub

ERRH:
3370    THAT_PROC = THIS_PROC
3380    That_Erl = Erl
3390    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3400    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
3570        With Me
3580          intRetVal = 0
3590          .cmdCancel.SetFocus
3600        End With
3610      End Select
3620    End If

        ' ** Shift keys.
3630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3640      Select Case intRetVal
          Case vbKeyTab
3650        With Me
3660          intRetVal = 0
3670          .curr_id.SetFocus
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

Private Sub cmdCancel_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

3810    With Me

3820      blnIsCancel = True
3830      If .NewRecord = True Then
3840  On Error Resume Next  ' ** If it's unavailable, e.g., new record with nothing in it, just continue to close.
3850        .Undo
3860  On Error GoTo ERRH
3870      Else
3880        DoCmd.RunCommand acCmdSelectRecord
3890        DoCmd.RunCommand acCmdDeleteRecord
3900      End If
3910      gblnMessage = False

3920      DoCmd.Close acForm, THIS_NAME

3930      DoCmd.SelectObject acForm, strCallingForm, False
3940      Select Case strCallingForm
          Case "frmJournal"
3950        Forms(strCallingForm).frmJournal_Sub3_Purchase.SetFocus
3960        Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.purchaseShareFace.SetFocus
3970      Case "frmJournal_Columns"
            ' ** Let JC_Frm_Assets() set the focus.
3980      End Select

3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
4020      Resume Next
4030    Case 3031, 3101
4040      DoCmd.RunCommand acCmdUndo
4050      Resume Next
4060    Case Else
4070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4080    End Select
4090    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4120    intShiftDown = (Shift And acShiftMask) > 0
4130    intAltDown = (Shift And acAltMask) > 0
4140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyTab
4170        With Me
4180          intRetVal = 0
4190          .cusip.SetFocus
4200        End With
4210      End Select
4220    End If

        ' ** Shift keys.
4230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4240      Select Case intRetVal
          Case vbKeyTab
4250        With Me
4260          intRetVal = 0
4270          .cmdOK.SetFocus
4280        End With
4290      End Select
4300    End If

EXITP:
4310    KeyCode = intRetVal
4320    Exit Sub

ERRH:
4330    intRetVal = 0
4340    THAT_PROC = THIS_PROC
4350    That_Erl = Erl
4360    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4370    Resume EXITP

End Sub

Private Sub cmdSave_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

4410    DoCmd.RunCommand acCmdSelectRecord
4420    DoCmd.RunCommand acCmdSaveRecord

EXITP:
4430    Exit Sub

ERRH:
4440    Select Case ERR.Number
        Case Else
4450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4460    End Select
4470    Resume EXITP

End Sub

Private Sub cusip_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4570        With Me
4580          intRetVal = 0
4590          .description_masterasset.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4650        With Me
4660          intRetVal = 0
4670          .cmdCancel.SetFocus
4680        End With
4690      End Select
4700    End If

EXITP:
4710    KeyCode = intRetVal
4720    Exit Sub

ERRH:
4730    intRetVal = 0
4740    THAT_PROC = THIS_PROC
4750    That_Erl = Erl
4760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4770    Resume EXITP

End Sub

Private Sub cusip_Exit(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cusip_Exit"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strCusip As String, strAssetDesc As String
        Dim blnContinue As Boolean

        Static intInstance As Integer

4810    blnContinue = True
4820    strCusip = vbNullString: strAssetDesc = vbNullString

4830    With Me
4840      If IsNull(.cusip) = False Then
4850        If Trim(.cusip) <> vbNullString Then

4860          Set dbs = CurrentDb
4870          With dbs
                ' ** MasterAsset, by specified [cusp].
4880            Set qdf = .QueryDefs("qryAssets_09")
4890            With qdf.Parameters
4900              ![cusp] = Me.cusip
4910            End With
4920            Set rst = qdf.OpenRecordset
4930            With rst
4940              If .BOF = True And .EOF = True Then
                    ' ** OK, nothing found.
4950                strGoodCusip = Trim(Me.cusip)
4960              Else
4970                intInstance = intInstance + 1
4980                .MoveFirst
4990                Select Case blnIsCancel
                    Case True
                      ' ** Skip the message.
5000                Case False
5010                  If intInstance <= 1 Then
5020                    blnContinue = False
5030                    strCusip = ![cusip]
5040                    strAssetDesc = ![description_masterasset]
5050                  Else
5060                    Me.cusip.Undo
5070                  End If
5080                End Select
5090                strGoodCusip = vbNullString
5100              End If
5110              .Close
5120            End With  ' ** rst.
5130            .Close
5140          End With  ' ** dbs.

5150          If blnContinue = False Then
5160            MsgBox "Cusip already exists." & Space(40) & vbCrLf & vbCrLf & strCusip & vbCrLf & _
                  strAssetDesc, vbInformation + vbOKOnly, "CUSIP Already Present"
5170            Cancel = -1
5180            .cusip.SetFocus
5190          End If

5200        End If
5210      End If
5220    End With  ' ** Me.

EXITP:
5230    Set rst = Nothing
5240    Set qdf = Nothing
5250    Set dbs = Nothing
5260    Exit Sub

ERRH:
5270    THAT_PROC = THIS_PROC
5280    That_Erl = Erl
5290    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5300    Resume EXITP

End Sub

Private Sub description_masterasset_AfterUpdate()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "description_masterasset_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String

5410    With Me
5420      If IsNull(.description_masterasset) = False Then
5430        If Trim(.description_masterasset) <> vbNullString Then
5440          strTmp01 = Trim(.description_masterasset)
5450          intPos01 = InStr(strTmp01, Chr(34))  ' ** No double-quotes allowed, so just replace with single-quote, apostrophe.
5460          If intPos01 > 0 Then
5470            Do While intPos01 > 0
5480              strTmp01 = Left(strTmp01, (intPos01 - 1)) & "'" & Mid(strTmp01, (intPos01 + 1))
5490              strTmp01 = Trim(strTmp01)
5500              intPos01 = InStr(strTmp01, Chr(34))
5510            Loop
5520            .description_masterasset = strTmp01
5530            Beep
5540          End If
5550        End If
5560      End If
5570    End With

EXITP:
5580    Exit Sub

ERRH:
5590    Select Case ERR.Number
        Case Else
5600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5610    End Select
5620    Resume EXITP

End Sub

Private Sub rate_entry_AfterUpdate()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "rate_entry_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String, dblTmp02 As Double

5710    With Me
5720      If IsNull(.rate_entry) = True Then
5730        .rate_entry = 0#
5740        .rate = 0#
5750      End If
5760      strTmp01 = CStr(.rate_entry)
5770      strTmp01 = Trim(strTmp01)
5780      If Right(strTmp01, 1) = "%" Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
5790      If IsNumeric(strTmp01) = True Then
5800        dblTmp02 = CDbl(strTmp01)  ' ** Remember, this is still a percentage, so the saved number will have 2 more decimal places!
5810        If dblTmp02 >= 100# Then
5820          MsgBox "Rate must be less than 100 percent.", vbInformation + vbOKOnly, "Invalid Entry"
5830          .rate_entry = Null
5840          .rate = Null
5850          .rate_entry.SetFocus
5860          blnSetToRate = True
5870          .TimerInterval = 100&
5880        Else
5890          intPos01 = InStr(strTmp01, ".")
5900          If intPos01 > 0 Then
                ' ** The user entered a decimal number.
5910            If Len(Mid(strTmp01, (intPos01 + 1))) > 3 Then
5920              MsgBox "You can only enter a rate percentage with up to 3 decimal places.", vbInformation + vbOKOnly, "Invalid Entry"
5930              .rate_entry = Null
5940              .rate = Null
5950              .rate_entry.SetFocus
5960              blnSetToRate = True
5970              .TimerInterval = 100&
5980            Else
5990              .rate = dblTmp02 / 100#  ' ** Here's where it gets divided by 100.
6000              .due.SetFocus
6010            End If
6020          Else
6030            If dblTmp02 <> 0# Then
6040              .rate = dblTmp02 / 100#  ' ** Here's where it gets divided by 100.
6050            Else
6060              .rate = 0#
6070            End If
6080            .due.SetFocus
6090          End If
6100        End If
6110      Else
6120        MsgBox "Please enter a valid rate percentage.", vbInformation + vbOKOnly, "Invalid Entry"
6130        .rate_entry = Null
6140        .rate = Null
6150        .rate_entry.SetFocus
6160        blnSetToRate = True
6170        .TimerInterval = 100&
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

Private Sub assettype_Enter()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_Enter"

6310    With Me
6320      If IsNull(.assettype) Then
6330        ExpandCombo  ' ** Module Procedure: modUtilities.
6340      End If
6350    End With

EXITP:
6360    Exit Sub

ERRH:
6370    THAT_PROC = THIS_PROC
6380    That_Erl = Erl
6390    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6400    Resume EXITP

End Sub

Private Sub assettype_AfterUpdate()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAssetType As String

6510    With Me
6520      If IsNull(.assettype) = True Then
6530        MsgBox "Please select a valid asset code.", vbInformation + vbOKOnly, "Entry Required"
6540        .assettype_description = Null
6550        .assettype.SetFocus
6560      Else
6570        strAssetType = .assettype
6580        Set dbs = CurrentDb
            ' ** AssetType, by specified [acod].
6590        Set qdf = dbs.QueryDefs("qryAssets_10")
6600        With qdf.Parameters
6610          ![acod] = strAssetType
6620        End With
6630        Set rst = qdf.OpenRecordset
6640        If rst.BOF = True And rst.EOF = True Then
              ' ** Not found?!
6650          MsgBox "Please select a valid asset code.", vbInformation + vbOKOnly, "Invalid Entry"
6660          .assettype_description = Null
6670          .assettype.SetFocus
6680        Else
6690          If strAssetType = "81" Then
6700            .assettype_description = "Interest Bearing"
6710          Else
6720            .assettype_description = rst![assettype_description]
6730          End If
6740        End If
6750        rst.Close
6760        dbs.Close
6770      End If
6780    End With

EXITP:
6790    Set rst = Nothing
6800    Set qdf = Nothing
6810    Set dbs = Nothing
6820    Exit Sub

ERRH:
6830    Select Case ERR.Number
        Case Else
6840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6850    End Select
6860    Resume EXITP

End Sub

Private Sub assettype_NotInList(NewData As String, Response As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_NotInList"

6910    With Me
6920      MsgBox "Please select a valid asset code.", vbInformation + vbOKOnly, "Invalid Entry"
6930      .assettype.SetFocus
6940      Response = acDataErrContinue
6950    End With

EXITP:
6960    Exit Sub

ERRH:
6970    THAT_PROC = THIS_PROC
6980    That_Erl = Erl
6990    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7000    Resume EXITP

End Sub

Private Sub assettype_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "assettype_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7120    intShiftDown = (Shift And acShiftMask) > 0
7130    intAltDown = (Shift And acAltMask) > 0
7140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7170        With Me
7180          intRetVal = 0
7190          .curr_id.SetFocus
7200        End With
7210      End Select
7220    End If

        ' ** Shift keys.
7230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7250        With Me
7260          intRetVal = 0
7270          .due.SetFocus
7280        End With
7290      End Select
7300    End If

EXITP:
7310    KeyCode = intRetVal
7320    Exit Sub

ERRH:
7330    intRetVal = 0
7340    THAT_PROC = THIS_PROC
7350    That_Erl = Erl
7360    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7370    Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim lngCurrID As Long

7410    With Me
7420      If IsNull(.curr_id) = True Then
7430        MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Entry Required"
7440        .curr_name = Null
7450        .curr_id.SetFocus
7460      Else
7470        lngCurrID = .curr_id
7480        .curr_name = .curr_id.Column(2)

7490      End If
7500    End With

EXITP:
7510    Exit Sub

ERRH:
7520    Select Case ERR.Number
        Case Else
7530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7540    End Select
7550    Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

7610    With Me
7620      MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
7630      .curr_id.SetFocus
7640      Response = acDataErrContinue
7650    End With

EXITP:
7660    Exit Sub

ERRH:
7670    THAT_PROC = THIS_PROC
7680    That_Erl = Erl
7690    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7700    Resume EXITP

End Sub

Private Sub curr_id_KeyDown(KeyCode As Integer, Shift As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_KeyDown"

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
7890          .cmdOK.SetFocus
7900        End With
7910      End Select
7920    End If

        ' ** Shift keys.
7930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7950        With Me
7960          intRetVal = 0
7970          .assettype.SetFocus
7980        End With
7990      End Select
8000    End If

EXITP:
8010    KeyCode = intRetVal
8020    Exit Sub

ERRH:
8030    intRetVal = 0
8040    THAT_PROC = THIS_PROC
8050    That_Erl = Erl
8060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8070    Resume EXITP

End Sub
