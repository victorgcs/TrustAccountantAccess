VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountProfile_Add_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountProfile_Add_Sub"

'VGC 10/26/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Fee Annually:     A {opgFeeFreq_optAnnual}
' **   AccountNo:        B {accountno}
' **   Cancel:           C {cmdCancel on frmAccountProfile_Add}
' **   Notes:            E {tabContacts_pgNotes}
' **   Bank Other:       H {opgLoc_optOther}
' **   Statement Freq:   K {cmdStatementFreq}
' **   Review Freq:      L {cmdReviewFreq}
' **   Fee Monthly:      M {opgFeeFreq_optMonth}
' **   Short Name:       N {shortname
' **   Short Name PQ:    N {shortname_pq} (Printer's Quotes)
' **   OK:               O {cmdOK on frmAccountProfile_Add}
' **   Choose RA:          (cmdChoose}
' **   Print Report:     P {cmdPrintReport on frmAccountProfile_Add}
' **   Fee Quarterly:    Q {opgFeeFreq_optQuarter}
' **   Related Accounts: R {related_accountno} (Disabled)
' **   Fee SemiAnnually: S {opgFeeFreq_optSemi}
' **   Account Type:     T {accounttype}
' **   Bank USA:         U {opgLoc_optUSA}
' **   Contact Name 1:   1 {tabContacts_pgContact1}
' **   Contact Name 2:   2 {tabContacts_pgContact2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Contact1 Other:   H {opgLocContact1_optOther}
' **   Save:             S {cmdSave}
' **   Contact1 USA:     U {opgLocContact1_optUSA}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Contact2 Other:   H {opgLocContact2_optOther}
' **   Contact2 USA:     U {opgLocContact2_optUSA}

' ** NOTE: Like Related Accounts, they can't choose a Default Asset on this form!

'WHAT ON EARTH IS 'STATE TYPE'? (statetype)

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private strCallingForm As String
Private strThisAccountNo As String, lngAssetsLbl_Offset As Long
Private blnNewHasSaved As Boolean, blnIsOpen As Boolean, blnDupe As Boolean
Private THAT_PROC As String, That_Erl As Double
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountProfile_Add" Then

160         strThisAccountNo = vbNullString
170         blnNewHasSaved = False: blnDupe = False
180         gblnClosing = False

190         lngAssetsLbl_Offset = .cmbAssets.Left - .cmbAssets_lbl.Left

200         Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

210         If IsNull(.Acct_State_Pref) = False Then
220           .Bank_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
230           .Contact1_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
240           .Contact2_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
250         End If

260         blnIsOpen = True
270         opgLoc_AfterUpdate  ' ** Procedure: Below.

280         blnIsOpen = True
290         opgLocContact1_AfterUpdate  ' ** Procedure: Below.

300         blnIsOpen = True
310         opgLocContact2_AfterUpdate  ' ** Procedure: Below.

320         If gblnSeparateCheckingAccounts = True Then
330           .Bank_Name.Visible = True
340           .Bank_City.Visible = True
350           .opgLoc.Visible = True
360           .Bank_State.Visible = True
370           .Bank_Country.Visible = True
380           .Bank_RoutingNumber.Visible = True
390           .Bank_AccountNumber.Visible = True
400           .FedIFNum1_lbl.Visible = True
410           .FedIFNum1.Visible = True
420           .FedIFNum2.Visible = True
430         Else
440           .Bank_Name.Visible = False
450           .Bank_City.Visible = False
460           .opgLoc.Enabled = False
470           .opgLoc.Visible = False
480           .opgLoc_box.Visible = False
490           .Bank_State.Visible = False
500           .Bank_Country.Visible = False
510           .Bank_RoutingNumber.Visible = False
520           .Bank_AccountNumber.Visible = False
530           .FedIFNum1_lbl.Visible = False
540           .FedIFNum1.Visible = False
550           .FedIFNum2.Visible = False
560           .cmbAssets.Top = .Bank_Name.Top
570           .cmbAssets.Left = .Bank_Name.Left
580           .cmbAssets_lbl.Top = .Bank_Name_lbl.Top
590           .cmbAssets_lbl.Left = (.cmbAssets.Left - lngAssetsLbl_Offset)
600           .LastCheckNum.Top = .cmbAssets.Top
610           .LastCheckNum_lbl.Top = .cmbAssets_lbl.Top
620         End If

630         fSetScrollBarPosVT Me, 1&  ' ** Module Function: modScrollBarFuncs.
640         .accountno.SetFocus

650       Else
660         Cancel = -1
670       End If
680     End With

690     If Cancel = -1 Then
700       Beep
710       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
720     End If

EXITP:
730     Exit Sub

ERRH:
740     THAT_PROC = THIS_PROC: That_Erl = Erl
750     Select Case ERR.Number
        Case Else
760       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770     End Select
780     Resume EXITP

End Sub

Private Sub Form_Load()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

810     With Me

820       CLR_DISABLED_FG = CLR_DKGRY
830       CLR_DISABLED_BG = CLR_LTTEAL

840       .description.ForeColor = CLR_DISABLED_FG
850       .description.BackColor = CLR_DISABLED_BG
860       .accountno2.ForeColor = CLR_DISABLED_FG
870       .accountno2.BackColor = CLR_DISABLED_BG
880       .shortname2.ForeColor = CLR_DISABLED_FG
890       .shortname2.BackColor = CLR_DISABLED_BG

900       .related_accountno.BorderColor = WIN_CLR_DISR
910       .related_accountno.BackStyle = acBackStyleTransparent
920       .cmbAssets.BorderColor = WIN_CLR_DISR
930       .cmbAssets.BackStyle = acBackStyleTransparent

940       .opgFeeFreq_opt_box.Visible = False
950       If IsNull(.feeFrequency) = False Then
960         If .feeFrequency > 0 Then
970           Select Case .feeFrequency
              Case 1
980             .opgFeeFreq_optMonth = True
990             .opgFeeFreq_opt_box.Left = .opgFeeFreq_optMonth_lbl.Left
1000            .opgFeeFreq_opt_box.Visible = True
1010          Case 2
1020            .opgFeeFreq_optQuarter = True
1030            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optQuarter_lbl.Left
1040            .opgFeeFreq_opt_box.Visible = True
1050          Case 3
1060            .opgFeeFreq_optSemi = True
1070            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optSemi_lbl.Left
1080            .opgFeeFreq_opt_box.Visible = True
1090          Case 4
1100            .opgFeeFreq_optAnnual = True
1110            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optAnnual_lbl.Left
1120            .opgFeeFreq_opt_box.Visible = True
1130          End Select
1140        End If
1150      End If

1160    End With

EXITP:
1170    Exit Sub

ERRH:
1180    THAT_PROC = THIS_PROC: That_Erl = Erl
1190    Select Case ERR.Number
        Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

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
        ' **   Notes:            E {tabContacts_pgNotes}
        ' **   Short Name:       N {shortname}
        ' **   Short Name PQ:    N {shortname_pq} (Printer's Quotes)
        ' **   Contact Name 1:   1 {tabContacts_pgContact1}
        ' **   Contact Name 2:   2 {tabContacts_pgContact2}

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   Contact1 Other:   H {opgLocContact1_optOther}
        ' **   Save:             S {cmdSave}
        ' **   Contact1 USA:     U {opgLocContact1_optUSA}

        ' ** Shortcut Ctrl-Shift to other forms and functionality.
        ' **   Contact2 Other:   H {opgLocContact2_optOther}
        ' **   Contact2 USA:     U {opgLocContact2_optUSA}

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyEscape
1370        With Me
1380          intRetVal = 0
1390          gblnClosing = True
1400          .Parent.cmdCancel_Click  ' ** Form Procedure: frmAccountProfile_Add.
1410        End With
1420      Case vbKeyPageDown
1430        With Me
1440          intRetVal = 0
1450          DoCmd.GoToPage 2, 0, 0
1460          .FocusHolder.SetFocus
1470          .Contact1_Name.SetFocus
1480          fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
1490        End With
1500      Case vbKeyPageUp
1510        With Me
1520          intRetVal = 0
1530          DoCmd.GoToPage 1, 0, 0
1540          .accountno.SetFocus
1550        End With
1560      End Select
1570    End If

        ' ** Alt keys.
1580    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1590      Select Case intRetVal
          Case vbKeyC
1600        With Me
1610          intRetVal = 0
1620          DoCmd.SelectObject acForm, .Parent.Name, False
1630          .Parent.cmdCancel.SetFocus
1640          .Parent.cmdCancel_Click
1650        End With
1660      Case vbKeyE
1670        With Me
1680          intRetVal = 0
1690          .FocusHolder.SetFocus
1700          .tabContacts = .tabContacts_pgNotes.PageIndex
1710          .Notes.SetFocus
1720          fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
1730        End With
1740      Case vbKeyN
1750        With Me
1760          intRetVal = 0
1770          If .shortname.Visible = True And .shortname.Enabled = True Then
1780            .shortname.SetFocus
1790          ElseIf .shortname_pq.Visible = True And .shortname_pq.Enabled = True Then
1800            .shortname_pq.SetFocus
1810          End If
1820        End With
1830      Case vbKey1
1840        With Me
1850          intRetVal = 0
1860          .FocusHolder.SetFocus
1870          .tabContacts = .tabContacts_pgContact1.PageIndex
1880          .Contact1_Name.SetFocus
1890          fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
1900        End With
1910      Case vbKey2
1920        With Me
1930          intRetVal = 0
1940          .FocusHolder.SetFocus
1950          .tabContacts = .tabContacts_pgContact2.PageIndex
1960          .Contact2_Name.SetFocus
1970          fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
1980        End With
1990      End Select
2000    End If

        ' ** Ctrl keys.
2010    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2020      Select Case intRetVal
          Case vbKeyH
2030        With Me
2040          If .opgLocContact1 <> .opgLocContact1_optOther.OptionValue Then
2050            .opgLocContact1 = .opgLocContact1_optOther.OptionValue
2060            opgLocContact1_AfterUpdate  ' ** Procedure: Below.
2070          End If
2080        End With
2090      Case vbKeyS
2100        intRetVal = 0
2110        cmdSave_Click  ' ** Procedure: Below.
2120      Case vbKeyU
2130        With Me
2140          If .opgLocContact1 <> .opgLocContact1_optUSA.OptionValue Then
2150            .opgLocContact1 = .opgLocContact1_optUSA.OptionValue
2160            opgLocContact1_AfterUpdate  ' ** Procedure: Below.
2170          End If
2180        End With
2190      Case vbKeyTab, vbKeyReturn
2200        With Me
2210          intRetVal = 0
2220          DoCmd.SelectObject acForm, .Parent.Name, False
2230          If .Parent.cmdPrintReport.Enabled = True Then
2240            .Parent.cmdPrintReport.SetFocus
2250          ElseIf .Parent.cmdOK.Enabled = True Then
2260            .Parent.cmdOK.SetFocus
2270          Else
2280            .Parent.cmdCancel.SetFocus
2290          End If
2300        End With
2310      End Select
2320    End If

        ' ** Ctrl-Shift keys.
2330    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyH
2350        With Me
2360          If .opgLocContact2 <> .opgLocContact2_optOther.OptionValue Then
2370            .opgLocContact2 = .opgLocContact2_optOther.OptionValue
2380            opgLocContact2_AfterUpdate  ' ** Procedure: Below.
2390          End If
2400        End With
2410      Case vbKeyU
2420        With Me
2430          If .opgLocContact2 <> .opgLocContact2_optUSA.OptionValue Then
2440            .opgLocContact2 = .opgLocContact2_optUSA.OptionValue
2450            opgLocContact2_AfterUpdate  ' ** Procedure: Below.
2460          End If
2470        End With
2480      Case vbKeyTab, vbKeyReturn
2490        With Me
2500          intRetVal = 0
2510          DoCmd.SelectObject acForm, .Parent.Name, False
2520          .Parent.cmdCancel.SetFocus
2530        End With
2540      End Select
2550    End If

EXITP:
2560    KeyCode = intRetVal
2570    Exit Sub

ERRH:
2580    intRetVal = 0
2590    THAT_PROC = THIS_PROC: That_Erl = Erl
2600    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
2610    Case Else
2620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2630    End Select
2640    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

        Dim msgResponse As VbMsgBoxResult

2710    If DataErr <> 0 Then
2720      Select Case DataErr
          Case 3164  ' ** Field cannot be updated.
            ' ** I have no idea what this is about!
2730        Response = acDataErrContinue
2740      Case 3314  ' ** The field '|' cannot contain a Null value because the Required property for this field is set to True.
2750        Select Case gblnClosing
            Case True
2760          Response = acDataErrContinue
2770        Case False
2780          With Me
2790            If IsNull(.accountno) = True And IsNull(.shortname) = True Then
2800              msgResponse = MsgBox("Do you wish to abandon this New Account?", vbQuestion + vbYesNo, "Abandon New Entry")
2810            Else
2820              msgResponse = vbNo
2830            End If
2840            If msgResponse <> vbYes Then
2850              MsgBox "A valid Account Number and Short Name are required to continue.", vbInformation + vbOKOnly, "Entry Required"
2860              Response = acDataErrContinue
2870            Else
2880              gblnClosing = True
2890              .Parent.cmdCancel_Click  ' ** Form Procedure: frmAccountProfile_Add.
2900              Response = acDataErrContinue
2910            End If
2920          End With
2930        End Select
2940      Case Else
2950        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2960        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970      End Select
2980    End If

EXITP:
2990    Exit Sub

ERRH:
3000    THAT_PROC = THIS_PROC: That_Erl = Erl
3010    Select Case ERR.Number
        Case Else
3020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3030    End Select
3040    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** Doesn't work from here, so moved to parent.
        ' ** Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3110    If strThisAccountNo <> vbNullString Then
3120      gstrActNo = strThisAccountNo
3130    End If

EXITP:
3140    Exit Sub

ERRH:
3150    THAT_PROC = THIS_PROC: That_Erl = Erl
3160    Select Case ERR.Number
        Case Else
3170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3180    End Select
3190    Resume EXITP

End Sub

Public Sub cmdSave_Click()
' ** This is a fake click event.

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3210    With Me
3220      If IsNull(.accountno) = False And IsNull(.shortname) = False Then

3230        DoCmd.RunCommand acCmdSaveRecord
3240        DoEvents

3250        If strThisAccountNo = vbNullString Then
3260          blnNewHasSaved = True
              '.Parent.HasSaved_Set strThisAccountNo, blnNewHasSaved  ' ** Form Procedure: frmAccountProfile.
3270        End If

3280      Else
            ' ** Nothing, for now.
3290      End If
3300    End With  ' ** Me.

EXITP:
3310    Exit Sub

ERRH:
3320    THAT_PROC = THIS_PROC: That_Erl = Erl
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub accountno_BeforeUpdate(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_BeforeUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strTmp01 As String, strTmp02 As String

3410    With Me
3420      blnDupe = False
3430      If IsNull(.accountno) = True Then
3440        If Trim(.accountno.text) <> vbNullString Then
3450          strTmp01 = Trim(.accountno.text)
3460        Else
3470          strTmp01 = vbNullString
3480        End If
3490      Else
3500        strTmp01 = .accountno
3510      End If
3520      If strTmp01 <> vbNullString Then
            ' ** Remove spaces in accountno.
3530        If InStr(strTmp01, " ") > 0 Then
3540          strTmp02 = StringReplace(strTmp01, " ", "")  ' ** Module Function: modUtilities.
3550  On Error Resume Next
3560          .accountno = strTmp02
3570          If ERR.Number <> 0 Then
                ' ** -2147352567  The macro or function set to the BeforeUpdate or ValidationRule property for this field is preventing Microsoft Access from saving the data in the field.
3580            .accountno.text = strTmp02
3590  On Error GoTo ERRH
3600          Else
3610  On Error GoTo ERRH
3620          End If
3630          strTmp01 = strTmp02
3640          MsgBox "The Account Number cannot have spaces." & vbCrLf & "They have been removed.", vbInformation + vbOKOnly, "Spaces Removed"
3650        End If
3660        If gblnAccountNoWithType = True Then
              ' ** Verify the user entered a valid account type prefix to the account number.
3670          Set dbs = CurrentDb
3680          If InStr(strTmp01, "-") > 0 Then
3690            If Len(strTmp01) > 15 Then
3700              Beep
3710              MsgBox "The Account Number is limited to 15 characters.", vbInformation + vbOKOnly, "Invalid Entry"
3720              .accountno.Undo
3730              Cancel = -1
3740            Else
3750              strTmp02 = Left(strTmp01, 2)
                  ' ** AccountType, by specified [acttyp].
3760              Set qdf = dbs.QueryDefs("qryAccountProfile_13")
3770              With qdf.Parameters
3780                ![acttyp] = strTmp02
3790              End With
3800              Set rst = qdf.OpenRecordset
3810              If rst.BOF = True And rst.EOF = True Then
3820                rst.Close
3830                Beep
3840                MsgBox "The Account Type you entered as an Account Number prefix does not exist." & vbCrLf & vbCrLf & _
                      "Please enter a valid Account Type to continue.", vbInformation + vbOKOnly, "Invalid Entry"
3850                .accountno.Undo
3860                Cancel = -1
3870              Else
3880                rst.Close
3890                .accounttype = strTmp02
3900                .accounttype.Requery
3910              End If
3920            End If
3930          Else
3940            If Len(strTmp01) > 12 Then
3950              Beep
3960              MsgBox "The Account Number is limited to 12 characters" & vbCrLf & _
                    "without the Account Type prefix.", vbInformation + vbOKOnly, "Invalid Entry"
3970              .accountno.Undo
3980              Cancel = -1
3990            Else
4000              If IsNull(.accounttype) = False Then
                    ' ** Add the chosen Account Type to accountno.
4010                strTmp02 = .accounttype & "-" & strTmp01
                    ' ** Account, by specified [actno].
4020                Set qdf = dbs.QueryDefs("qryAccountProfile_05")
4030                With qdf.Parameters
4040                  ![actno] = strTmp02
4050                End With
4060                Set rst = qdf.OpenRecordset
4070                If rst.BOF = True And rst.EOF = True Then
                      ' ** No hit, it's new.
4080                  rst.Close
4090                  .accountno = strTmp02
4100                  strTmp01 = strTmp02
4110                Else
                      ' ** The account number already exists.
4120                  rst.MoveFirst
4130                  Beep
4140                  MsgBox "An account with this number already exists." & vbCrLf & _
                        "   " & rst![accountno] & "   " & rst![shortname], vbInformation + vbOKOnly, "Account Exists"
4150                  rst.Close
4160                  blnDupe = True
4170                  .accountno.Undo
4180                  .accountno.Undo
4190                  .Undo
4200                  .Undo
4210                  Cancel = -1
4220                End If
4230              Else
                    ' ** They haven't yet chosen an Account Type.
                    ' ** Account, for AccountNoWithType, just non-type portion, by specified [actno].
4240                Set qdf = dbs.QueryDefs("qryAccountProfile_14")
4250                With qdf.Parameters
4260                  ![actno] = strTmp01
4270                End With
4280                Set rst = qdf.OpenRecordset
4290                If rst.BOF = True And rst.EOF = True Then
                      ' ** No hit, it's new.
4300                  rst.Close
4310                Else
4320                  rst.MoveFirst
4330                  strTmp02 = rst![accountno]
4340                  If InStr(strTmp02, "-") > 0 Then
4350                    strTmp02 = Left(strTmp02, 2)
4360                    Beep
4370                    MsgBox "Though you haven't yet chosen an Account Type for this new account," & vbCrLf & _
                          "an account exists with the Account Type prefix " & strTmp02 & ":" & vbCrLf & _
                          "   " & rst![accountno] & "   " & rst![shortname], vbInformation + vbOKOnly, "Account Exists"
4380                    rst.Close
                        ' ** Let it go through for now.
4390                  Else
4400                    Beep
4410                    MsgBox "An account with this number already exists." & vbCrLf & _
                          "   " & rst![accountno] & "   " & rst![shortname], vbInformation + vbOKOnly, "Account Exists"
4420                    rst.Close
4430                    blnDupe = True
4440                    .accountno.Undo
4450                    .accountno.Undo
4460                    .Undo
4470                    .Undo
4480                    Cancel = -1
4490                  End If
4500                End If
4510              End If
4520            End If
4530          End If
4540          dbs.Close
4550        Else
4560          If Len(strTmp01) > 15 Then
4570            Beep
4580            MsgBox "The Account Number is limited to 15 characters.", vbInformation + vbOKOnly, "Invalid Entry"
4590            .accountno.Undo
4600            Cancel = -1
4610          End If
4620        End If
4630        If Cancel = 0 Then
              ' ** Verify the user entered a unique account number.
4640          Set dbs = CurrentDb
              ' ** Account, by specified [actno].
4650          Set qdf = dbs.QueryDefs("qryAccountProfile_05")
4660          With qdf.Parameters
4670            ![actno] = strTmp01
4680          End With
4690          Set rst = qdf.OpenRecordset
4700          If rst.BOF = True And rst.EOF = True Then
                ' ** No hit, it's new.
4710            rst.Close
4720          Else
                ' ** The account number already exists.
4730            Beep
4740            MsgBox "An account with this number already exists." & vbCrLf & _
                  "   " & rst![accountno] & "   " & rst![shortname], vbInformation + vbOKOnly, "Account Exists"
4750            rst.Close
4760            blnDupe = True
4770            .accountno.Undo
4780            .accountno.Undo
4790            .Undo
4800            .Undo
4810            Cancel = -1
4820          End If
4830          dbs.Close
4840        End If
4850      End If
4860    End With

EXITP:
4870    Set rst = Nothing
4880    Set qdf = Nothing
4890    Set dbs = Nothing
4900    Exit Sub

ERRH:
4910    THAT_PROC = THIS_PROC: That_Erl = Erl
4920    Select Case ERR.Number
        Case Else
4930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4940    End Select
4950    Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5020    intShiftDown = (Shift And acShiftMask) > 0
5030    intAltDown = (Shift And acAltMask) > 0
5040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5070        With Me
5080          intRetVal = 0
5090          If blnDupe = False Then
5100            If .shortname.Enabled = True Then
5110              .shortname.SetFocus
5120            Else
                  ' ** I don't know what this is about.
5130            End If
5140          End If
5150        End With
5160      End Select
5170    End If

        ' ** Shift keys.
5180    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5190      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5200        With Me
5210          intRetVal = 0
5220          If blnDupe = False Then
5230            DoCmd.SelectObject acForm, .Parent.Name, False
5240            .Parent.cmdCancel.SetFocus
5250          End If
5260        End With
5270      End Select
5280    End If

EXITP:
5290    KeyCode = intRetVal
5300    Exit Sub

ERRH:
5310    intRetVal = 0
5320    THAT_PROC = THIS_PROC: That_Erl = Erl
5330    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
5340    Case Else
5350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5360    End Select
5370    Resume EXITP

End Sub

Private Sub accountno_Exit(Cancel As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_Exit"

5410    With Me
5420      If blnDupe = False Then
5430        If IsNull(.accountno) = False And IsNull(.accounttype) = False And IsNull(.shortname) = False Then
5440          If .Dirty = True Then
5450            cmdSave_Click  ' ** Procedure: Above.
5460          End If
5470          .Parent.cmdOK.Enabled = True
5480          If .Parent.cmdPrintReport.Enabled = False Then
5490            .Parent.cmdPrintReport.Enabled = True
5500            .Parent.cmdPrintReport_raised_img.Visible = True
5510            .Parent.cmdPrintReport_raised_semifocus_dots_img.Visible = False
5520            .Parent.cmdPrintReport_raised_focus_img.Visible = False
5530            .Parent.cmdPrintReport_raised_focus_dots_img.Visible = False
5540            .Parent.cmdPrintReport_sunken_focus_dots_img.Visible = False
5550            .Parent.cmdPrintReport_raised_img_dis.Visible = False
5560          End If
5570        Else
5580          .Parent.cmdOK.Enabled = False
5590          If .Parent.cmdPrintReport.Enabled = True Then
5600            .Parent.cmdPrintReport.Enabled = False
5610            .Parent.cmdPrintReport_raised_img_dis.Visible = True
5620            .Parent.cmdPrintReport_raised_img.Visible = False
5630            .Parent.cmdPrintReport_raised_semifocus_dots_img.Visible = False
5640            .Parent.cmdPrintReport_raised_focus_img.Visible = False
5650            .Parent.cmdPrintReport_raised_focus_dots_img.Visible = False
5660            .Parent.cmdPrintReport_sunken_focus_dots_img.Visible = False
5670            .Parent.cmdPrintReport_raised_img_dis.Visible = False
5680          End If
5690        End If
5700      End If
5710    End With

EXITP:
5720    Exit Sub

ERRH:
5730    THAT_PROC = THIS_PROC: That_Erl = Erl
5740    Select Case ERR.Number
        Case Else
5750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5760    End Select
5770    Resume EXITP

End Sub

Private Sub related_accountno_lbl4_GotFocus()
' ** Text box.
' ** This is the small one, just holding focus.
' ** related_accountno_lbl2 is just for the white background and border.

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "related_accountno_lbl4_GotFocus"

5810    With Me
5820      .related_accountno_lbl3.Visible = True  ' ** This is the one with 'Choose -->'.
5830      .related_accountno_lbl2.Visible = True  ' ** This is the white background and border.
5840      .related_accountno.Visible = False
5850      blnDupe = False
5860    End With

EXITP:
5870    Exit Sub

ERRH:
5880    THAT_PROC = THIS_PROC: That_Erl = Erl
5890    Select Case ERR.Number
        Case Else
5900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5910    End Select
5920    Resume EXITP

End Sub

Private Sub related_accountno_lbl4_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Text box.

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "related_accountno_lbl4_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6020    intShiftDown = (Shift And acShiftMask) > 0
6030    intAltDown = (Shift And acAltMask) > 0
6040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6070        With Me
6080          intRetVal = 0
6090          If IsNull(.accounttype) = True Or IsNull(.shortname) = True Then
6100            .shortname.SetFocus
6110          Else
6120            .cmdChoose.SetFocus
6130          End If
6140        End With
6150      End Select
6160    End If

        ' ** Shift keys.
6170    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6180      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6190        With Me
6200          intRetVal = 0
6210          .accountno.SetFocus
6220        End With
6230      End Select
6240    End If

EXITP:
6250    KeyCode = intRetVal
6260    Exit Sub

ERRH:
6270    intRetVal = 0
6280    THAT_PROC = THIS_PROC: That_Erl = Erl
6290    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
6300    Case Else
6310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6320    End Select
6330    Resume EXITP

End Sub

Private Sub related_accountno_lbl4_LostFocus()
' ** Text box.

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "related_accountno_lbl4_LostFocus"

6410    With Me
6420      .related_accountno_lbl3.Visible = False  ' ** This is the one with 'Choose -->'.
6430      .related_accountno_lbl2.Visible = False  ' ** This is the white background and border.
6440      .related_accountno.Visible = True
6450    End With

EXITP:
6460    Exit Sub

ERRH:
6470    THAT_PROC = THIS_PROC: That_Erl = Erl
6480    Select Case ERR.Number
        Case Else
6490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6500    End Select
6510    Resume EXITP

End Sub

Public Sub cmdChoose_Click()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChoose_Click"

        Dim strAccts As String

6610    With Me

6620      If IsNull(.accounttype) = True Then
6630        MsgBox "Please choose an Account Type before selecting Related Accounts.", vbInformation + vbOKOnly, "Entry Required"
6640      ElseIf IsNull(.shortname) = True Then
6650        MsgBox "Please enter the account's Short Name before selecting Related Accounts.", vbInformation + vbOKOnly, "Entry Required"
6660      Else

6670        If IsNull(.related_accountno) = False Then
6680          If Trim(.related_accountno) <> vbNullString Then
6690            strAccts = Trim(.related_accountno)
6700          Else
6710            strAccts = vbNullString
6720          End If
6730        Else
6740          strAccts = vbNullString
6750        End If

6760        cmdSave_Click ' ** Procedure: Above.

6770        DoCmd.OpenForm "frmAccountProfile_RelAccts", acNormal, , , , acDialog, strCallingForm & "~" & strAccts

6780        DoEvents

            ' ** Make sure this form is on top.
6790        DoCmd.SelectObject acForm, strCallingForm, False

6800      End If

6810    End With  ' ** Me.

EXITP:
6820    Exit Sub

ERRH:
6830    THAT_PROC = THIS_PROC: That_Erl = Erl
6840    Select Case ERR.Number
        Case Else
6850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6860    End Select
6870    Resume EXITP

End Sub

Private Sub cmdChoose_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChoose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6920    intShiftDown = (Shift And acShiftMask) > 0
6930    intAltDown = (Shift And acAltMask) > 0
6940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6960      Select Case intRetVal
          Case vbKeyTab
6970        With Me
6980          intRetVal = 0
6990          .shortname.SetFocus
7000        End With
7010      Case vbKeyReturn
7020        With Me
7030          If IsNull(.accounttype) = True Or IsNull(.shortname) = True Then
7040            intRetVal = 0
7050            .shortname.SetFocus
7060          Else
                ' ** Let it go through to invoke cmdChoose_Click.
7070          End If
7080        End With
7090      End Select
7100    End If

        ' ** Shift keys.
7110    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7120      Select Case intRetVal
          Case vbKeyTab
7130        With Me
7140          intRetVal = 0
7150          If IsNull(.related_accountno) = True And .related_accountno_lbl4.Visible = True Then
7160            .related_accountno_lbl4.SetFocus
7170          ElseIf .related_accountno_lbl4.Visible = False Then
7180            .related_accountno.SetFocus
7190          End If
7200        End With
7210      End Select
7220    End If

EXITP:
7230    KeyCode = intRetVal
7240    Exit Sub

ERRH:
7250    intRetVal = 0
7260    THAT_PROC = THIS_PROC: That_Erl = Erl
7270    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
7280    Case Else
7290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7300    End Select
7310    Resume EXITP

End Sub

Private Sub shortname_AfterUpdate()

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_AfterUpdate"

        Dim strTmp01 As String, strTmp02 As String

7410    With Me
7420      blnDupe = False
7430      If IsNull(.shortname) = False Then
7440        If Trim(.shortname) <> vbNullString Then
7450          strTmp01 = Trim(.shortname)
7460          strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
7470          If strTmp02 <> strTmp01 Then
7480            .shortname = strTmp02
7490            .shortname_pq.Visible = True
7500            .shortname_pq.Enabled = True
7510            .CaseNum.SetFocus
7520            .shortname.Visible = False
7530            .shortname.Enabled = False
7540          End If
7550        End If
7560      End If
7570    End With

EXITP:
7580    Exit Sub

ERRH:
7590    THAT_PROC = THIS_PROC: That_Erl = Erl
7600    Select Case ERR.Number
        Case Else
7610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7620    End Select
7630    Resume EXITP

End Sub

Private Sub shortname_Exit(Cancel As Integer)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_Exit"

7710    With Me
7720      If IsNull(.accountno) = False And IsNull(.shortname) = False And IsNull(.accounttype) = False Then
7730        If .Dirty = True Then
7740          cmdSave_Click  ' ** Procedure: Above.
7750        End If
7760        .Parent.cmdOK.Enabled = True
7770        If .Parent.cmdPrintReport.Enabled = False Then
7780          .Parent.cmdPrintReport.Enabled = True
7790          .Parent.cmdPrintReport_raised_img.Visible = True
7800          .Parent.cmdPrintReport_raised_semifocus_dots_img.Visible = False
7810          .Parent.cmdPrintReport_raised_focus_img.Visible = False
7820          .Parent.cmdPrintReport_raised_focus_dots_img.Visible = False
7830          .Parent.cmdPrintReport_sunken_focus_dots_img.Visible = False
7840          .Parent.cmdPrintReport_raised_img_dis.Visible = False
7850        End If
7860      Else
7870        .Parent.cmdOK.Enabled = False
7880        If .Parent.cmdPrintReport.Enabled = True Then
7890          .Parent.cmdPrintReport.Enabled = False
7900          .Parent.cmdPrintReport_raised_img_dis.Visible = True
7910          .Parent.cmdPrintReport_raised_img.Visible = False
7920          .Parent.cmdPrintReport_raised_semifocus_dots_img.Visible = False
7930          .Parent.cmdPrintReport_raised_focus_img.Visible = False
7940          .Parent.cmdPrintReport_raised_focus_dots_img.Visible = False
7950          .Parent.cmdPrintReport_sunken_focus_dots_img.Visible = False
7960          .Parent.cmdPrintReport_raised_img_dis.Visible = False
7970        End If
7980      End If
7990    End With

EXITP:
8000    Exit Sub

ERRH:
8010    THAT_PROC = THIS_PROC: That_Erl = Erl
8020    Select Case ERR.Number
        Case Else
8030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8040    End Select
8050    Resume EXITP

End Sub

Private Sub shortname_pq_AfterUpdate()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_pq_AfterUpdate"

        Dim strTmp01 As String, strTmp02 As String

        ' ** Single-quote, double-quote, printer's quote, printers quote, printers' quotes.
        ' **  '  Chr(39)
        ' **  "  Chr(34)
        ' **  ‘  Chr(145)
        ' **  ’  Chr(146)
        ' **  “  Chr(147)
        ' **  ”  Chr(148)

8110    With Me
8120      If IsNull(.shortname_pq) = False Then
8130        If Trim(.shortname_pq) <> vbNullString Then
8140          strTmp01 = Trim(.shortname_pq)
8150          If InStr(strTmp01, Chr(145)) = 0 And InStr(strTmp01, Chr(146)) = 0 And _
                  InStr(strTmp01, Chr(147)) = 0 And InStr(strTmp01, Chr(148)) = 0 Then
                ' ** They've removed the printers quotes entirely.
8160            strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
8170            If strTmp02 <> strTmp01 Then
                  ' ** And replaced them with regular quotes! Put them back.
8180              .shortname_pq = strTmp02
8190              .CaseNum.SetFocus
8200            Else
8210              .shortname.Visible = True
8220              .shortname.Enabled = True
8230              .CaseNum.SetFocus
8240              .shortname_pq.Visible = False
8250              .shortname_pq.Enabled = False
8260            End If
8270          Else
8280            strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
8290            If strTmp02 <> strTmp01 Then
8300              .shortname_pq = strTmp02
8310              .CaseNum.SetFocus
8320            End If
8330          End If
8340        End If
8350      End If
8360    End With

EXITP:
8370    Exit Sub

ERRH:
8380    THAT_PROC = THIS_PROC: That_Erl = Erl
8390    Select Case ERR.Number
        Case Else
8400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8410    End Select
8420    Resume EXITP

End Sub

Private Sub shortname_pq_Exit(Cancel As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_pq_Exit"

8510    With Me
8520      If IsNull(.accountno) = False And IsNull(.shortname_pq) = False And IsNull(.accounttype) = False Then
8530        If .Dirty = True Then
8540          cmdSave_Click  ' ** Procedure: Above.
8550        End If
8560        .Parent.cmdOK.Enabled = True
8570      Else
8580        .Parent.cmdOK.Enabled = False
8590      End If
8600    End With

EXITP:
8610    Exit Sub

ERRH:
8620    THAT_PROC = THIS_PROC: That_Erl = Erl
8630    Select Case ERR.Number
        Case Else
8640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8650    End Select
8660    Resume EXITP

End Sub

Private Sub legalname_AfterUpdate()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_AfterUpdate"

        Dim strTmp01 As String, strTmp02 As String

8710    With Me
8720      If IsNull(.legalname) = False Then
8730        If Trim(.legalname) <> vbNullString Then
8740          strTmp01 = Trim(.legalname)
8750          strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
8760          If strTmp02 <> strTmp01 Then
8770            .legalname = strTmp02
8780            .legalname_pq.Visible = True
8790            .legalname_pq.Enabled = True
8800            .cotrustee.SetFocus
8810            .legalname.Visible = False
8820            .legalname.Enabled = False
8830          End If
8840        End If
8850      End If
8860    End With

EXITP:
8870    Exit Sub

ERRH:
8880    THAT_PROC = THIS_PROC: That_Erl = Erl
8890    Select Case ERR.Number
        Case Else
8900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8910    End Select
8920    Resume EXITP

End Sub

Private Sub legalname_pq_AfterUpdate()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_pq_AfterUpdate"

        Dim strTmp01 As String, strTmp02 As String

        ' ** Single-quote, double-quote, printer's quote, printers quote, printers' quotes.
        ' **  '  Chr(39)
        ' **  "  Chr(34)
        ' **  ‘  Chr(145)
        ' **  ’  Chr(146)
        ' **  “  Chr(147)
        ' **  ”  Chr(148)

9010    With Me
9020      If IsNull(.legalname_pq) = False Then
9030        If Trim(.legalname_pq) <> vbNullString Then
9040          strTmp01 = Trim(.legalname_pq)
9050          If InStr(strTmp01, Chr(145)) = 0 And InStr(strTmp01, Chr(146)) = 0 And _
                  InStr(strTmp01, Chr(147)) = 0 And InStr(strTmp01, Chr(148)) = 0 Then
                ' ** They've removed the printers quotes entirely.
9060            strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
9070            If strTmp02 <> strTmp01 Then
                  ' ** And replaced them with regular quotes! Put them back.
9080              .legalname_pq = strTmp02
9090              .cotrustee.SetFocus
9100            Else
9110              .legalname.Visible = True
9120              .legalname.Enabled = True
9130              .cotrustee.SetFocus
9140              .legalname_pq.Visible = False
9150              .legalname_pq.Enabled = False
9160            End If
9170          Else
9180            strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
9190            If strTmp02 <> strTmp01 Then
9200              .legalname_pq = strTmp02
9210              .cotrustee.SetFocus
9220            End If
9230          End If
9240        End If
9250      End If
9260    End With

EXITP:
9270    Exit Sub

ERRH:
9280    THAT_PROC = THIS_PROC: That_Erl = Erl
9290    Select Case ERR.Number
        Case Else
9300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9310    End Select
9320    Resume EXITP

End Sub

Private Sub cotrustee_Enter()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cotrustee_Enter"

9410    With Me
9420      If .cotrustee.text = vbNullString Then
9430        ExpandCombo .cotrustee  ' ** Module Procedure: modUtilities.
9440      End If
9450    End With

EXITP:
9460    Exit Sub

ERRH:
9470    THAT_PROC = THIS_PROC: That_Erl = Erl
9480    Select Case ERR.Number
        Case Else
9490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9500    End Select
9510    Resume EXITP

End Sub

Private Sub cotrustee_NotInList(NewData As String, Response As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cotrustee_NotInList"

9610    With Me
9620      .cotrustee = Null
9630    End With

EXITP:
9640    Exit Sub

ERRH:
9650    THAT_PROC = THIS_PROC: That_Erl = Erl
9660    Select Case ERR.Number
        Case Else
9670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9680    End Select
9690    Resume EXITP

End Sub

Private Sub investmentobj_Enter()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "investmentobj_Enter"

9710    With Me
9720      If IsNull(.investmentobj) Then
9730        ExpandCombo .investmentobj  ' ** Module Procedure: modUtilities.
9740      End If
9750    End With

EXITP:
9760    Exit Sub

ERRH:
9770    THAT_PROC = THIS_PROC: That_Erl = Erl
9780    Select Case ERR.Number
        Case Else
9790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9800    End Select
9810    Resume EXITP

End Sub

Private Sub accounttype_Enter()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_Enter"

9910    With Me
9920      If .accounttype.text = vbNullString Then
9930        ExpandCombo .accounttype  ' ** Module Procedure: modUtilities.
9940      End If
9950    End With

EXITP:
9960    Exit Sub

ERRH:
9970    THAT_PROC = THIS_PROC: That_Erl = Erl
9980    Select Case ERR.Number
        Case Else
9990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10000   End Select
10010   Resume EXITP

End Sub

Private Sub accounttype_NotInList(NewData As String, Response As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_NotInList"

10110   With Me
10120     .accounttype = Null
10130     MsgBox "Please select a valid account type.", vbInformation + vbOKOnly, "Invalid Entry"
10140     Response = acDataErrContinue
10150   End With

EXITP:
10160   Exit Sub

ERRH:
10170   THAT_PROC = THIS_PROC: That_Erl = Erl
10180   Select Case ERR.Number
        Case Else
10190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10200   End Select
10210   Resume EXITP

End Sub

Private Sub accounttype_AfterUpdate()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_AfterUpdate"

10310   With Me
10320     If IsNull(.accounttype) = False Then
10330       If gblnAccountNoWithType = True Then
              ' ** Verify the user entered a valid account type prefix to the account number.
10340         If InStr(.accountno, "-") > 0 Then
10350           If Mid(.accountno, 3, 1) = "-" Then
10360             If Left(.accountno, 2) <> .accounttype Then
10370               .accountno = .accounttype & Mid(accountno, 3)
10380             End If
10390           Else
10400             .accountno = .accounttype & "-" & .accountno
10410           End If
10420         Else
10430           .accountno = .accounttype & "-" & Trim(.accountno)
10440         End If
10450       End If
10460     End If
10470     If IsNull(.accountno) = False And IsNull(.shortname) = False And IsNull(.accounttype) = False Then
10480       cmdSave_Click  ' ** Procedure: Above.
10490       .Parent.cmdOK.Enabled = True
10500     Else
10510       .Parent.cmdOK.Enabled = False
10520     End If
10530   End With

EXITP:
10540   Exit Sub

ERRH:
10550   THAT_PROC = THIS_PROC: That_Erl = Erl
10560   Select Case ERR.Number
        Case Else
10570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10580   End Select
10590   Resume EXITP

End Sub

Private Sub accounttype_Exit(Cancel As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_Exit"

10610   With Me
10620     If IsNull(.accountno) = False And IsNull(.shortname) = False And IsNull(.accounttype) = False Then
10630       If .Dirty = True Then
10640         cmdSave_Click  ' ** Procedure: Above.
10650       End If
10660       .Parent.cmdOK.Enabled = True
10670       If .Parent.cmdPrintReport.Enabled = False Then
10680         .Parent.cmdPrintReport.Enabled = True
10690         .Parent.cmdPrintReport_raised_img.Visible = True
10700         .Parent.cmdPrintReport_raised_semifocus_dots_img.Visible = False
10710         .Parent.cmdPrintReport_raised_focus_img.Visible = False
10720         .Parent.cmdPrintReport_raised_focus_dots_img.Visible = False
10730         .Parent.cmdPrintReport_sunken_focus_dots_img.Visible = False
10740         .Parent.cmdPrintReport_raised_img_dis.Visible = False
10750       End If
10760     Else
10770       .Parent.cmdOK.Enabled = False
10780       If .Parent.cmdPrintReport.Enabled = True Then
10790         .Parent.cmdPrintReport.Enabled = False
10800         .Parent.cmdPrintReport_raised_img_dis.Visible = True
10810         .Parent.cmdPrintReport_raised_img.Visible = False
10820         .Parent.cmdPrintReport_raised_semifocus_dots_img.Visible = False
10830         .Parent.cmdPrintReport_raised_focus_img.Visible = False
10840         .Parent.cmdPrintReport_raised_focus_dots_img.Visible = False
10850         .Parent.cmdPrintReport_sunken_focus_dots_img.Visible = False
10860         .Parent.cmdPrintReport_raised_img_dis.Visible = False
10870       End If
10880     End If
10890   End With

EXITP:
10900   Exit Sub

ERRH:
10910   THAT_PROC = THIS_PROC: That_Erl = Erl
10920   Select Case ERR.Number
        Case Else
10930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10940   End Select
10950   Resume EXITP

End Sub

Private Sub cmbSweep_AfterUpdate()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbSweep_AfterUpdate"

11010   With Me
11020     If .cmbSweep.text = "Yes" Then
11030       .account_SWEEP.Value = True
11040     Else
11050       .account_SWEEP.Value = False
11060     End If
11070   End With

EXITP:
11080   Exit Sub

ERRH:
11090   THAT_PROC = THIS_PROC: That_Erl = Erl
11100   Select Case ERR.Number
        Case Else
11110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11120   End Select
11130   Resume EXITP

End Sub

Private Sub documentdate_BeforeUpdate(Cancel As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "documentdate_BeforeUpdate"

11210   With Me
11220     If .documentdate.text <> "__/__/____" And .documentdate.text <> vbNullString Then
11230       If Left(.documentdate.text, 2) > 12 Then
11240         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
11250         Cancel = -1
11260       End If
11270     End If
11280   End With

EXITP:
11290   Exit Sub

ERRH:
11300   THAT_PROC = THIS_PROC: That_Erl = Erl
11310   Select Case ERR.Number
        Case Else
11320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11330   End Select
11340   Resume EXITP

End Sub

Private Sub amendments_Enter()

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "amendments_Enter"

11410   With Me
11420     If .amendments.text = vbNullString Then
11430       ExpandCombo .amendments  ' ** Module Procedure: modUtilities.
11440     End If
11450   End With

EXITP:
11460   Exit Sub

ERRH:
11470   THAT_PROC = THIS_PROC: That_Erl = Erl
11480   Select Case ERR.Number
        Case Else
11490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11500   End Select
11510   Resume EXITP

End Sub

Private Sub amendments_NotInList(NewData As String, Response As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "amendments_NotInList"

11610   With Me
11620     .amendments = Null
11630   End With

EXITP:
11640   Exit Sub

ERRH:
11650   THAT_PROC = THIS_PROC: That_Erl = Erl
11660   Select Case ERR.Number
        Case Else
11670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11680   End Select
11690   Resume EXITP

End Sub

Private Sub appointmentdate_BeforeUpdate(Cancel As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "appointmentdate_BeforeUpdate"

11710   With Me
11720     If .appointmentdate.text <> "__/__/____" And .appointmentdate.text <> vbNullString Then
11730       If Left(.appointmentdate.text, 2) > 12 Then
11740         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
11750         Cancel = -1
11760       End If
11770     End If
11780   End With

EXITP:
11790   Exit Sub

ERRH:
11800   THAT_PROC = THIS_PROC: That_Erl = Erl
11810   Select Case ERR.Number
        Case Else
11820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11830   End Select
11840   Resume EXITP

End Sub

Private Sub courtsupervised_Enter()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "courtsupervised_Enter"

11910   With Me
11920     If .courtsupervised.text = vbNullString Then
11930       ExpandCombo .courtsupervised  ' ** Module Procedure: modUtilities.
11940     End If
11950   End With

EXITP:
11960   Exit Sub

ERRH:
11970   THAT_PROC = THIS_PROC: That_Erl = Erl
11980   Select Case ERR.Number
        Case Else
11990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12000   End Select
12010   Resume EXITP

End Sub

Private Sub courtsupervised_NotInList(NewData As String, Response As Integer)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "courtsupervised_NotInList"

12110   With Me
12120     .courtsupervised = Null
12130   End With

EXITP:
12140   Exit Sub

ERRH:
12150   THAT_PROC = THIS_PROC: That_Erl = Erl
12160   Select Case ERR.Number
        Case Else
12170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12180   End Select
12190   Resume EXITP

End Sub

Private Sub jurisdiction_Enter()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "jurisdiction_Enter"

12210   With Me
12220     If .jurisdiction.text = vbNullString Then
12230       ExpandCombo .jurisdiction  ' ** Module Procedure: modUtilities.
12240     End If
12250   End With

EXITP:
12260   Exit Sub

ERRH:
12270   THAT_PROC = THIS_PROC: That_Erl = Erl
12280   Select Case ERR.Number
        Case Else
12290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12300   End Select
12310   Resume EXITP

End Sub

Private Sub jurisdiction_NotInList(NewData As String, Response As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "jurisdiction_NotInList"

12410   With Me
12420     .jurisdiction = Null
12430     MsgBox "Must select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
12440     Response = acDataErrContinue
12450   End With

EXITP:
12460   Exit Sub

ERRH:
12470   THAT_PROC = THIS_PROC: That_Erl = Erl
12480   Select Case ERR.Number
        Case Else
12490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12500   End Select
12510   Resume EXITP

End Sub

Private Sub cmdStatementFreq_Click()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStatementFreq_Click"

        Dim strDocName As String

12610   With Me
12620     strThisAccountNo = .accountno
12630     cmdSave_Click  ' ** Procedure: Above.
12640     strDocName = "frmAccountProfile_StatementFreq"
12650     DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm & "~" & strThisAccountNo
12660   End With

EXITP:
12670   Exit Sub

ERRH:
12680   THAT_PROC = THIS_PROC: That_Erl = Erl
12690   Select Case ERR.Number
        Case Else
12700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12710   End Select
12720   Resume EXITP

End Sub

Private Sub opgFeeFreq_optMonth_AfterUpdate()
' ** Value: 1.

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFeeFreq_optMonth_AfterUpdate"

12810   With Me
12820     .opgFeeFreq_opt_box.Visible = False
12830     Select Case .opgFeeFreq_optMonth
          Case True
12840       .opgFeeFreq_optQuarter = False
12850       .opgFeeFreq_optSemi = False
12860       .opgFeeFreq_optAnnual = False
12870       .feeFrequency = 1
12880       .opgFeeFreq_opt_box.Left = .opgFeeFreq_optMonth_lbl.Left
12890       .opgFeeFreq_opt_box.Visible = True
12900     Case False
12910       .feeFrequency = 0
12920     End Select
12930   End With

EXITP:
12940   Exit Sub

ERRH:
12950   THAT_PROC = THIS_PROC: That_Erl = Erl
12960   Select Case ERR.Number
        Case Else
12970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12980   End Select
12990   Resume EXITP

End Sub

Private Sub opgFeeFreq_optQuarter_AfterUpdate()
' ** Value: 2.

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFeeFreq_optQuarter_AfterUpdate"

13010   With Me
13020     .opgFeeFreq_opt_box.Visible = False
13030     Select Case .opgFeeFreq_optQuarter
          Case True
13040       .opgFeeFreq_optMonth = False
13050       .opgFeeFreq_optSemi = False
13060       .opgFeeFreq_optAnnual = False
13070       .feeFrequency = 2
13080       .opgFeeFreq_opt_box.Left = .opgFeeFreq_optQuarter_lbl.Left
13090       .opgFeeFreq_opt_box.Visible = True
13100     Case False
13110       .feeFrequency = 0
13120     End Select
13130   End With

EXITP:
13140   Exit Sub

ERRH:
13150   THAT_PROC = THIS_PROC: That_Erl = Erl
13160   Select Case ERR.Number
        Case Else
13170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13180   End Select
13190   Resume EXITP

End Sub

Private Sub opgFeeFreq_optSemi_AfterUpdate()
' ** Value: 3.

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFeeFreq_optSemi_AfterUpdate"

13210   With Me
13220     .opgFeeFreq_opt_box.Visible = False
13230     Select Case .opgFeeFreq_optSemi
          Case True
13240       .opgFeeFreq_optMonth = False
13250       .opgFeeFreq_optQuarter = False
13260       .opgFeeFreq_optAnnual = False
13270       .feeFrequency = 3
13280       .opgFeeFreq_opt_box.Left = .opgFeeFreq_optSemi_lbl.Left
13290       .opgFeeFreq_opt_box.Visible = True
13300     Case False
13310       .feeFrequency = 0
13320     End Select
13330   End With

EXITP:
13340   Exit Sub

ERRH:
13350   THAT_PROC = THIS_PROC: That_Erl = Erl
13360   Select Case ERR.Number
        Case Else
13370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13380   End Select
13390   Resume EXITP

End Sub

Private Sub opgFeeFreq_optAnnual_AfterUpdate()
' ** Value: 4.

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFeeFreq_optAnnual_AfterUpdate"

13410   With Me
13420     .opgFeeFreq_opt_box.Visible = False
13430     Select Case .opgFeeFreq_optAnnual
          Case True
13440       .opgFeeFreq_optMonth = False
13450       .opgFeeFreq_optQuarter = False
13460       .opgFeeFreq_optSemi = False
13470       .feeFrequency = 4
13480       .opgFeeFreq_opt_box.Left = .opgFeeFreq_optAnnual_lbl.Left
13490       .opgFeeFreq_opt_box.Visible = True
13500     Case False
13510       .feeFrequency = 0
13520     End Select
13530   End With

EXITP:
13540   Exit Sub

ERRH:
13550   THAT_PROC = THIS_PROC: That_Erl = Erl
13560   Select Case ERR.Number
        Case Else
13570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13580   End Select
13590   Resume EXITP

End Sub

Private Sub cmdReviewFreq_Click()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReviewFreq_Click"

        Dim strDocName As String

13610   With Me
13620     strThisAccountNo = .accountno
13630     cmdSave_Click  ' ** Procedure: Above.
13640     strDocName = "frmAccountProfile_ReviewFreq"
13650     DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm & "~" & strThisAccountNo
13660   End With

EXITP:
13670   Exit Sub

ERRH:
13680   THAT_PROC = THIS_PROC: That_Erl = Erl
13690   Select Case ERR.Number
        Case Else
13700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13710   End Select
13720   Resume EXITP

End Sub

Private Sub Schedule_ID_Enter()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_ID_Enter"

13810   With Me
13820     If .Schedule_ID.text = vbNullString Then
13830       ExpandCombo .Schedule_ID  ' ** Module Procedure: modUtilities.
13840     End If
13850   End With

EXITP:
13860   Exit Sub

ERRH:
13870   THAT_PROC = THIS_PROC: That_Erl = Erl
13880   Select Case ERR.Number
        Case Else
13890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13900   End Select
13910   Resume EXITP

End Sub

Private Sub Schedule_ID_NotInList(NewData As String, Response As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_ID_NotInList"

14010   With Me
14020     .Schedule_ID = Null
14030   End With

EXITP:
14040   Exit Sub

ERRH:
14050   THAT_PROC = THIS_PROC: That_Erl = Erl
14060   Select Case ERR.Number
        Case Else
14070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14080   End Select
14090   Resume EXITP

End Sub

Private Sub discretion_Enter()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "discretion_Enter"

14110   With Me
14120     If .discretion.text = vbNullString Then
14130       ExpandCombo .discretion  ' ** Module Procedure: modUtilities.
14140     End If
14150   End With

EXITP:
14160   Exit Sub

ERRH:
14170   THAT_PROC = THIS_PROC: That_Erl = Erl
14180   Select Case ERR.Number
        Case Else
14190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14200   End Select
14210   Resume EXITP

End Sub

Private Sub discretion_NotInList(NewData As String, Response As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "discretion_NotInList"

14310   With Me
14320     .discretion = Null
14330   End With

EXITP:
14340   Exit Sub

ERRH:
14350   THAT_PROC = THIS_PROC: That_Erl = Erl
14360   Select Case ERR.Number
        Case Else
14370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14380   End Select
14390   Resume EXITP

End Sub

Private Sub adminno_Enter()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "adminno_Enter"

14410   With Me
14420     If .adminno.text = vbNullString Then
14430       ExpandCombo .adminno  ' ** Module Procedure: modUtilities.
14440     End If
14450   End With

EXITP:
14460   Exit Sub

ERRH:
14470   THAT_PROC = THIS_PROC: That_Erl = Erl
14480   Select Case ERR.Number
        Case Else
14490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14500   End Select
14510   Resume EXITP

End Sub

Private Sub adminno_NotInList(NewData As String, Response As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "adminno_NotInList"

14610   With Me
14620     .adminno = Null
14630     MsgBox "Please select an Administrator from the list.", vbInformation + vbOKOnly, "Invalid Entry"
14640     Response = acDataErrContinue
14650   End With

EXITP:
14660   Exit Sub

ERRH:
14670   THAT_PROC = THIS_PROC: That_Erl = Erl
14680   Select Case ERR.Number
        Case Else
14690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14700   End Select
14710   Resume EXITP

End Sub

Private Sub adminno_KeyDown(KeyCode As Integer, Shift As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "adminno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14820   intShiftDown = (Shift And acShiftMask) > 0
14830   intAltDown = (Shift And acAltMask) > 0
14840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14870       With Me
14880         intRetVal = 0
14890         If .dateclosed.Enabled = True Then
14900           .dateclosed.SetFocus
14910         ElseIf .Bank_Name.Visible = True Then
14920           .Bank_Name.SetFocus
14930         Else
14940           .FocusHolder.SetFocus
14950           .tabContacts = .tabContacts_pgContact1.PageIndex
14960           .Contact1_Name.SetFocus
14970           fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
14980         End If
14990       End With
15000     End Select
15010   End If

        ' ** Shift keys.
15020   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15030     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15040       With Me
15050         intRetVal = 0
15060         .discretion.SetFocus
15070       End With
15080     End Select
15090   End If

EXITP:
15100   KeyCode = intRetVal
15110   Exit Sub

ERRH:
15120   intRetVal = 0
15130   THAT_PROC = THIS_PROC: That_Erl = Erl
15140   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
15150   Case Else
15160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15170   End Select
15180   Resume EXITP

End Sub

Private Sub dateclosed_KeyDown(KeyCode As Integer, Shift As Integer)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "dateclosed_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15220   intShiftDown = (Shift And acShiftMask) > 0
15230   intAltDown = (Shift And acAltMask) > 0
15240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15270       With Me
15280         intRetVal = 0
15290         If .Bank_Name.Visible = True Then
15300           .Bank_Name.SetFocus
15310         Else
15320           .FocusHolder.SetFocus
15330           .tabContacts = .tabContacts_pgContact1.PageIndex
15340           .Contact1_Name.SetFocus
15350           fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
15360         End If
15370       End With
15380     End Select
15390   End If

        ' ** Shift keys.
15400   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15410     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15420       With Me
15430         intRetVal = 0
15440         .adminno.SetFocus
15450       End With
15460     End Select
15470   End If

EXITP:
15480   KeyCode = intRetVal
15490   Exit Sub

ERRH:
15500   intRetVal = 0
15510   THAT_PROC = THIS_PROC: That_Erl = Erl
15520   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
15530   Case Else
15540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15550   End Select
15560   Resume EXITP

End Sub

Private Sub Bank_Name_AfterUpdate()

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_Name_AfterUpdate"

        Dim strTmp01 As String

15610   With Me
15620     If IsNull(.Bank_Name) = False Then
15630       strTmp01 = FixQuotes(.Bank_Name)  ' ** Module Function: modStringFuncs.
15640       If Compare_StringA_StringB(.Bank_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
15650         .Bank_Name = strTmp01
15660         cmdSave_Click  ' ** Procedure: Above.
15670       End If
15680     End If
15690   End With

EXITP:
15700   Exit Sub

ERRH:
15710   THAT_PROC = THIS_PROC: That_Erl = Erl
15720   Select Case ERR.Number
        Case Else
15730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15740   End Select
15750   Resume EXITP

End Sub

Private Sub Bank_City_AfterUpdate()

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_City_AfterUpdate"

        Dim strTmp01 As String

15810   With Me
15820     If IsNull(.Bank_City) = False Then
15830       strTmp01 = FixQuotes(.Bank_City)  ' ** Module Function: modStringFuncs.
15840       If Compare_StringA_StringB(.Bank_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
15850         .Bank_City = strTmp01
15860         cmdSave_Click  ' ** Procedure: Above.
15870       End If
15880     End If
15890   End With

EXITP:
15900   Exit Sub

ERRH:
15910   THAT_PROC = THIS_PROC: That_Erl = Erl
15920   Select Case ERR.Number
        Case Else
15930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15940   End Select
15950   Resume EXITP

End Sub

Private Sub opgLoc_AfterUpdate()

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_AfterUpdate"

16010   With Me
16020     Select Case .opgLoc
          Case .opgLoc_optUSA.OptionValue
16030       .opgLoc_optUSA_lbl.FontBold = True
16040       .opgLoc_optOther_lbl.FontBold = False
16050       .Bank_State.Enabled = True
16060       .Bank_State.BorderColor = CLR_LTBLU2
16070       .Bank_State.BackStyle = acBackStyleNormal
16080       .Bank_State_lbl.BackStyle = acBackStyleNormal
16090       .Bank_State_lbl_box.Visible = False
16100       .Bank_Country.Enabled = False
16110       .Bank_Country.BorderColor = WIN_CLR_DISR
16120       .Bank_Country.BackStyle = acBackStyleTransparent
16130       .Bank_Country_lbl.BackStyle = acBackStyleTransparent
16140       .Bank_Country_lbl_box.Visible = True
16150       Select Case blnIsOpen
            Case True
16160         blnIsOpen = False
16170       Case False
16180         .Bank_State.SetFocus
16190       End Select
16200     Case .opgLoc_optOther.OptionValue
16210       .opgLoc_optUSA_lbl.FontBold = False
16220       .opgLoc_optOther_lbl.FontBold = True
16230       .Bank_State.Enabled = False
16240       .Bank_State.BorderColor = WIN_CLR_DISR
16250       .Bank_State.BackStyle = acBackStyleTransparent
16260       .Bank_State_lbl.BackStyle = acBackStyleTransparent
16270       .Bank_State_lbl_box.Visible = True
16280       .Bank_Country.Enabled = True
16290       .Bank_Country.BorderColor = CLR_LTBLU2
16300       .Bank_Country.BackStyle = acBackStyleNormal
16310       .Bank_Country_lbl.BackStyle = acBackStyleNormal
16320       .Bank_Country_lbl_box.Visible = False
16330       Select Case blnIsOpen
            Case True
16340         blnIsOpen = False
16350       Case False
16360         .Bank_Country.SetFocus
16370       End Select
16380     End Select
16390   End With

EXITP:
16400   Exit Sub

ERRH:
16410   Select Case ERR.Number
        Case Else
16420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
16430   End Select
16440   Resume EXITP

End Sub

Private Sub opgLoc_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16520   intShiftDown = (Shift And acShiftMask) > 0
16530   intAltDown = (Shift And acAltMask) > 0
16540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16570       With Me
16580         intRetVal = 0
16590         If .Bank_State.Enabled = True Then
16600           .Bank_State.SetFocus
16610         Else
16620           .Bank_Country.SetFocus
16630         End If
16640       End With
16650     End Select
16660   End If

        ' ** Shift keys.
16670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16690       With Me
16700         intRetVal = 0
16710         .Bank_City.SetFocus
16720       End With
16730     End Select
16740   End If

EXITP:
16750   KeyCode = intRetVal
16760   Exit Sub

ERRH:
16770   intRetVal = 0
16780   THAT_PROC = THIS_PROC
16790   That_Erl = Erl
16800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16810   Resume EXITP

End Sub

Private Sub opgLoc_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16920   intShiftDown = (Shift And acShiftMask) > 0
16930   intAltDown = (Shift And acAltMask) > 0
16940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16970       With Me
16980         intRetVal = 0
16990         If .Bank_State.Enabled = True Then
17000           .Bank_State.SetFocus
17010         Else
17020           .Bank_Country.SetFocus
17030         End If
17040       End With
17050     End Select
17060   End If

        ' ** Shift keys.
17070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17090       With Me
17100         intRetVal = 0
17110         .Bank_City.SetFocus
17120       End With
17130     End Select
17140   End If

EXITP:
17150   KeyCode = intRetVal
17160   Exit Sub

ERRH:
17170   intRetVal = 0
17180   THAT_PROC = THIS_PROC
17190   That_Erl = Erl
17200   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17210   Resume EXITP

End Sub

Private Sub Bank_State_Enter()

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_State_Enter"

17310   With Me
17320     If .Bank_State.text = vbNullString Then
17330       ExpandCombo .Bank_State  ' ** Module Procedure: modUtilities.
17340     End If
17350   End With

EXITP:
17360   Exit Sub

ERRH:
17370   THAT_PROC = THIS_PROC: That_Erl = Erl
17380   Select Case ERR.Number
        Case Else
17390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17400   End Select
17410   Resume EXITP

End Sub

Private Sub Bank_State_AfterUpdate()

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_State_AfterUpdate"

17510   With Me
17520     If IsNull(.Bank_State) = False Then
17530       If Trim(.Bank_State) <> vbNullString Then
17540         .Bank_State.Value = FormatUpperCase(Nz(.Bank_State, vbNullString))
17550         Select Case IsNull(.Acct_State_Pref)
              Case True
17560           .Acct_State_Pref = .Bank_State
17570         Case False
17580           If .Acct_State_Pref <> .Bank_State Then
17590             .Acct_State_Pref = .Bank_State
17600           End If
17610         End Select
17620         .Contact1_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
17630         .Contact2_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
17640       End If
17650     End If
17660   End With

EXITP:
17670   Exit Sub

ERRH:
17680   THAT_PROC = THIS_PROC: That_Erl = Erl
17690   Select Case ERR.Number
        Case Else
17700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17710   End Select
17720   Resume EXITP

End Sub

Private Sub Bank_State_NotInList(NewData As String, Response As Integer)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_State_NotInList"

17810   With Me
17820     .Bank_State = Null
17830     MsgBox "Please select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
17840     Response = acDataErrContinue
17850   End With

EXITP:
17860   Exit Sub

ERRH:
17870   THAT_PROC = THIS_PROC: That_Erl = Erl
17880   Select Case ERR.Number
        Case Else
17890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17900   End Select
17910   Resume EXITP

End Sub

Private Sub Bank_RoutingNumber_AfterUpdate()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_RoutingNumber_AfterUpdate"

18010   With Me
18020     If IsNull(.Bank_RoutingNumber) = False Then
18030       If Trim(.Bank_RoutingNumber) <> vbNullString Then
18040         If Len(Trim(.Bank_RoutingNumber)) > 9 Then
18050           .Bank_RoutingNumber = Left(Trim(.Bank_RoutingNumber), 9)
18060         End If
18070       End If
18080     End If
18090   End With

EXITP:
18100   Exit Sub

ERRH:
18110   THAT_PROC = THIS_PROC
18120   That_Erl = Erl
18130   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18140   Resume EXITP

End Sub

Private Sub PageBreak_1_to_2_lbl_Click()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "PageBreak_1_to_2_lbl_Click"

18210   DoCmd.GoToPage 2, 0, 0

EXITP:
18220   Exit Sub

ERRH:
18230   THAT_PROC = THIS_PROC: That_Erl = Erl
18240   Select Case ERR.Number
        Case Else
18250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18260   End Select
18270   Resume EXITP

End Sub

Private Sub PageBreak_1_to_2_lbl2_Click()

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "PageBreak_1_to_2_lbl2_Click"

18310   DoCmd.GoToPage 1, 0, 0

EXITP:
18320   Exit Sub

ERRH:
18330   THAT_PROC = THIS_PROC: That_Erl = Erl
18340   Select Case ERR.Number
        Case Else
18350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18360   End Select
18370   Resume EXITP

End Sub

Private Sub tabContacts_Change()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "tabContacts_Change"

18410   DoCmd.GoToPage 2, 0, 0

EXITP:
18420   Exit Sub

ERRH:
18430   THAT_PROC = THIS_PROC: That_Erl = Erl
18440   Select Case ERR.Number
        Case Else
18450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18460   End Select
18470   Resume EXITP

End Sub

Private Sub Contact1_Name_AfterUpdate()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Name_AfterUpdate"

        Dim strTmp01 As String

18510   With Me
18520     If IsNull(.Contact1_Name) = False Then
18530       strTmp01 = FixQuotes(.Contact1_Name)  ' ** Module Function: modStringFuncs.
18540       If Compare_StringA_StringB(.Contact1_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
18550         .Contact1_Name = strTmp01
18560         cmdSave_Click ' ** Procedure: Above.
18570       End If
18580     End If
18590   End With

EXITP:
18600   Exit Sub

ERRH:
18610   THAT_PROC = THIS_PROC: That_Erl = Erl
18620   Select Case ERR.Number
        Case Else
18630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18640   End Select
18650   Resume EXITP

End Sub

Private Sub Contact1_Name_KeyDown(KeyCode As Integer, Shift As Integer)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18720   intShiftDown = (Shift And acShiftMask) > 0
18730   intAltDown = (Shift And acAltMask) > 0
18740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18770       With Me
18780         intRetVal = 0
18790         DoCmd.GoToPage 2, 0, 0
18800         .FocusHolder.SetFocus
18810         .Contact1_Address1.SetFocus  ' ** This alone scrolls the page up a bit!
18820         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
18830       End With
18840     End Select
18850   End If

        ' ** Shift keys.
18860   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18870     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18880       With Me
18890         intRetVal = 0
18900         DoCmd.GoToPage 1, 0, 0
18910         If .FedIFNum2.Visible = True And .FedIFNum2.Enabled = True Then
18920           .FedIFNum2.SetFocus
18930         Else
18940           .dateclosed.SetFocus
18950         End If
18960         fSetScrollBarPosVT Me, 1&  ' ** Module Function: modScrollBarFuncs.
18970       End With
18980     End Select
18990   End If

EXITP:
19000   KeyCode = intRetVal
19010   Exit Sub

ERRH:
19020   intRetVal = 0
19030   THAT_PROC = THIS_PROC: That_Erl = Erl
19040   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
19050   Case Else
19060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19070   End Select
19080   Resume EXITP

End Sub

Private Sub Contact1_Address1_AfterUpdate()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Address1_AfterUpdate"

        Dim strTmp01 As String

19110   With Me
19120     If IsNull(.Contact1_Address1) = False Then
19130       strTmp01 = FixQuotes(.Contact1_Address1)  ' ** Module Function: modStringFuncs.
19140       If Compare_StringA_StringB(.Contact1_Address1, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
19150         .Contact1_Address1 = strTmp01
19160         cmdSave_Click  ' ** Procedure: Above.
19170       End If
19180     End If
19190   End With

EXITP:
19200   Exit Sub

ERRH:
19210   THAT_PROC = THIS_PROC: That_Erl = Erl
19220   Select Case ERR.Number
        Case Else
19230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19240   End Select
19250   Resume EXITP

End Sub

Private Sub Contact1_Address2_AfterUpdate()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Address2_AfterUpdate"

        Dim strTmp01 As String

19310   With Me
19320     If IsNull(.Contact1_Address2) = False Then
19330       strTmp01 = FixQuotes(.Contact1_Address2)  ' ** Module Function: modStringFuncs.
19340       If Compare_StringA_StringB(.Contact1_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
19350         .Contact1_Address2 = strTmp01
19360         cmdSave_Click  ' ** Procedure: Above.
19370       End If
19380     End If
19390   End With

EXITP:
19400   Exit Sub

ERRH:
19410   THAT_PROC = THIS_PROC: That_Erl = Erl
19420   Select Case ERR.Number
        Case Else
19430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19440   End Select
19450   Resume EXITP

End Sub

Private Sub Contact1_City_AfterUpdate()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_City_AfterUpdate"

        Dim strTmp01 As String

19510   With Me
19520     If IsNull(.Contact1_City) = False Then
19530       strTmp01 = FixQuotes(.Contact1_City)  ' ** Module Function: modStringFuncs.
19540       If Compare_StringA_StringB(.Contact1_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
19550         .Contact1_City = strTmp01
19560         cmdSave_Click  ' ** Procedure: Above.
19570       End If
19580     End If
19590   End With

EXITP:
19600   Exit Sub

ERRH:
19610   THAT_PROC = THIS_PROC: That_Erl = Erl
19620   Select Case ERR.Number
        Case Else
19630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19640   End Select
19650   Resume EXITP

End Sub

Public Sub opgLocContact1_AfterUpdate()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact1_AfterUpdate"

19710   With Me
19720     Select Case .opgLocContact1
          Case .opgLocContact1_optUSA.OptionValue
19730       .FocusHolder.SetFocus
19740       .Contact1_State.Enabled = True
19750       .Contact1_State.BackStyle = acBackStyleNormal
19760       .Contact1_State.BorderColor = CLR_LTBLU2
19770       .Contact1_State_lbl.BackStyle = acBackStyleNormal
19780       .Contact1_Zip.Enabled = True
19790       .Contact1_Zip.BackStyle = acBackStyleNormal
19800       .Contact1_Zip.BorderColor = CLR_LTBLU2
19810       .Contact1_Zip_lbl.BackStyle = acBackStyleNormal
19820       .Contact1_Country.Enabled = False
19830       .Contact1_Country.BackStyle = acBackStyleTransparent
19840       .Contact1_Country.BorderColor = WIN_CLR_DISR
19850       .Contact1_Country_lbl.BackStyle = acBackStyleTransparent
19860       .Contact1_PostalCode.Enabled = False
19870       .Contact1_PostalCode.BackStyle = acBackStyleTransparent
19880       .Contact1_PostalCode.BorderColor = WIN_CLR_DISR
19890       .Contact1_PostalCode_lbl.BackStyle = acBackStyleTransparent
19900       Select Case blnIsOpen
            Case True
19910         blnIsOpen = False
19920       Case False
19930         .Contact1_State.SetFocus
19940         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
19950       End Select
19960     Case .opgLocContact1_optOther.OptionValue
19970       .FocusHolder.SetFocus
19980       .Contact1_Country.Enabled = True
19990       .Contact1_Country.BackStyle = acBackStyleNormal
20000       .Contact1_Country.BorderColor = CLR_LTBLU2
20010       .Contact1_Country_lbl.BackStyle = acBackStyleNormal
20020       .Contact1_PostalCode.Enabled = True
20030       .Contact1_PostalCode.BackStyle = acBackStyleNormal
20040       .Contact1_PostalCode.BorderColor = CLR_LTBLU2
20050       .Contact1_PostalCode_lbl.BackStyle = acBackStyleNormal
20060       .Contact1_State.Enabled = False
20070       .Contact1_State.BackStyle = acBackStyleTransparent
20080       .Contact1_State.BorderColor = WIN_CLR_DISR
20090       .Contact1_State_lbl.BackStyle = acBackStyleTransparent
20100       .Contact1_Zip.Enabled = False
20110       .Contact1_Zip.BackStyle = acBackStyleTransparent
20120       .Contact1_Zip.BorderColor = WIN_CLR_DISR
20130       .Contact1_Zip_lbl.BackStyle = acBackStyleTransparent
20140       Select Case blnIsOpen
            Case True
20150         blnIsOpen = False
20160       Case False
20170         .Contact1_Country.SetFocus
20180         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
20190       End Select
20200     End Select
20210   End With

EXITP:
20220   Exit Sub

ERRH:
20230   THAT_PROC = THIS_PROC
20240   That_Erl = Erl
20250   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20260   Resume EXITP

End Sub

Private Sub opgLocContact1_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact1_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20320   intShiftDown = (Shift And acShiftMask) > 0
20330   intAltDown = (Shift And acAltMask) > 0
20340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20370       With Me
20380         intRetVal = 0
20390         If .Contact1_State.Enabled = True Then
20400           .Contact1_State.SetFocus
20410         Else
20420           .Contact1_Country.SetFocus
20430         End If
20440       End With
20450     End Select
20460   End If

        ' ** Shift keys.
20470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20490       With Me
20500         intRetVal = 0
20510         .Contact1_City.SetFocus
20520       End With
20530     End Select
20540   End If

EXITP:
20550   KeyCode = intRetVal
20560   Exit Sub

ERRH:
20570   intRetVal = 0
20580   THAT_PROC = THIS_PROC
20590   That_Erl = Erl
20600   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20610   Resume EXITP

End Sub

Private Sub opgLocContact1_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact1_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20720   intShiftDown = (Shift And acShiftMask) > 0
20730   intAltDown = (Shift And acAltMask) > 0
20740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20770       With Me
20780         intRetVal = 0
20790         If .Contact1_State.Enabled = True Then
20800           .Contact1_State.SetFocus
20810         Else
20820           .Contact1_Country.SetFocus
20830         End If
20840       End With
20850     End Select
20860   End If

        ' ** Shift keys.
20870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20880     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20890       With Me
20900         intRetVal = 0
20910         .Contact1_City.SetFocus
20920       End With
20930     End Select
20940   End If

EXITP:
20950   KeyCode = intRetVal
20960   Exit Sub

ERRH:
20970   intRetVal = 0
20980   THAT_PROC = THIS_PROC
20990   That_Erl = Erl
21000   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21010   Resume EXITP

End Sub

Private Sub Contact1_State_Enter()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_State_Enter"

21110   With Me
21120     If .Contact1_State.text = vbNullString Then
21130       ExpandCombo .Contact1_State  ' ** Module Procedure: modUtilities.
21140     End If
21150   End With

EXITP:
21160   Exit Sub

ERRH:
21170   THAT_PROC = THIS_PROC: That_Erl = Erl
21180   Select Case ERR.Number
        Case Else
21190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21200   End Select
21210   Resume EXITP

End Sub

Private Sub Contact1_State_AfterUpdate()

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_State_AfterUpdate"

21310   With Me
21320     If IsNull(.Contact1_State) = False Then
21330       If Trim(.Contact1_State) <> vbNullString Then
21340         .Contact1_State.Value = FormatUpperCase(Nz(.Contact1_State, vbNullString))
21350         Select Case IsNull(.Acct_State_Pref)
              Case True
21360           .Acct_State_Pref = .Contact1_State
21370         Case False
21380           If .Acct_State_Pref <> .Contact1_State Then
21390             .Acct_State_Pref = .Contact1_State
21400           End If
21410         End Select
21420         .Bank_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
21430         .Contact2_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
21440       End If
21450     End If
21460   End With

EXITP:
21470   Exit Sub

ERRH:
21480   THAT_PROC = THIS_PROC: That_Erl = Erl
21490   Select Case ERR.Number
        Case Else
21500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21510   End Select
21520   Resume EXITP

End Sub

Private Sub Contact1_State_NotInList(NewData As String, Response As Integer)

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_State_NotInList"

21610   With Me
21620     .Contact1_State = Null
21630     MsgBox "Please select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
21640     Response = acDataErrContinue
21650   End With

EXITP:
21660   Exit Sub

ERRH:
21670   THAT_PROC = THIS_PROC: That_Erl = Erl
21680   Select Case ERR.Number
        Case Else
21690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21700   End Select
21710   Resume EXITP

End Sub

Private Sub Contact1_Zip_Exit(Cancel As Integer)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Zip_Exit"

21810   With Me
21820     If .Contact1_Zip <> "" Then
21830       If Right(.Contact1_Zip.text, 4) <> "____" And Right(.Contact1_Zip.text, 4) <> "    " And _
                (Not IsNumeric(Mid(.Contact1_Zip.text, 7, 1)) Or Not IsNumeric(Mid(.Contact1_Zip.text, 8, 1)) Or _
                Not IsNumeric(Mid(.Contact1_Zip.text, 9, 1)) Or Not IsNumeric(Mid(.Contact1_Zip.text, 10, 1))) Then
21840         MsgBox "Please enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
21850         Cancel = -1
21860       End If
21870     End If
21880   End With

EXITP:
21890   Exit Sub

ERRH:
21900   THAT_PROC = THIS_PROC: That_Erl = Erl
21910   Select Case ERR.Number
        Case Else
21920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21930   End Select
21940   Resume EXITP

End Sub

Private Sub Contact1_Email_KeyDown(KeyCode As Integer, Shift As Integer)

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Email_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22020   intShiftDown = (Shift And acShiftMask) > 0
22030   intAltDown = (Shift And acAltMask) > 0
22040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22070       With Me
22080         intRetVal = 0
22090         .GoToPage 1, 0, 0
22100         .accountno.SetFocus  ' ** To make sure the Detail gets scrolled up.
22110         fSetScrollBarPosVT Me, 1&  ' ** Module Function: modScrollBarFuncs.
22120       End With
22130     End Select
22140   End If

        ' ** Shift keys.
22150   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22170       With Me
22180         DoCmd.GoToPage 2, 0, 0
22190         .FocusHolder.SetFocus
22200         .Contact1_Fax.SetFocus  ' ** This alone scrolls the page up a bit!
22210         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
22220       End With
22230     End Select
22240   End If

EXITP:
22250   KeyCode = intRetVal
22260   Exit Sub

ERRH:
22270   intRetVal = 0
22280   THAT_PROC = THIS_PROC: That_Erl = Erl
22290   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
22300   Case Else
22310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22320   End Select
22330   Resume EXITP

End Sub

Private Sub Contact2_Name_AfterUpdate()

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Name_AfterUpdate"

        Dim strTmp01 As String

22410   With Me
22420     If IsNull(.Contact2_Name) = False Then
22430       strTmp01 = FixQuotes(.Contact2_Name)  ' ** Module Function: modStringFuncs.
22440       If Compare_StringA_StringB(.Contact2_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
22450         .Contact2_Name = strTmp01
22460         cmdSave_Click  ' ** Procedure: Above.
22470       End If
22480     End If
22490   End With

EXITP:
22500   Exit Sub

ERRH:
22510   THAT_PROC = THIS_PROC: That_Erl = Erl
22520   Select Case ERR.Number
        Case Else
22530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22540   End Select
22550   Resume EXITP

End Sub

Private Sub Contact2_Name_KeyDown(KeyCode As Integer, Shift As Integer)

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22620   intShiftDown = (Shift And acShiftMask) > 0
22630   intAltDown = (Shift And acAltMask) > 0
22640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22670       With Me
22680         intRetVal = 0
22690         .Contact2_Address1.SetFocus
22700       End With
22710     End Select
22720   End If

        ' ** Shift keys.
22730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22750       With Me
22760         intRetVal = 0
22770         .Contact1_Email.SetFocus
22780       End With
22790     End Select
22800   End If

EXITP:
22810   KeyCode = intRetVal
22820   Exit Sub

ERRH:
22830   intRetVal = 0
22840   THAT_PROC = THIS_PROC: That_Erl = Erl
22850   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
22860   Case Else
22870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22880   End Select
22890   Resume EXITP

End Sub

Private Sub Contact2_Address1_AfterUpdate()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Address1_AfterUpdate"

        Dim strTmp01 As String

22910   With Me
22920     If IsNull(.Contact2_Address1) = False Then
22930       strTmp01 = FixQuotes(.Contact2_Address1)  ' ** Module Function: modStringFuncs.
22940       If Compare_StringA_StringB(.Contact2_Address1, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
22950         .Contact2_Address1 = strTmp01
22960         cmdSave_Click  ' ** Procedure: Above.
22970       End If
22980     End If
22990   End With

EXITP:
23000   Exit Sub

ERRH:
23010   THAT_PROC = THIS_PROC: That_Erl = Erl
23020   Select Case ERR.Number
        Case Else
23030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23040   End Select
23050   Resume EXITP

End Sub

Private Sub Contact2_Address2_AfterUpdate()

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Address2_AfterUpdate"

        Dim strTmp01 As String

23110   With Me
23120     If IsNull(.Contact2_Address2) = False Then
23130       strTmp01 = FixQuotes(.Contact2_Address2)  ' ** Module Function: modStringFuncs.
23140       If Compare_StringA_StringB(.Contact2_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
23150         .Contact2_Address2 = strTmp01
23160         cmdSave_Click  ' ** Procedure: Above.
23170       End If
23180     End If
23190   End With

EXITP:
23200   Exit Sub

ERRH:
23210   THAT_PROC = THIS_PROC: That_Erl = Erl
23220   Select Case ERR.Number
        Case Else
23230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23240   End Select
23250   Resume EXITP

End Sub

Private Sub Contact2_City_AfterUpdate()

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_City_AfterUpdate"

        Dim strTmp01 As String

23310   With Me
23320     If IsNull(.Contact2_City) = False Then
23330       strTmp01 = FixQuotes(.Contact2_City)  ' ** Module Function: modStringFuncs.
23340       If Compare_StringA_StringB(.Contact2_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
23350         .Contact2_City = strTmp01
23360         cmdSave_Click  ' ** Procedure: Above.
23370       End If
23380     End If
23390   End With

EXITP:
23400   Exit Sub

ERRH:
23410   THAT_PROC = THIS_PROC: That_Erl = Erl
23420   Select Case ERR.Number
        Case Else
23430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23440   End Select
23450   Resume EXITP

End Sub

Public Sub opgLocContact2_AfterUpdate()

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact2_AfterUpdate"

23510   With Me
23520     Select Case .opgLocContact2
          Case .opgLocContact2_optUSA.OptionValue
23530       .FocusHolder.SetFocus
23540       .Contact2_State.Enabled = True
23550       .Contact2_State.BackStyle = acBackStyleNormal
23560       .Contact2_State.BorderColor = CLR_LTBLU2
23570       .Contact2_State_lbl.BackStyle = acBackStyleNormal
23580       .Contact2_Zip.Enabled = True
23590       .Contact2_Zip.BackStyle = acBackStyleNormal
23600       .Contact2_Zip.BorderColor = CLR_LTBLU2
23610       .Contact2_Zip_lbl.BackStyle = acBackStyleNormal
23620       .Contact2_Country.Enabled = False
23630       .Contact2_Country.BackStyle = acBackStyleTransparent
23640       .Contact2_Country.BorderColor = WIN_CLR_DISR
23650       .Contact2_Country_lbl.BackStyle = acBackStyleTransparent
23660       .Contact2_PostalCode.Enabled = False
23670       .Contact2_PostalCode.BackStyle = acBackStyleTransparent
23680       .Contact2_PostalCode.BorderColor = WIN_CLR_DISR
23690       .Contact2_PostalCode_lbl.BackStyle = acBackStyleTransparent
23700       Select Case blnIsOpen
            Case True
23710         blnIsOpen = False
23720       Case False
23730         .Contact2_State.SetFocus
23740         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
23750       End Select
23760     Case .opgLocContact2_optOther.OptionValue
23770       .FocusHolder.SetFocus
23780       .Contact2_Country.Enabled = True
23790       .Contact2_Country.BackStyle = acBackStyleNormal
23800       .Contact2_Country.BorderColor = CLR_LTBLU2
23810       .Contact2_Country_lbl.BackStyle = acBackStyleNormal
23820       .Contact2_PostalCode.Enabled = True
23830       .Contact2_PostalCode.BackStyle = acBackStyleNormal
23840       .Contact2_PostalCode.BorderColor = CLR_LTBLU2
23850       .Contact2_PostalCode_lbl.BackStyle = acBackStyleNormal
23860       .Contact2_State.Enabled = False
23870       .Contact2_State.BackStyle = acBackStyleTransparent
23880       .Contact2_State.BorderColor = WIN_CLR_DISR
23890       .Contact2_State_lbl.BackStyle = acBackStyleTransparent
23900       .Contact2_Zip.Enabled = False
23910       .Contact2_Zip.BackStyle = acBackStyleTransparent
23920       .Contact2_Zip.BorderColor = WIN_CLR_DISR
23930       .Contact2_Zip_lbl.BackStyle = acBackStyleTransparent
23940       Select Case blnIsOpen
            Case True
23950         blnIsOpen = False
23960       Case False
23970         .Contact2_Country.SetFocus
23980         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
23990       End Select
24000     End Select
24010   End With

EXITP:
24020   Exit Sub

ERRH:
24030   THAT_PROC = THIS_PROC
24040   That_Erl = Erl
24050   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24060   Resume EXITP

End Sub

Private Sub opgLocContact2_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact2_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24120   intShiftDown = (Shift And acShiftMask) > 0
24130   intAltDown = (Shift And acAltMask) > 0
24140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24170       With Me
24180         intRetVal = 0
24190         If .Contact2_State.Enabled = True Then
24200           .Contact2_State.SetFocus
24210         Else
24220           .Contact2_Country.SetFocus
24230         End If
24240       End With
24250     End Select
24260   End If

        ' ** Shift keys.
24270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24290       With Me
24300         intRetVal = 0
24310         .Contact2_City.SetFocus
24320       End With
24330     End Select
24340   End If

EXITP:
24350   KeyCode = intRetVal
24360   Exit Sub

ERRH:
24370   intRetVal = 0
24380   THAT_PROC = THIS_PROC
24390   That_Erl = Erl
24400   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24410   Resume EXITP

End Sub

Private Sub opgLocContact2_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact2_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24520   intShiftDown = (Shift And acShiftMask) > 0
24530   intAltDown = (Shift And acAltMask) > 0
24540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24570       With Me
24580         intRetVal = 0
24590         If .Contact2_State.Enabled = True Then
24600           .Contact2_State.SetFocus
24610         Else
24620           .Contact2_Country.SetFocus
24630         End If
24640       End With
24650     End Select
24660   End If

        ' ** Shift keys.
24670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24690       With Me
24700         intRetVal = 0
24710         .Contact2_City.SetFocus
24720       End With
24730     End Select
24740   End If

EXITP:
24750   KeyCode = intRetVal
24760   Exit Sub

ERRH:
24770   intRetVal = 0
24780   THAT_PROC = THIS_PROC
24790   That_Erl = Erl
24800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24810   Resume EXITP

End Sub

Private Sub Contact2_State_Enter()

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_State_Enter"

24910   With Me
24920     If .Contact2_State.text = vbNullString Then
24930       ExpandCombo .Contact2_State  ' ** Module Procedure: modUtilities.
24940     End If
24950   End With

EXITP:
24960   Exit Sub

ERRH:
24970   THAT_PROC = THIS_PROC: That_Erl = Erl
24980   Select Case ERR.Number
        Case Else
24990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25000   End Select
25010   Resume EXITP

End Sub

Private Sub Contact2_State_AfterUpdate()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_State_AfterUpdate"

25110   With Me
25120     If IsNull(.Contact2_State) = False Then
25130       If Trim(.Contact2_State) <> vbNullString Then
25140         .Contact2_State.Value = FormatUpperCase(Nz(.Contact2_State, vbNullString))
25150         Select Case IsNull(.Acct_State_Pref)
              Case True
25160           .Acct_State_Pref = .Contact2_State
25170         Case False
25180           If .Acct_State_Pref <> .Contact2_State Then
25190             .Acct_State_Pref = .Contact2_State
25200           End If
25210         End Select
25220         .Bank_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
25230         .Contact1_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
25240       End If
25250     End If
25260   End With

EXITP:
25270   Exit Sub

ERRH:
25280   THAT_PROC = THIS_PROC: That_Erl = Erl
25290   Select Case ERR.Number
        Case Else
25300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25310   End Select
25320   Resume EXITP

End Sub

Private Sub Contact2_State_NotInList(NewData As String, Response As Integer)

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_State_NotInList"

25410   With Me
25420     .Contact2_State = Null
25430     MsgBox "Please select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
25440     Response = acDataErrContinue
25450   End With

EXITP:
25460   Exit Sub

ERRH:
25470   THAT_PROC = THIS_PROC: That_Erl = Erl
25480   Select Case ERR.Number
        Case Else
25490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25500   End Select
25510   Resume EXITP

End Sub

Private Sub Contact2_Zip_Exit(Cancel As Integer)

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Zip_Exit"

25610   With Me
25620     If .Contact2_Zip <> "" Then
25630       If Right(.Contact2_Zip.text, 4) <> "____" And Right(.Contact2_Zip.text, 4) <> "    " And _
                (Not IsNumeric(Mid(.Contact2_Zip.text, 7, 1)) Or Not IsNumeric(Mid(.Contact2_Zip.text, 8, 1)) Or _
                Not IsNumeric(Mid(.Contact2_Zip.text, 9, 1)) Or Not IsNumeric(Mid(.Contact2_Zip.text, 10, 1))) Then
25640         MsgBox "Please enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
25650         Cancel = -1
25660       End If
25670     End If
25680   End With

EXITP:
25690   Exit Sub

ERRH:
25700   THAT_PROC = THIS_PROC: That_Erl = Erl
25710   Select Case ERR.Number
        Case Else
25720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25730   End Select
25740   Resume EXITP

End Sub

Private Sub Contact2_Email_KeyDown(KeyCode As Integer, Shift As Integer)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Email_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25820   intShiftDown = (Shift And acShiftMask) > 0
25830   intAltDown = (Shift And acAltMask) > 0
25840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25870       With Me
25880         intRetVal = 0
25890         .Notes.SetFocus
25900       End With
25910     End Select
25920   End If

        ' ** Shift keys.
25930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25950       With Me
25960         intRetVal = 0
25970         .Contact2_Fax.SetFocus
25980       End With
25990     End Select
26000   End If

EXITP:
26010   KeyCode = intRetVal
26020   Exit Sub

ERRH:
26030   intRetVal = 0
26040   THAT_PROC = THIS_PROC: That_Erl = Erl
26050   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
26060   Case Else
26070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26080   End Select
26090   Resume EXITP

End Sub

Private Sub Notes_GotFocus()

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "Notes_GotFocus"

26110   With Me
26120     .Notes.SelLength = 0
26130     .Notes.SelStart = 999
26140   End With

EXITP:
26150   Exit Sub

ERRH:
26160   THAT_PROC = THIS_PROC: That_Erl = Erl
26170   Select Case ERR.Number
        Case Else
26180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26190   End Select
26200   Resume EXITP

End Sub

Private Sub Notes_KeyDown(KeyCode As Integer, Shift As Integer)

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "Notes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26320   intShiftDown = (Shift And acShiftMask) > 0
26330   intAltDown = (Shift And acAltMask) > 0
26340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26360     Select Case intRetVal
          Case vbKeyTab
26370       With Me
26380         intRetVal = 0
26390         DoCmd.SelectObject acForm, .Parent.Name, False
26400         If .Parent.cmdPrintReport.Enabled = True Then
26410           .Parent.cmdPrintReport.SetFocus
26420         ElseIf .Parent.cmdOK.Enabled = True Then
26430           .Parent.cmdOK.SetFocus
26440         Else
26450           .Parent.cmdCancel.SetFocus
26460         End If
26470       End With
26480     End Select
26490   End If

        ' ** Shift keys.
26500   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26510     Select Case intRetVal
          Case vbKeyTab
26520       With Me
26530         intRetVal = 0
26540         .Contact2_Email.SetFocus
26550       End With
26560     End Select
26570   End If

EXITP:
26580   KeyCode = intRetVal
26590   Exit Sub

ERRH:
26600   intRetVal = 0
26610   THAT_PROC = THIS_PROC: That_Erl = Erl
26620   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
26630   Case Else
26640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26650   End Select
26660   Resume EXITP

End Sub
