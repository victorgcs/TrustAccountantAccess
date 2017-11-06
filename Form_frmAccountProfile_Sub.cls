VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountProfile_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountProfile_Sub"

'VGC 10/26/2017: CHANGES!

' #########################
' ## Use VBA_RenumErrh().  36840
' #########################

' ** Shortcut Alt keys responsive from this form:
' **   Fee Annually:     A {opgFeeFreq_optAnnual}
' **   AccountNo:        B {accountno}
' **   Close:            C {cmdClose on frmAccountProfile}
' **   Notes:            E {tabContacts_pgNotes}
' **   Bank Other:       H {opgLoc_optOther}
' **   Statement Freq:   K {cmdStatementFreq}
' **   Review Freq:      L {cmdReviewFreq}
' **   Fee Monthly:      M {opgFeeFreq_optMonth}
' **   Short Name:       N {shortname}
' **   Short Name PQ:    N {shortname_pq} (Printer's Quotes)
' **   Choose RA:        O {cmdChoose}
' **   Print Report:     P {cmdPrintReport on frmAccountProfile}
' **   Fee Quarterly:    Q {opgFeeFreq_optQuarter}
' **   Related Accounts: R {related_accountno}
' **   Fee Semiannually: S {opgFeeFreq_optSemi}
' **   Account Type:     T {accounttype}
' **   Bank USA:         U {opgLoc_optUSA}
' **   Exit:             X {cmdClose on frmAccountProfile}
' **   Contact Name 1:   1 {tabContacts_pgContact1}
' **   Contact Name 2:   2 {tabContacts_pgContact2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Contact1 Other:   H {opgLocContact1_optOther}
' **   Next Record:      N {MoveRec}
' **   Previous Record:  P {MoveRec}
' **   Save:             S {cmdSave}
' **   Contact1 USA:     U {opgLocContact1_optUSA}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Contact2 Other:   H {opgLocContact2_optOther}
' **   Contact2 USA:     U {opgLocContact2_optUSA}

'WHAT ON EARTH IS 'STATE TYPE'? (statetype)

' ** Demo Zips:
' ** Zip1: 55114-6538
' ** Zip2: 55103

' ** Queries with DateClosed:
'QRY: 'qryAccount_Contacts_18a' dateclosed  'EXCLUDED!
'QRY: 'qryAccount_Contacts_18b' dateclosed  'EXCLUDED!
'QRY: 'qryAccountBalance_01a' dateclosed)   'EXCLUDED!
'QRY: 'qryAccountBalance_02a' dateclosed)   'EXCLUDED!
'QRY: 'qryAccountExport_02_01' dateclosed      'INCLUDED!
'QRY: 'qryAccountNoDropDown_01' dateclosed  'EXCLUDED!
'QRY: 'qryAccountNoDropDown_02' dateclosed  'EXCLUDED!
'QRY: 'qryAccountNoDropDown_03' dateclosed  'EXCLUDED!
'QRY: 'qryAccountNoDropDown_04' dateclosed  'EXCLUDED!
'QRY: 'qryAccountNoDropDown_05' dateclosed  'EXCLUDED!
'QRY: 'qryAssetList_01_02' dateclosed)      'EXCLUDED!
'QRY: 'qryCashControl_01' dateclosed)       'EXCLUDED!
'QRY: 'qryCashControl_02c' dateclosed)      'EXCLUDED!
'QRY: 'qryCheckReconcile_03a' dateclosed    'EXCLUDED!
'QRY: 'qryCheckReconcile_03b' dateclosed    'EXCLUDED!
'QRY: 'qryCheckReconcile_03c' dateclosed    'EXCLUDED!
'QRY: 'qryCheckReconcile_03d' dateclosed    'EXCLUDED!
'QRY: 'qryMap_06' dateclosed                'INCLUDED!
'QRY: 'qryMap_Misc_LTCL_03' dateclosed      'INCLUDED!
'QRY: 'qryMap_Misc_STCGL_03' dateclosed     'INCLUDED!
'QRY: 'qryQualifyingAccountsForStatement' dateclosed)  'EXCLUDED!
'QRY: 'qryStatementBalance_11' dateclosed   'EXCLUDED!
'QRY: 'qryStatementBalance_12' dateclosed   'EXCLUDED!
'QRY: 'qryStatementParameters_20' dateclosed)  'EXCLUDED!

' In ('qryAccount_Contacts_18a','qryAccount_Contacts_18b','qryAccountBalance_01a','qryAccountBalance_02a','qryAccountExport_02_01','qryAccountNoDropDown_01','qryAccountNoDropDown_02','qryAccountNoDropDown_03','qryAccountNoDropDown_04','qryAccountNoDropDown_05','qryAssetList_01_02','qryCashControl_01','qryCashControl_02c','qryCheckReconcile_03a','qryCheckReconcile_03b','qryCheckReconcile_03c','qryCheckReconcile_03d','qryMap_06','qryMap_Misc_LTCL_03','qryMap_Misc_STCGL_03','qryQualifyingAccountsForStatement','qryStatementBalance_11','qryStatementBalance_12','qryStatementParameters_20')

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnSysAcct As Boolean, blnActNoSel As Boolean, blnViewOnly As Boolean
Private blnExit As Boolean, blnEscPressed As Boolean, blnIsOpen As Boolean
Private blnAccountHasChanged As Boolean, blnNewHasSaved As Boolean
Private strThisAccountNo As String
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim lngAssetsLbl_Offset As Long

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountProfile" Then

            ' ** Check for Related Accounts.
160         RelAcct_Chk  ' ** Function: below.

170         blnViewOnly = False

            ' ** Initialize - Later it will change to True when a
            ' ** user changes the status in any one of the data forms.
180         blnAccountHasChanged = False
190         .Parent.AcctChanged_Set blnAccountHasChanged  ' ** Form Procedure: frmAccountProfile.

200         If .ActiveAssets > 0 Then
210           .dateclosed.Enabled = False
220           .dateclosed.BorderColor = WIN_CLR_DISR
230           .dateclosed_lbl.BackStyle = acBackStyleTransparent
240         Else
250           .dateclosed.Enabled = True
260           .dateclosed.BorderColor = CLR_LTBLU2
270           .dateclosed_lbl.BackStyle = acBackStyleNormal
280         End If
290         If IsNull(.dateclosed) = False Then
300           .dateclosed.Enabled = True
310           .dateclosed.BorderColor = CLR_LTBLU2
320           .dateclosed_lbl.BackStyle = acBackStyleNormal
330         End If

340         lngAssetsLbl_Offset = .cmbAssets.Left - .cmbAssets_lbl.Left

350         Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

360         If IsNull(.Acct_State_Pref) = False Then
370           .Bank_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
380           .Contact1_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
390           .Contact2_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
400         End If

            ' ** This will get superceded by SysAcct_Set(), below!
410         blnIsOpen = True
420         opgLoc_AfterUpdate  ' ** Procedure: Below.

430         blnIsOpen = True
440         opgLocContact1_AfterUpdate  ' ** Procedure: Below.

450         blnIsOpen = True
460         opgLocContact2_AfterUpdate  ' ** Procedure: Below.

470         If gblnSeparateCheckingAccounts = True Then
480           .Bank_Name.Visible = True
490           .Bank_City.Visible = True
500           .opgLoc.Visible = True
510           .Bank_State.Visible = True
520           .Bank_Country.Visible = True
530           .Bank_RoutingNumber.Visible = True
540           .Bank_AccountNumber.Visible = True
550           .FedIFNum1_lbl.Visible = True
560           .FedIFNum1.Visible = True
570           .FedIFNum2.Visible = True
580         Else
590           .Bank_Name.Visible = False
600           .Bank_City.Visible = False
610           .opgLoc.Enabled = False
620           .opgLoc.Visible = False
630           .opgLoc_box.Visible = False
640           .Bank_State.Visible = False
650           .Bank_Country.Visible = False
660           .Bank_RoutingNumber.Visible = False
670           .Bank_AccountNumber.Visible = False
680           .FedIFNum1_lbl.Visible = False
690           .FedIFNum1.Visible = False
700           .FedIFNum2.Visible = False
710           .cmbAssets.Top = .Bank_Name.Top
720           .cmbAssets.Left = .Bank_Name.Left
730           .cmbAssets_lbl.Top = .Bank_Name_lbl.Top
740           .cmbAssets_lbl.Left = (.cmbAssets.Left - lngAssetsLbl_Offset)
750           .LastCheckNum.Top = .cmbAssets.Top
760           .LastCheckNum_lbl.Top = .cmbAssets_lbl.Top
770         End If

780       Else
790         Cancel = -1
800       End If
810     End With

820     If Cancel = -1 Then
830       Beep
840       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
850     End If

EXITP:
860     Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
900     Resume EXITP

End Sub

Private Sub Form_Load()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

910     With Me

920       CLR_DISABLED_FG = CLR_DKGRY
930       CLR_DISABLED_BG = CLR_LTTEAL

940       .description.ForeColor = CLR_DISABLED_FG
950       .description.BackColor = CLR_DISABLED_BG
960       .accountno2.ForeColor = CLR_DISABLED_FG
970       .accountno2.BackColor = CLR_DISABLED_BG
980       .shortname2.ForeColor = CLR_DISABLED_FG
990       .shortname2.BackColor = CLR_DISABLED_BG

1000      .opgFeeFreq_opt_box.Visible = False
1010      If IsNull(.feeFrequency) = False Then
1020        If .feeFrequency > 0 Then
1030          Select Case .feeFrequency
              Case 1
1040            .opgFeeFreq_optMonth = True
1050            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optMonth_lbl.Left
1060            .opgFeeFreq_opt_box.Visible = True
1070          Case 2
1080            .opgFeeFreq_optQuarter = True
1090            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optQuarter_lbl.Left
1100            .opgFeeFreq_opt_box.Visible = True
1110          Case 3
1120            .opgFeeFreq_optSemi = True
1130            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optSemi_lbl.Left
1140            .opgFeeFreq_opt_box.Visible = True
1150          Case 4
1160            .opgFeeFreq_optAnnual = True
1170            .opgFeeFreq_opt_box.Left = .opgFeeFreq_optAnnual_lbl.Left
1180            .opgFeeFreq_opt_box.Visible = True
1190          End Select
1200        End If
1210      End If

1220      .accountno.SetFocus

1230    End With

EXITP:
1240    Exit Sub

ERRH:
860     THAT_PROC = THIS_PROC
870     That_Erl = Erl
880     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
890     Resume EXITP

End Sub

Private Sub Form_Current()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim blnTmp01 As Boolean

1310    blnExit = False: blnEscPressed = False
1320    blnAccountHasChanged = False: blnNewHasSaved = False

1330    With Me

1340      blnTmp01 = blnSysAcct
1350      If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Or _
              .accountno = "SUSPENSE" Or .accountno = "99-SUSPENSE" Then
1360        blnSysAcct = True
1370      Else
1380        blnSysAcct = False
1390      End If
1400      .Parent.SysAcct_Set blnSysAcct  ' ** Form Procedure: frmAccountProfile.

1410      If IsNull(.accountno) = False Then  ' ** Should no longer ever be empty.
1420        If .accountno <> vbNullString Then
1430          strThisAccountNo = .accountno
1440        Else
1450          strThisAccountNo = vbNullString
1460          blnNewHasSaved = False
1470        End If
1480      Else
1490        strThisAccountNo = vbNullString
1500        blnNewHasSaved = False
1510      End If
1520      .Parent.HasSaved_Set strThisAccountNo, blnNewHasSaved  ' ** Form Procedure: frmAccountProfile.

1530      If IsNull(.related_accountno) = True And blnSysAcct = False Then
1540        .related_accountno_lbl2.Visible = True
1550        .related_accountno_lbl3.Visible = False  ' ** Doesn't become visible until they're in the field.
1560        .related_accountno_lbl4.Visible = True   ' ** Small focus-holding field, but not for entry.
1570        .related_accountno.Visible = False
1580      Else
1590        .related_accountno_lbl2.Visible = False
1600        .related_accountno_lbl3.Visible = False
1610        .related_accountno_lbl4.Visible = False
1620        .related_accountno.Visible = True
1630      End If

1640      Select Case .S_PQuotes
          Case True
1650        .shortname_pq.Visible = True
1660        .shortname_pq.Enabled = True
1670        .shortname.Visible = False
1680        .shortname.Enabled = False
1690      Case False
1700        If .shortname_pq.Visible = True Then
1710          .shortname_pq.Visible = False
1720          .shortname_pq.Enabled = False
1730          .shortname.Visible = True
1740          .shortname.Enabled = True
1750        End If
1760      End Select

1770      Select Case .L_PQuotes
          Case True
1780        .legalname_pq.Visible = True
1790        .legalname_pq.Enabled = True
1800        .legalname.Visible = False
1810        .legalname.Enabled = False
1820      Case False
1830        If .legalname_pq.Visible = True Then
1840          .legalname_pq.Visible = False
1850          .legalname_pq.Enabled = False
1860          .legalname.Visible = True
1870          .legalname.Enabled = True
1880        End If
1890      End Select

1900      SysAcct_Set_AP blnSysAcct, blnViewOnly, CLR_DISABLED_FG, CLR_DISABLED_BG, Me  ' ** Module Procedure: modAcctProfileFuncs.

1910      If gblnSeparateCheckingAccounts = True Then
1920        blnIsOpen = True
1930        opgLoc_AfterUpdate  ' ** Procedure: Below.
1940      End If

1950      .cmbAssets.Requery

          ' ** VGC 11/01/2010: For now, I'm using an unused field for the default asset.
1960      If IsNull(.taxlot) = False Then
1970        If IsNumeric(.taxlot) = True Then
1980          If Val(.taxlot) > 0 Then
1990            .cmbAssets = CLng(Val(.taxlot))
2000          Else
2010            .cmbAssets = Null  ' ** If it's Zero, it drops assets!?
2020          End If
2030        Else
2040          .cmbAssets = Null
2050        End If
2060      Else
2070        .cmbAssets = Null
2080      End If

2090      If blnTmp01 = False Then
            ' ** If it just came from a non-system account, there's
            ' ** always a residual dark label border still showing.
2100        fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
2110        DoEvents
2120        fSetScrollBarPosVT Me, 1&  ' ** Module Function: modScrollBarFuncs.
2130        .Repaint
2140      End If

2150      .accountno.SetFocus
2160      blnActNoSel = True
2170      .TimerInterval = 250&

2180    End With

EXITP:
2190    Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 3022    ' ** The changes you requested to the table were not successful because they
880       ERR.Clear  ' ** would create duplicate values in the index, primary key, or relationship.
890       ERR.Raise vbObjectError + 27  ' **  vbObjectError = -2147221504  WHAT DOES THIS LINE MEAN? 'Automation Error'
900     Case 2455  ' ** You entered an expressions that has an invalid reference to the property |.
          ' ** Do nothing.
910     Case Else
920       THAT_PROC = THIS_PROC
930       That_Erl = Erl
940       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
950     End Select
960     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strControl As String
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality.
        ' **   Notes:            E {tabContacts_pgNotes}
        ' **   Short Name:       N {shortname}
        ' **   Short Name PQ:    N {shortname_pq} (Printer's Quotes)
        ' **   Contact Name 1:   1 {tabContacts_pgContact1}
        ' **   Contact Name 2:   2 {tabContacts_pgContact2}

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   Contact1 Other:   H {opgLocContact1_optOther}
        ' **   Next Record:      N {Form_KeyDown}
        ' **   Previous Record:  P {Form_KeyDown}
        ' **   Save:             S {cmdSave}
        ' **   Contact1 USA:     U {opgLocContact1_optUSA}

        ' ** Shortcut Ctrl-Shift to other forms and functionality.
        ' **   Contact2 Other:   H {opgLocContact2_optOther}
        ' **   Contact2 USA:     U {opgLocContact2_optUSA}

2250  On Error Resume Next
2260    strControl = Screen.ActiveControl.Name
2270  On Error GoTo ERRH

        ' ** Plain keys.
2280    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2290      Select Case intRetVal
          Case vbKeyEscape
2300        With Me
2310          intRetVal = 0
2320          blnEscPressed = True
2330          .Parent.EscPressed_Set blnEscPressed  ' ** Form Procedure: frmAccountProfile.
2340          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountProfile.
2350        End With
2360      Case vbKeyPageDown
2370        With Me
2380          intRetVal = 0
2390          DoCmd.GoToPage 2, 0, 0
2400          .Contact1_Name.SetFocus
2410          fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
2420        End With
2430      Case vbKeyPageUp
2440        With Me
2450          If strControl = "accountno" Then
2460            intRetVal = 0
2470            .GoToPage 2, 0, 0
2480            If strControl <> "accountno" Then
2490              .accountno.SetFocus
2500            End If
2510          Else
2520            intRetVal = 0
2530            .GoToPage 1, 0, 0
2540          End If
2550        End With
2560      End Select
2570    End If

        ' ** Alt keys.
2580    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2590      Select Case intRetVal
          Case vbKeyC
2600        With Me
2610          DoCmd.SelectObject acForm, .Parent.Name, False
2620          .Parent.cmdClose.SetFocus
2630          .Parent.cmdClose_Click
2640        End With
2650      Case vbKeyE
2660        With Me
2670          intRetVal = 0
2680          .FocusHolder.SetFocus
2690          .tabContacts = .tabContacts_pgNotes.PageIndex
2700          .Notes.SetFocus
2710          fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
2720        End With
2730      Case vbKeyN
2740        With Me
2750          intRetVal = 0
2760          If .shortname.Visible = True And .shortname.Enabled = True Then
2770            .shortname.SetFocus
2780          ElseIf .shortname_pq.Visible = True And .shortname_pq.Enabled = True Then
2790            .shortname_pq.SetFocus
2800          End If
2810        End With
2820      Case vbKeyX
2830        With Me
2840          intRetVal = 0
2850          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountProfile.
2860        End With
2870      Case vbKey1
2880        With Me
2890          intRetVal = 0
2900          .FocusHolder.SetFocus
2910          .tabContacts = .tabContacts_pgContact1.PageIndex
2920          .Contact1_Name.SetFocus
2930          fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
2940        End With
2950      Case vbKey2
2960        With Me
2970          intRetVal = 0
2980          .FocusHolder.SetFocus
2990          .tabContacts = .tabContacts_pgContact2.PageIndex
3000          .Contact2_Name.SetFocus
3010          fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
3020        End With
3030      End Select
3040    End If

        ' ** Ctrl keys.
3050    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3060      Select Case intRetVal
          Case vbKeyH
3070        With Me
3080          If .opgLocContact1 <> .opgLocContact1_optOther.OptionValue Then
3090            .opgLocContact1 = .opgLocContact1_optOther.OptionValue
3100            opgLocContact1_AfterUpdate  ' ** Procedure: Below.
3110          End If
3120        End With
3130      Case vbKeyN
3140        intRetVal = 0
3150        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
3160      Case vbKeyP
3170        intRetVal = 0
3180        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
3190      Case vbKeyS
3200        intRetVal = 0
3210        cmdSave_Click THIS_PROC ' ** Procedure: Below.
3220      Case vbKeyU
3230        With Me
3240          If .opgLocContact1 <> .opgLocContact1_optUSA.OptionValue Then
3250            .opgLocContact1 = .opgLocContact1_optUSA.OptionValue
3260            opgLocContact1_AfterUpdate  ' ** Procedure: Below.
3270          End If
3280        End With
3290      Case vbKeyTab, vbKeyReturn
3300        With Me
3310          intRetVal = 0
3320          DoCmd.SelectObject acForm, .Parent.Name, False
3330          .Parent.cmdClose.SetFocus
3340        End With
3350      End Select
3360    End If

        ' ** Ctrl-Shift keys.
3370    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3380      Select Case intRetVal
          Case vbKeyH
3390        With Me
3400          If .opgLocContact2 <> .opgLocContact2_optOther.OptionValue Then
3410            .opgLocContact2 = .opgLocContact2_optOther.OptionValue
3420            opgLocContact2_AfterUpdate  ' ** Procedure: Below.
3430          End If
3440        End With
3450      Case vbKeyU
3460        With Me
3470          If .opgLocContact2 <> .opgLocContact2_optUSA.OptionValue Then
3480            .opgLocContact2 = .opgLocContact2_optUSA.OptionValue
3490            opgLocContact2_AfterUpdate  ' ** Procedure: Below.
3500          End If
3510        End With
3520      Case vbKeyTab, vbKeyReturn
3530        With Me
3540          intRetVal = 0
3550          DoCmd.SelectObject acForm, .Parent.Name, False
3560          .Parent.cmdClose.SetFocus
3570        End With
3580      End Select
3590    End If

EXITP:
3600    KeyCode = intRetVal
3610    Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
910     Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim varTmp00 As Variant, strTmp01 As String

3710    With Me
3720      If IsNull(.accounttype) = True And IsNull(.accountno) = False Then
3730        MsgBox "Please enter an account type to continue", vbInformation + vbOKOnly, "Entry Required"
3740        .accounttype.SetFocus
3750        Cancel = -1
3760      Else
3770        If IsNull(.shortname) = True And IsNull(.accountno) = False Then
3780          MsgBox "Please enter a Short Name to continue", vbInformation + vbOKOnly, "Entry Required"
3790          .shortname.SetFocus
3800          Cancel = -1
3810        Else
3820          If gblnAccountNoWithType = True Then
                ' ** Verify the user entered a valid account type prefix to the account number.
3830            If InStr(.accountno, "-") > 0 Then
3840              Set dbs = CurrentDb
3850              Set rst = dbs.OpenRecordset("SELECT accounttype.* FROM accounttype WHERE accounttype.accounttype = '" & _
                    Left(.accountno, 2) & "'")
3860              If rst.EOF Then
3870                rst.Close
3880                dbs.Close
3890                MsgBox "The Account Type you entered as an Account Number prefix does not exist." & vbCrLf & vbCrLf & _
                      "Please enter a valid Account Type to continue.", vbInformation + vbOKOnly, "Invalid Entry"
3900                .accountno.Undo
3910                Cancel = -1
3920              Else
3930                rst.Close
3940                dbs.Close
3950                .accounttype = Left(.accountno, 2)
3960                .accounttype.Requery
3970              End If
3980            Else
                  ' ** Add the chosen Account Type to accountno.
3990              strTmp01 = .accounttype & "-" & Trim(.accountno)
4000              varTmp00 = DLookup("[shortname]", "account", "[accountno] = '" & strTmp01 & "'")
4010              If IsNull(varTmp00) = True Then
4020                .accountno = strTmp01
4030                DoEvents
4040              Else
4050                Beep
4060                MsgBox "An account with this number already exists." & vbCrLf & _
                      "   " & strTmp01 & "   " & CStr(varTmp00), vbInformation + vbOKOnly, "Account Exists"
4070                .accountno.Undo
4080                Cancel = -1
4090              End If
4100            End If
4110          End If
              ' ** In case they flipped around, make sure only one remains.
4120          If gblnSeparateCheckingAccounts = True Then
4130            Select Case .opgLoc
                Case .opgLoc_optUSA.OptionValue
4140              If IsNull(.Bank_Country) = False Then
4150                .Bank_Country = Null
4160              End If
4170            Case .opgLoc_optOther.OptionValue
4180              If IsNull(.Bank_State) = False Then
4190                .Bank_State = Null
4200              End If
4210            End Select
4220          End If
4230          Select Case .opgLocContact1
              Case .opgLocContact1_optUSA.OptionValue
4240            If IsNull(.Contact1_Country) = False Then
4250              .Contact1_Country = Null
4260            End If
4270            If IsNull(.Contact1_PostalCode) = False Then
4280              .Contact1_PostalCode = Null
4290            End If
4300          Case .opgLocContact1_optOther.OptionValue
4310            If IsNull(.Contact1_State) = False Then
4320              .Contact1_State = Null
4330            End If
4340            If IsNull(.Contact1_Zip) = False Then
4350              .Contact1_Zip = Null
4360            End If
4370          End Select
4380          Select Case .opgLocContact2
              Case .opgLocContact2_optUSA.OptionValue
4390            If IsNull(.Contact2_Country) = False Then
4400              .Contact2_Country = Null
4410            End If
4420            If IsNull(.Contact2_PostalCode) = False Then
4430              .Contact2_PostalCode = Null
4440            End If
4450          Case .opgLocContact2_optOther.OptionValue
4460            If IsNull(.Contact2_State) = False Then
4470              .Contact2_State = Null
4480            End If
4490            If IsNull(.Contact2_Zip) = False Then
4500              .Contact2_Zip = Null
4510            End If
4520          End Select
4530        End If
4540      End If
4550    End With

EXITP:
4560    Set rst = Nothing
4570    Set dbs = Nothing
4580    Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
880     Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
890     Case Else
900       THAT_PROC = THIS_PROC
910       That_Erl = Erl
920       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
930     End Select
940     Resume EXITP

End Sub

Private Sub Form_Timer()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

4610    With Me
4620      .TimerInterval = 0
4630      If blnActNoSel = True Then
4640        .accountno.SetFocus
4650        .accountno.SelLength = 0
4660        .accountno.SelStart = 19
4670      End If
4680    End With

EXITP:
4690    Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

4710    If DataErr <> 0 Then
4720      Select Case DataErr
          Case 13, 91, 94, 2001, 2105, 2110, 2116, 2135, 2169, 2474, 2501, 3020, 3314, 7753, 8519, 8530
            ' ** 2001: You Canceled the previous operation.
            ' ** 2501: The '|' action was Canceled.
            ' ** 3314: The field '|' can't contain a Null value because the Required property for this field is set to True.
4730        Response = acDataErrContinue
4740      Case 8050  ' ** Do you want to save changes to the design of '|'?
            ' ** Do nothing.
4750      Case 3021  ' ** No current record.
4760        If Me.CurrentRecord > 1 Then
4770          DoCmd.GoToRecord , , acPrevious
4780        End If
4790      Case 3314  ' ** The field '|' can't contain a Null value because the Required property for this field is set to True.
4800        DoCmd.RunCommand acCmdUndo
4810        If Screen.ActiveControl.Name <> "accountno" Then
4820          Me.accountno.SetFocus
4830        End If
4840        Response = acDataErrContinue
4850      Case 2046  ' ** The command or action isn't available now (first or last record).
4860        DoCmd.CancelEvent
4870      Case 2113, 2279
            ' ** 2113: The value you entered isn't valid for this field.
            ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
4880        Select Case Screen.ActiveControl.Name
            Case "Contact1_Phone1", "Contact1_Phone2", "Contact2_Phone1", "Contact2_Phone2", "Contact1_Fax", "Contact2_Fax"
4890          MsgBox "Please enter valid phone number.", vbInformation + vbOKOnly, "Invalid Entry"
4900        Case "Contact1_Zip", "Contact2_Zip"
4910          MsgBox "Please enter valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
4920        Case "Contact1_State", "Contact2_State"
4930          MsgBox "Please enter valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
4940        Case "documentdate", "appointmentdate"
4950          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Entry"
4960        End Select
4970        Response = acDataErrContinue
4980      Case 2108  ' ** You must save the field before you execute the GoToControl action, the GoToControl method, or the SetFocus method.
4990        Me.accountno.text = ""
5000        DoCmd.RunCommand acCmdUndo
5010      Case 3101  ' ** The Microsoft Jet database engine cannot find a record in the table '|' with key matching field(s) '|'.
5020        DoCmd.CancelEvent
5030      Case 2237  ' ** The text you entered isn't an item in the list.
5040        DoCmd.CancelEvent
5050      Case 3162  ' ** You tried to assign the Null value to a variable that isn't a Variant data type.
5060        MsgBox "The " & Screen.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
            ' ** Do nothing.
5070      Case 3022
            ' ** The changes you requested to the table were not successful because they
            ' ** would create duplicate values in the index, primary key, or relationship.
5080        MsgBox "The Account Number you entered already exists." & vbCrLf & vbCrLf & _
              "Please enter a unique Account Number to continue.", vbInformation + vbOKOnly, "Duplicate Account Number"
5090        Me.accountno.SetFocus
5100        Me.accountno.text = ""
5110      Case 7787  ' ** The record your are modifying has been changed by another user.
            ' ** Ignore.
5120        Response = acDataErrContinue
5130      Case Else
5140        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
5150        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
5160      End Select
5170    End If

        ' ** AcDataError enumeration:
        ' **   0  acDataErrContinue  Ignore the error and continue without displaying the default Microsoft Access
        ' **                         error message. A custom error message may be displayed in place of the default
        ' **                         error message.
        ' **   1  acDataErrDisplay   Display the default Microsoft Access error message. (Default)
        ' **   2  acDataErrAdded     Don't display the default Microsoft Access error message. The entry may be
        ' **                         added to the combo box list in the NotInList event procedure. After the entry
        ' **                         is added, Microsoft Access updates the list by requerying the combo box.
        ' **                         Microsoft Access then rechecks the string against the combo box list, and saves
        ' **                         the value in the NewData argument in the field the combo box is bound to. If
        ' **                         the string is not in the list, then Microsoft Access displays an error message.

EXITP:
5180    Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** Doesn't work from here, so moved to parent.
        ' ** Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

5210    If strThisAccountNo <> vbNullString Then
5220      gstrActNo = strThisAccountNo  ' ** Only used for hidden transactions and account profile.
5230    End If

EXITP:
5240    Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Public Sub cmdSave_Click(strProc As String)
' ** This is a fake click event.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

        Dim intPos01 As Integer, intLen As Integer
        Dim intX As Integer

5310    With Me

5320      intLen = Len(strProc)
5330      intPos01 = 0
5340      For intX = intLen To 1 Step -1
5350        If Mid(strProc, intX, 1) = "_" Then
5360          intPos01 = intX
5370          Exit For
5380        End If
5390      Next
5400      strProc = Left(strProc, (intPos01 - 1))

5410      DoCmd.RunCommand acCmdSaveRecord
5420      DoEvents

5430      If strThisAccountNo = vbNullString Then
            ' ** Leave strThisAccountNo empty, so it's known as new throughout this module.
5440        blnNewHasSaved = True
5450        .Parent.HasSaved_Set strThisAccountNo, blnNewHasSaved  ' ** Form Procedure: frmAccountProfile.
5460      End If

          ' ** When hitting Enter a couple of times after accounttype on a new account,
          ' ** it took me to a new record, and no way to get back to where I was to finish it!
5470      Select Case strProc
          Case "Form"  ' ** Form_KeyDown().
5480        strProc = vbNullString
5490  On Error Resume Next
5500        strProc = Screen.ActiveControl.Name
5510  On Error GoTo ERRH
5520        If strProc <> vbNullString Then
5530          .Controls(strProc).SetFocus
5540        End If
5550      Case Else
5560        Select Case strProc
            Case "cmdClose"
5570          DoCmd.SelectObject acForm, .Parent.Name, False
5580          .Parent.cmdClose.SetFocus
5590        Case Else
5600          .Controls(strProc).SetFocus
              ' ** cmdChoose_Click()
              ' ** accounttype_AfterUpdate().
              ' ** county_Enter().
              ' ** cmdStatementFreq_Click().
              ' ** cmdReviewFreq_Click().
              ' ** dateclosed_AfterUpdate().
5610        End Select
5620      End Select

5630    End With  ' ** Me.

EXITP:
5640    Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
880     Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
890     Case 3164  ' ** Field cannot be updated.
          ' ** A required field value is still missing.
900     Case 3314  ' ** The field '|' can't contain a Null value because the Required property for this field is set to True.
          ' ** Ignore for now.
910     Case 3709  ' ** The search key was not found in any record.
          ' ** I have no idea what this means!
920     Case Else
930       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
940     End Select
950     Resume EXITP

End Sub

Private Sub accountno_GotFocus()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_GotFocus"

5710    With Me
5720      .accountno.SelLength = 0
5730      .accountno.SelStart = 19
5740    End With

EXITP:
5750    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub accountno_Change()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_Change"

        Dim blnContinue As Boolean

5810    With Me

5820      blnContinue = True

5830      If gblnAccountNoWithType = True Then
5840        If Len(.accountno.text) > 15 Then
5850          blnContinue = False
5860          MsgBox "The Account Number is limited to 15 characters.", vbInformation + vbOKOnly, "Invalid Entry"
5870          .accountno.Undo
5880        End If
5890      Else
5900        If Len(.accountno.text) > 12 Then
5910          blnContinue = False
5920          MsgBox "AccountNo is limited to 12 characters.", vbInformation + vbOKOnly, "Invalid Entry"
5930          .accountno.Undo
5940        End If
5950      End If

5960      If blnContinue = True Then
5970        If .accountno.OldValue = "INCOME O/U" Or .accountno.OldValue = "99-INCOME O/U" Then
5980          blnContinue = False
5990          MsgBox "You cannot change this account.", vbInformation + vbOKOnly, "Invalid Action"
6000          .accountno.Undo
6010        ElseIf .accountno.OldValue = "SUSPENSE" Or .accountno.OldValue = "99-SUSPENSE" Then
6020          blnContinue = False
6030          MsgBox "You cannot change this account.", vbInformation + vbOKOnly, "Invalid Action"
6040          .accountno.Undo
6050        Else
6060          AccountChange  ' ** Procedure: Below.
6070        End If
6080      End If

6090    End With

EXITP:
6100    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub accountno_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6220    intShiftDown = (Shift And acShiftMask) > 0
6230    intAltDown = (Shift And acAltMask) > 0
6240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6270        With Me
6280          intRetVal = 0
6290          Select Case blnSysAcct
              Case True
6300            .shortname.SetFocus
6310          Case False
6320            If IsNull(.related_accountno) = True Then
6330              If .related_accountno_lbl4.Enabled = False Then .related_accountno_lbl4.Enabled = True
6340              If .related_accountno_lbl4.Visible = False Then .related_accountno_lbl4.Visible = True
6350              .related_accountno_lbl4.SetFocus
6360            Else
6370              .related_accountno.SetFocus
6380            End If
6390          End Select
6400        End With
6410      End Select
6420    End If

        ' ** Shift keys.
6430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6450        With Me
6460          intRetVal = 0
6470          DoCmd.SelectObject acForm, .Parent.Name, False
6480          .Parent.cmdClose.SetFocus
6490        End With
6500      End Select
6510    End If

EXITP:
6520    KeyCode = intRetVal
6530    Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub accountno_AfterUpdate()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountCompare As String
        Dim strAcctName As String, strAcctDate As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

6610    blnContinue = True

6620    With Me

          ' ** Remove spaces in accountno.
6630      strAccountCompare = StringReplace(.accountno, " ", "")  ' ** Module Function: modUtilities.
6640      If (strAccountCompare <> .accountno) And (.accountno <> "INCOME O/U") And (.accountno <> "99-INCOME O/U") Then
6650        .accountno.text = strAccountCompare
6660        msgResponse = MsgBox("Account Number cannot have spaces!" & vbCrLf & "They have been removed.", vbExclamation + vbOKOnly, "Error")
6670      End If

6680      If (IsNull(.accountno) = True Or Trim(.accountno.text) = vbNullString) And blnAccountHasChanged = True Then
6690        If blnEscPressed = False And blnExit = False Then
6700          blnContinue = False
6710          .accountno.text = vbNullString
6720          MsgBox "Please enter a valid account number.", vbInformation + vbOKOnly, "Entry Required"
6730          DoCmd.CancelEvent
6740          .accountno.SetFocus
6750        End If
6760      End If

6770      If blnContinue = True Then
6780        Set dbs = CurrentDb
6790        Set qdf = dbs.QueryDefs("qryAccountProfile_05")
6800        With qdf.Parameters
6810          ![actno] = Me.accountno
6820        End With
6830        Set rst = qdf.OpenRecordset
6840        If rst.BOF = True And rst.EOF = True Then
              ' ** No hit, it's new.
6850        Else
6860          If .NewRecord = True Or .accountno.OldValue <> .accountno Then
6870            blnContinue = False
6880            strAcctName = vbNullString: strAcctDate = vbNullString
6890            If IsNull(rst![shortname]) = True Then strAcctName = "{unknown}" Else strAcctName = rst![shortname]
6900            If IsNull(rst![predate]) = False Then strAcctDate = Format(rst![predate], "mm/dd/yyyy")
6910            MsgBox "The Account Number you entered already exists:" & vbCrLf & _
                  "  Name: " & strAcctName & vbCrLf & _
                  "  Dated: " & strAcctDate & vbCrLf & vbCrLf & _
                  "Please enter a unique Account Number to continue.", vbInformation + vbOKOnly, "Duplicate Account Number"
6920            DoCmd.CancelEvent
6930            .accountno.SetFocus
6940  On Error Resume Next
6950            .accountno.text = ""
6960  On Error GoTo ERRH
6970          End If
6980        End If
6990        rst.Close
7000        dbs.Close

7010      End If

7020    End With

EXITP:
7030    Set rst = Nothing
7040    Set dbs = Nothing
7050    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub related_accountno_lbl4_GotFocus()
' ** Text box.
' ** This is the small one, just holding focus.
' ** related_accountno_lbl2 is just for the white background and border.

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "related_accountno_lbl4_GotFocus"

7110    With Me
7120      .related_accountno_lbl3.Visible = True  ' ** This is the one with 'Choose -->'.
7130    End With

EXITP:
7140    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub related_accountno_lbl4_KeyDown(KeyCode As Integer, Shift As Integer)
' ** Text box.

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "related_accountno_lbl4_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7220    intShiftDown = (Shift And acShiftMask) > 0
7230    intAltDown = (Shift And acAltMask) > 0
7240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7270        With Me
7280          intRetVal = 0
7290          If IsNull(.accounttype) = True Or IsNull(.shortname) = True Then
7300            .shortname.SetFocus
7310          Else
7320            .cmdChoose.SetFocus
7330          End If
7340        End With
7350      End Select
7360    End If

        ' ** Shift keys.
7370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7380      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7390        With Me
7400          intRetVal = 0
7410          .accountno.SetFocus
7420        End With
7430      End Select
7440    End If

EXITP:
7450    KeyCode = intRetVal
7460    Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub related_accountno_lbl4_LostFocus()
' ** Text box.

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "related_accountno_lbl4_LostFocus"

7510    With Me
7520      .related_accountno_lbl3.Visible = False  ' ** This is the one with 'Choose -->'.
7530    End With

EXITP:
7540    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub related_accountno_Change()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "related_accountno_Change"

7610    AccountChange  ' ** Procedure: Below.

EXITP:
7620    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Public Sub cmdChoose_Click()
' ** Public for GoToReports.

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChoose_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strRelAccts As String, strQry As String
        Dim intOrderBy As Integer
        Dim varTmp00 As Variant

        Const opgOrderBy_optNumber    As Integer = 1
        Const opgOrderBy_optName      As Integer = 2
        Const opgOrderBy_optSelected  As Integer = 3

7710    With Me

7720      If IsNull(.accounttype) = True Then
7730        MsgBox "Please choose an Account Type before selecting Related Accounts.", vbInformation + vbOKOnly, "Entry Required"
7740      ElseIf IsNull(.shortname) = True Then
7750        MsgBox "Please enter the account's Short Name before selecting Related Accounts.", vbInformation + vbOKOnly, "Entry Required"
7760      Else

7770        DoCmd.Hourglass True
7780        DoEvents

7790        If IsNull(.related_accountno) = False Then
7800          If Trim(.related_accountno) <> vbNullString Then
7810            strRelAccts = Trim(.related_accountno)
7820          Else
7830            strRelAccts = vbNullString
7840          End If
7850        Else
7860          strRelAccts = vbNullString
7870        End If

7880        cmdSave_Click THIS_PROC  ' ** Procedure: Above.

7890        Set dbs = CurrentDb
7900        With dbs
              ' ** Empty tmpAccount_RelAccts.
7910          Set qdf = .QueryDefs("qryAccountProfile_RelAccts_02")
7920          qdf.Execute
7930          Set qdf = Nothing
              ' ** tblPreference_User, linked to tblPreference_Control, just opgOrderBy on
              ' ** frmAccountProfile_RelAccts, by specified CurrentAppName(), CurrentUser().
7940          varTmp00 = DLookup("[prefuser_integer]", "qryAccountProfile_RelAccts_07")  '##dbs_id
7950          Select Case IsNull(varTmp00)
              Case True
7960            intOrderBy = opgOrderBy_optNumber
7970          Case False
7980            intOrderBy = varTmp00
7990          End Select
8000          Select Case intOrderBy
              Case opgOrderBy_optNumber
                ' ** Append qryAccountProfile_RelAccts_03 (Account, linked to
                ' ** qryAccountMenu_01_10 (xx), qryAccountNoDropDown_06 (xx), accountno,
                ' ** Desc, with cash criteria; sort: accountno) to tmpAccount_RelAccts.
8010            strQry = "qryAccountProfile_RelAccts_03a"
8020          Case opgOrderBy_optName
                ' ** Append qryAccountProfile_RelAccts_04 (Account, linked to
                ' ** qryAccountMenu_01_10 (xx), qryAccountNoDropDown_06 (xx), accountno,
                ' ** shortname, with cash criteria; sort: shortname) to tmpAccount_RelAccts.
8030            strQry = "qryAccountProfile_RelAccts_04a"
8040          Case opgOrderBy_optSelected
                ' ** Append qryAccountProfile_RelAccts_05 (Account, linked to
                ' ** qryAccountMenu_01_10 (xx), qryAccountNoDropDown_06 (xx), accountno,
                ' ** shortname, with cash criteria; sort: IsRelated) to tmpAccount_RelAccts.
8050            strQry = "qryAccountProfile_RelAccts_05a"
8060          End Select
8070          Set qdf = .QueryDefs(strQry)
8080          qdf.Execute
8090          Set qdf = Nothing
8100          .Close
8110        End With  ' ** dbs.
8120        Set qdf = Nothing
8130        Set dbs = Nothing

8140        .Parent.RelAccts_Set True  ' ** Form Procedure: frmAccountProfile.

8150        DoCmd.OpenForm "frmAccountProfile_RelAccts", acNormal, , , , acDialog, strCallingForm & "~" & strRelAccts

8160        DoEvents
8170        If .related_accountno <> strRelAccts Then
8180          AccountChange  ' ** Procedure: Below.
8190        End If

            ' ** Make sure this form is on top.
8200        DoCmd.SelectObject acForm, strCallingForm, False

8210      End If

8220    End With  ' ** Me.

EXITP:
8230    Set qdf = Nothing
8240    Set dbs = Nothing
8250    Exit Sub

ERRH:
870     DoCmd.Hourglass False
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub cmdChoose_KeyDown(KeyCode As Integer, Shift As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChoose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8320    intShiftDown = (Shift And acShiftMask) > 0
8330    intAltDown = (Shift And acAltMask) > 0
8340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8360      Select Case intRetVal
          Case vbKeyTab
8370        With Me
8380          intRetVal = 0
8390          .shortname.SetFocus
8400        End With
8410      Case vbKeyReturn
8420        With Me
8430          If IsNull(.accounttype) = True Or IsNull(.shortname) = True Then
8440            intRetVal = 0
8450            .shortname.SetFocus
8460          Else
                ' ** Let it go through to invoke cmdChoose_Click.
8470          End If
8480        End With
8490      End Select
8500    End If

        ' ** Shift keys.
8510    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8520      Select Case intRetVal
          Case vbKeyTab
8530        With Me
8540          intRetVal = 0
8550          If IsNull(.related_accountno) = True And .related_accountno_lbl4.Visible = True Then
8560            .related_accountno_lbl4.SetFocus
8570          ElseIf .related_accountno_lbl4.Visible = False Then
8580            .related_accountno.SetFocus
8590          End If
8600        End With
8610      End Select
8620    End If

EXITP:
8630    KeyCode = intRetVal
8640    Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub shortname_GotFocus()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_GotFocus"

8710    With Me
8720      .shortname.SelLength = 0
8730      .shortname.SelStart = 99
8740    End With

EXITP:
8750    Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub shortname_AfterUpdate()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_AfterUpdate"

        Dim strTmp01 As String, strTmp02 As String

8810    With Me
8820      If IsNull(.shortname) = False Then
8830        If Trim(.shortname) <> vbNullString Then
8840          strTmp01 = Trim(.shortname)
8850          strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
8860          If strTmp02 <> strTmp01 Then
8870            .shortname = strTmp02
8880            .shortname_pq.Visible = True
8890            .shortname_pq.Enabled = True
8900            .CaseNum.SetFocus
8910            .shortname.Visible = False
8920            .shortname.Enabled = False
8930          End If
8940        End If
8950      End If
8960    End With

EXITP:
8970    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub shortname_Change()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_Change"

9010    With Me
9020      If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
9030        MsgBox "You cannot change this acount.", vbInformation + vbOKOnly, "Invalid Action"
9040        .shortname.Undo
9050      ElseIf .accountno = "SUSPENSE" Or .accountno = "99-SUSPENSE" Then
9060        MsgBox "You cannot change this acount.", vbInformation + vbOKOnly, "Invalid Action"
9070        .shortname.Undo
9080      Else
9090        AccountChange  ' ** Procedure: Below.
9100      End If
9110    End With

EXITP:
9120    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub shortname_pq_GotFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_pq_GotFocus"

9210    With Me
9220      .shortname_pq.SelLength = 0
9230      .shortname_pq.SelStart = 99
9240    End With

EXITP:
9250    Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub shortname_pq_AfterUpdate()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_pq_AfterUpdate"

        Dim strTmp01 As String, strTmp02 As String

        ' ** Single-quote, double-quote, printer's quote, printers quote, printers' quotes.
        ' **  '  Chr(39)
        ' **  "  Chr(34)
        ' **  ‘  Chr(145)
        ' **  ’  Chr(146)
        ' **  “  Chr(147)
        ' **  ”  Chr(148)

9310    With Me
9320      If IsNull(.shortname_pq) = False Then
9330        If Trim(.shortname_pq) <> vbNullString Then
9340          strTmp01 = Trim(.shortname_pq)
9350          If InStr(strTmp01, Chr(145)) = 0 And InStr(strTmp01, Chr(146)) = 0 And _
                  InStr(strTmp01, Chr(147)) = 0 And InStr(strTmp01, Chr(148)) = 0 Then
                ' ** They've removed the printers quotes entirely.
9360            strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
9370            If strTmp02 <> strTmp01 Then
                  ' ** And replaced them with regular quotes! Put them back.
9380              .shortname_pq = strTmp02
9390              .CaseNum.SetFocus
9400            Else
9410              .shortname.Visible = True
9420              .shortname.Enabled = True
9430              .CaseNum.SetFocus
9440              .shortname_pq.Visible = False
9450              .shortname_pq.Enabled = False
9460            End If
9470          Else
9480            strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
9490            If strTmp02 <> strTmp01 Then
9500              .shortname_pq = strTmp02
9510              .CaseNum.SetFocus
9520            End If
9530          End If
9540        End If
9550      End If
9560    End With

EXITP:
9570    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub shortname_pq_Change()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_pq_Change"

9610    With Me
9620      If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
9630        MsgBox "You cannot change this acount.", vbInformation + vbOKOnly, "Invalid Action"
9640        .shortname_pq.Undo
9650      ElseIf .accountno = "SUSPENSE" Or .accountno = "99-SUSPENSE" Then
9660        MsgBox "You cannot change this acount.", vbInformation + vbOKOnly, "Invalid Action"
9670        .shortname_pq.Undo
9680      Else
9690        AccountChange  ' ** Procedure: Below.
9700      End If
9710    End With

EXITP:
9720    Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub legalname_GotFocus()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_GotFocus"

9810    With Me
9820      .legalname.SelLength = 0
9830      .legalname.SelStart = 99
9840    End With

EXITP:
9850    Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub legalname_AfterUpdate()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_AfterUpdate"

        Dim strTmp01 As String, strTmp02 As String

9910    With Me
9920      If IsNull(.legalname) = False Then
9930        If Trim(.legalname) <> vbNullString Then
9940          strTmp01 = Trim(.legalname)
9950          strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
9960          If strTmp02 <> strTmp01 Then
9970            .legalname = strTmp02
9980            .legalname_pq.Visible = True
9990            .legalname_pq.Enabled = True
10000           .cotrustee.SetFocus
10010           .legalname.Visible = False
10020           .legalname.Enabled = False
10030         End If
10040       End If
10050     End If
10060   End With

EXITP:
10070   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub legalname_Change()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_Change"

10110   AccountChange  ' ** Procedure: Below.

EXITP:
10120   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub legalname_pq_GotFocus()

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_pq_GotFocus"

10210   With Me
10220     .legalname_pq.SelLength = 0
10230     .legalname_pq.SelStart = 99
10240   End With

EXITP:
10250   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub legalname_pq_AfterUpdate()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_pq_AfterUpdate"

        Dim strTmp01 As String, strTmp02 As String

        ' ** Single-quote, double-quote, printer's quote, printers quote, printers' quotes.
        ' **  '  Chr(39)
        ' **  "  Chr(34)
        ' **  ‘  Chr(145)
        ' **  ’  Chr(146)
        ' **  “  Chr(147)
        ' **  ”  Chr(148)

10310   With Me
10320     If IsNull(.legalname_pq) = False Then
10330       If Trim(.legalname_pq) <> vbNullString Then
10340         strTmp01 = Trim(.legalname_pq)
10350         If InStr(strTmp01, Chr(145)) = 0 And InStr(strTmp01, Chr(146)) = 0 And _
                  InStr(strTmp01, Chr(147)) = 0 And InStr(strTmp01, Chr(148)) = 0 Then
                ' ** They've removed the printers quotes entirely.
10360           strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
10370           If strTmp02 <> strTmp01 Then
                  ' ** And replaced them with regular quotes! Put them back.
10380             .legalname_pq = strTmp02
10390             .cotrustee.SetFocus
10400           Else
10410             .legalname.Visible = True
10420             .legalname.Enabled = True
10430             .cotrustee.SetFocus
10440             .legalname_pq.Visible = False
10450             .legalname_pq.Enabled = False
10460           End If
10470         Else
10480           strTmp02 = FixQuotes(strTmp01)  ' ** Module Function: modStringFuncs.
10490           If strTmp02 <> strTmp01 Then
10500             .legalname_pq = strTmp02
10510             .cotrustee.SetFocus
10520           End If
10530         End If
10540       End If
10550     End If
10560   End With

EXITP:
10570   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub legalname_pq_Change()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "legalname_pq_Change"

10610   AccountChange  ' ** Procedure: Below.

EXITP:
10620   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub cotrustee_Enter()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cotrustee_Enter"

10710   With Me
10720     If .cotrustee.text = vbNullString Then
10730       ExpandCombo .cotrustee  ' ** Module Procedure: modUtilities.
10740     End If
10750   End With

EXITP:
10760   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub cotrustee_NotInList(NewData As String, Response As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cotrustee_NotInList"

10810   With Me
10820     .cotrustee = Null
10830   End With

EXITP:
10840   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub cotrustee_Change()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cotrustee_Change"

10910   AccountChange  ' ** Procedure: Below.

EXITP:
10920   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub tin_Change()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "tin_Change"

11010   AccountChange  ' ** Procedure: Below.

EXITP:
11020   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub investmentobj_Change()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "investmentobj_Change"

11110   AccountChange  ' ** Procedure: Below.

EXITP:
11120   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub investmentobj_Enter()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "investmentobj_Enter"

11210   With Me
11220     If IsNull(.investmentobj) Then
11230       ExpandCombo .investmentobj  ' ** Module Procedure: modUtilities.
11240     End If
11250   End With

EXITP:
11260   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub accounttype_Enter()

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_Enter"

11310   With Me
11320     If .accounttype.text = vbNullString Then
11330       ExpandCombo .accounttype  ' ** Module Procedure: modUtilities.
11340     End If
11350   End With

EXITP:
11360   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub accounttype_NotInList(NewData As String, Response As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_NotInList"

11410   With Me
11420     .accounttype = Null
11430     MsgBox "Please select a valid account type.", vbInformation + vbOKOnly, "Invalid Entry"
11440     Response = acDataErrContinue
11450   End With

EXITP:
11460   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub accounttype_AfterUpdate()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "accounttype_AfterUpdate"

        Dim blnOkChange As Boolean

11510   blnOkChange = True  ' ** Unless proven otherwise.

11520   With Me
11530     If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
11540       blnOkChange = False
11550       .accounttype.Value = 99
11560       MsgBox "You cannot change this acount.", vbInformation + vbOKOnly, "Invalid Action"
11570       .accounttype.Undo
11580     ElseIf .accountno = "SUSPENSE" Or .accountno = "99-SUSPENSE" Then
11590       blnOkChange = False
11600       .accounttype.Value = 99
11610       MsgBox "You cannot change this acount.", vbInformation + vbOKOnly, "Invalid Action"
11620       .accounttype.Undo
11630     Else
11640       If gblnAccountNoWithType = True Then
              ' ** Verify the user entered a valid account type prefix to the account number.
11650         If InStr(.accountno, "-") > 0 Then
11660           If Mid(.accountno, 3, 1) = "-" Then
11670             If Left(.accountno, 2) <> .accounttype Then
11680               .accountno = .accounttype & Mid(accountno, 3)
11690             End If
11700           Else
11710             .accountno = .accounttype & "-" & .accountno
11720           End If
11730         Else
11740           .accountno = .accounttype & "-" & Trim(.accountno)
11750         End If
11760       End If
11770       If blnOkChange = True Then
11780         AccountChange  ' ** Procedure: Below.
11790         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
11800       End If
11810     End If
11820   End With

EXITP:
11830   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
880     Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
890     Case Else
900       THAT_PROC = THIS_PROC
910       That_Erl = Erl
920       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
930     End Select

940     Resume EXITP

End Sub

Private Sub description_GotFocus()
' ** Only accessible on a system account, and then it's locked.

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

11910   With Me
11920     .description.SelLength = 0
11930     .description.SelStart = 99
11940   End With

EXITP:
11950   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub cmbSweep_AfterUpdate()

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbSweep_AfterUpdate"

12010   With Me
12020     If .cmbSweep.text = "Yes" Then
12030       .account_SWEEP.Value = True
12040     Else
12050       .account_SWEEP.Value = False
12060     End If
12070   End With

EXITP:
12080   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub documentdate_BeforeUpdate(Cancel As Integer)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "documentdate_BeforeUpdate"

12110   With Me
12120     If .documentdate.text <> "__/__/____" And .documentdate.text <> vbNullString Then
12130       If Left(.documentdate.text, 2) > 12 Then
12140         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
12150         Cancel = -1
12160       End If
12170     End If
12180   End With

EXITP:
12190   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub documentdate_Change()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "documentdate_Change"

12210   With Me
12220     If IsDate(.documentdate) Then
12230       AccountChange  ' ** Procedure: Below.
12240     End If
12250   End With

EXITP:
12260   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub documentdate_Exit(Cancel As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "documentdate_Exit"

12310   ProcessMe  ' ** Procedure: Below.

EXITP:
12320   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub amendments_Enter()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "amendments_Enter"

12410   With Me
12420     If .amendments.text = vbNullString Then
12430       ExpandCombo .amendments  ' ** Module Procedure: modUtilities.
12440     End If
12450   End With

EXITP:
12460   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub amendments_NotInList(NewData As String, Response As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "amendments_NotInList"

12510   With Me
12520     .amendments = Null
12530   End With

EXITP:
12540   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub amendments_Change()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "amendments_Change"

12610   AccountChange  ' ** Procedure: Below.

EXITP:
12620   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub appointmentdate_BeforeUpdate(Cancel As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "appointmentdate_BeforeUpdate"

12710   With Me
12720     If .appointmentdate.text <> "__/__/____" And .appointmentdate.text <> vbNullString Then
12730       If Left(.appointmentdate.text, 2) > 12 Then
12740         MsgBox "Please enter a valid date.", vbInformation + vbOKOnly, "Invalid Date"
12750         Cancel = -1
12760       End If
12770     End If
12780   End With

EXITP:
12790   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub appointmentdate_Change()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "appointmentdate_Change"

12810   AccountChange  ' ** Procedure: Below.

EXITP:
12820   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub appointmentdate_Exit(Cancel As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "appointmentdate_Exit"

12910   ProcessMe  ' ** Procedure: Below.

EXITP:
12920   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub courtsupervised_Enter()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "courtsupervised_Enter"

13010   With Me
13020     If .courtsupervised.text = vbNullString Then
13030       ExpandCombo .courtsupervised  ' ** Module Procedure: modUtilities.
13040     End If
13050   End With

EXITP:
13060   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub courtsupervised_NotInList(NewData As String, Response As Integer)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "courtsupervised_NotInList"

13110   With Me
13120     .courtsupervised = Null
13130   End With

EXITP:
13140   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub courtsupervised_Change()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "courtsupervised_Change"

13210   AccountChange  ' ** Procedure: Below.

EXITP:
13220   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub courtsupervised_Exit(Cancel As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "courtsupervised_Exit"

13310   ProcessMe  ' ** Procedure: Below.

EXITP:
13320   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub jurisdiction_Enter()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "jurisdiction_Enter"

13410   With Me
13420     If .jurisdiction.text = vbNullString Then
13430       ExpandCombo .jurisdiction  ' ** Module Procedure: modUtilities.
13440     End If
13450   End With

EXITP:
13460   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub jurisdiction_NotInList(NewData As String, Response As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "jurisdiction_NotInList"

13510   With Me
13520     .jurisdiction = Null
13530     MsgBox "Must select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
13540     Response = acDataErrContinue
13550   End With

EXITP:
13560   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub jurisdiction_Change()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "jurisdiction_Change"

13610   AccountChange  ' ** Procedure: Below.

EXITP:
13620   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub jurisdiction_Exit(Cancel As Integer)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "jurisdiction_Exit"

13710   ProcessMe  ' ** Procedure: Below.

EXITP:
13720   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub county_Enter()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "county_Enter"

13810   cmdSave_Click THIS_PROC  ' ** Procedure: Above.

EXITP:
13820   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
880     Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
890     Case Else
900       THAT_PROC = THIS_PROC
910       That_Erl = Erl
920       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
930     End Select

940     Resume EXITP

End Sub

Private Sub county_Change()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "county_Change"

13910   AccountChange  ' ** Procedure: Below.

EXITP:
13920   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub county_Exit(Cancel As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "county_Exit"

14010   ProcessMe  ' ** Procedure: Below.

EXITP:
14020   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub cmdStatementFreq_Click()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStatementFreq_Click"

        Dim strDocName As String

        ' ** This routine is for any required processes.
14110   With Me
14120     If ProcessMe = True Then  ' ** Function: Below.
14130       strThisAccountNo = .accountno
14140       cmdSave_Click THIS_PROC  ' ** Procedure: Above.
14150       strDocName = "frmAccountProfile_StatementFreq"
14160       DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm & "~" & strThisAccountNo
14170     End If
14180   End With

EXITP:
14190   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
880     Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
890     Case Else
900       THAT_PROC = THIS_PROC
910       That_Erl = Erl
920       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
930     End Select
940     Resume EXITP

End Sub

Private Sub cmdStatementFreq_KeyDown(KeyCode As Integer, Shift As Integer)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStatementFreq_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14220   intShiftDown = (Shift And acShiftMask) > 0
14230   intAltDown = (Shift And acAltMask) > 0
14240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14260     Select Case intRetVal
          Case vbKeyTab
14270       With Me
14280         intRetVal = 0
14290         Select Case blnSysAcct
              Case True
14300           .FocusHolder.SetFocus
14310           .tabContacts = .tabContacts_pgNotes.PageIndex
14320           .Notes.SetFocus
14330           fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
14340         Case False
14350           .opgFeeFreq_optMonth.SetFocus
14360         End Select
14370       End With
14380     End Select
14390   End If

        ' ** Shift keys.
14400   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14410     Select Case intRetVal
          Case vbKeyTab
14420       With Me
14430         intRetVal = 0
14440         Select Case blnSysAcct
              Case True
14450           .description.SetFocus
14460         Case False
14470           .county.SetFocus
14480         End Select
14490       End With
14500     End Select
14510   End If

EXITP:
14520   KeyCode = intRetVal
14530   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub opgFeeFreq_optMonth_AfterUpdate()
' ** Value: 1.

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFeeFreq_optMonth_AfterUpdate"

14610   With Me
14620     .opgFeeFreq_opt_box.Visible = False
14630     Select Case .opgFeeFreq_optMonth
          Case True
14640       .opgFeeFreq_optQuarter = False
14650       .opgFeeFreq_optSemi = False
14660       .opgFeeFreq_optAnnual = False
14670       .feeFrequency = 1
14680       .opgFeeFreq_opt_box.Left = .opgFeeFreq_optMonth_lbl.Left
14690       .opgFeeFreq_opt_box.Visible = True
14700     Case False
14710       .feeFrequency = 0
14720     End Select
14730   End With

EXITP:
14740   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub opgFeeFreq_optQuarter_AfterUpdate()
' ** Value: 2.

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFeeFreq_optQuarter_AfterUpdate"

14810   With Me
14820     .opgFeeFreq_opt_box.Visible = False
14830     Select Case .opgFeeFreq_optQuarter
          Case True
14840       .opgFeeFreq_optMonth = False
14850       .opgFeeFreq_optSemi = False
14860       .opgFeeFreq_optAnnual = False
14870       .feeFrequency = 2
14880       .opgFeeFreq_opt_box.Left = .opgFeeFreq_optQuarter_lbl.Left
14890       .opgFeeFreq_opt_box.Visible = True
14900     Case False
14910       .feeFrequency = 0
14920     End Select
14930   End With

EXITP:
14940   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub opgFeeFreq_optSemi_AfterUpdate()
' ** Value: 3.

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFeeFreq_optSemi_AfterUpdate"

15010   With Me
15020     .opgFeeFreq_opt_box.Visible = False
15030     Select Case .opgFeeFreq_optSemi
          Case True
15040       .opgFeeFreq_optMonth = False
15050       .opgFeeFreq_optQuarter = False
15060       .opgFeeFreq_optAnnual = False
15070       .feeFrequency = 3
15080       .opgFeeFreq_opt_box.Left = .opgFeeFreq_optSemi_lbl.Left
15090       .opgFeeFreq_opt_box.Visible = True
15100     Case False
15110       .feeFrequency = 0
15120     End Select
15130   End With

EXITP:
15140   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub opgFeeFreq_optAnnual_AfterUpdate()
' ** Value: 4.

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFeeFreq_optAnnual_AfterUpdate"

15210   With Me
15220     .opgFeeFreq_opt_box.Visible = False
15230     Select Case .opgFeeFreq_optAnnual
          Case True
15240       .opgFeeFreq_optMonth = False
15250       .opgFeeFreq_optQuarter = False
15260       .opgFeeFreq_optSemi = False
15270       .feeFrequency = 4
15280       .opgFeeFreq_opt_box.Left = .opgFeeFreq_optAnnual_lbl.Left
15290       .opgFeeFreq_opt_box.Visible = True
15300     Case False
15310       .feeFrequency = 0
15320     End Select
15330   End With

EXITP:
15340   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub cmdReviewFreq_Click()

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReviewFreq_Click"

        Dim strDocName As String

15410   With Me
          ' ** This routine is for any required processes.
15420     If ProcessMe = True Then  ' ** Function: Below.
15430       strThisAccountNo = .accountno
15440       cmdSave_Click THIS_PROC  ' ** Procedure: Above.
15450       strDocName = "frmAccountProfile_ReviewFreq"
15460       DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm & "~" & strThisAccountNo
15470     End If
15480   End With

EXITP:
15490   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
880     Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
890     Case Else
900       THAT_PROC = THIS_PROC
910       That_Erl = Erl
920       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
930     End Select
940     Resume EXITP

End Sub

Private Sub Schedule_ID_Enter()

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_ID_Enter"

15510   With Me
15520     If .[Schedule_ID].text = vbNullString Then
15530       ExpandCombo .[Schedule_ID]  ' ** Module Procedure: modUtilities.
15540     End If
15550   End With

EXITP:
15560   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Schedule_ID_NotInList(NewData As String, Response As Integer)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_ID_NotInList"

15610   With Me
15620     .[Schedule_ID] = Null
15630   End With

EXITP:
15640   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Schedule_ID_Change()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_ID_Change"

15710   AccountChange  ' ** Procedure: Below.

EXITP:
15720   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Schedule_ID_Exit(Cancel As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_ID_Exit"

15810   ProcessMe  ' ** Procedure: Below.

EXITP:
15820   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub discretion_GotFocus()

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "discretion_GotFocus"

15910   With Me
15920     .discretion.SelLength = 0
15930     .discretion.SelStart = 9
15940   End With

EXITP:
15950   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub discretion_Enter()

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "discretion_Enter"

16010   With Me
16020     If .discretion.text = vbNullString Then
16030       ExpandCombo .discretion  ' ** Module Procedure: modUtilities.
16040     End If
16050   End With

EXITP:
16060   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub discretion_NotInList(NewData As String, Response As Integer)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "discretion_NotInList"

16110   With Me
16120     .discretion = Null
16130   End With

EXITP:
16140   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub discretion_Change()

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "discretion_Change"

16210   AccountChange  ' ** Procedure: Below.

EXITP:
16220   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub discretion_Exit(Cancel As Integer)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "discretion_Exit"

16310   ProcessMe  ' ** Procedure: Below.

EXITP:
16320   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub adminno_Enter()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "adminno_Enter"

16410   With Me
16420     If .adminno.text = vbNullString Then
16430       ExpandCombo .adminno  ' ** Module Procedure: modUtilities.
16440     End If
16450   End With

EXITP:
16460   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub adminno_NotInList(NewData As String, Response As Integer)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "adminno_NotInList"

16510   With Me
16520     .adminno = Null
16530     MsgBox "Please select an Administrator from the list.", vbInformation + vbOKOnly, "Invalid Entry"
16540     Response = acDataErrContinue
16550   End With

EXITP:
16560   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
900     Resume EXITP

End Sub

Private Sub adminno_KeyDown(KeyCode As Integer, Shift As Integer)

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "adminno_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16620   intShiftDown = (Shift And acShiftMask) > 0
16630   intAltDown = (Shift And acAltMask) > 0
16640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16670       With Me
16680         intRetVal = 0
16690         If .dateclosed.Enabled = True Then
16700           .dateclosed.SetFocus
16710         ElseIf .Bank_Name.Visible = True Then
16720           .Bank_Name.SetFocus
16730         Else
16740           .cmbAssets.SetFocus
16750         End If
16760       End With
16770     End Select
16780   End If

        ' ** Shift keys.
16790   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16800     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16810       With Me
16820         intRetVal = 0
16830         .discretion.SetFocus
16840       End With
16850     End Select
16860   End If

EXITP:
16870   KeyCode = intRetVal
16880   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub dateclosed_BeforeUpdate(Cancel As Integer)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "dateclosed_BeforeUpdate"

        Dim blnCashRound As Boolean, blnCostRound As Boolean
        Dim dblICash As Double, dblPCash As Double, dblCost As Double
        Dim msgResponse As VbMsgBoxResult

16910   With Me
16920     If IsNull(.dateclosed) = False Then
16930       If .ActiveAssets > 0 Then
              ' ** The ActiveAssets field comes from  a DLookup() in qryAccountProfile_01 to qryAccountProfile_02,
              ' ** which is the ActiveAssets table, grouped, with cnt.
16940         Cancel = -1
16950         MsgBox "This account still has " & CStr(.ActiveAssets) & " Active Assets.", vbInformation + vbOKOnly, "Invalid Entry"
16960         .dateclosed.SetFocus
16970       Else
16980         blnCashRound = False: blnCostRound = False
16990         If (Nz(.ICash, 0) + Nz(.PCash, 0)) <> 0 Then 'Or .Cost > 0 Then  ' ** Pull Cost out of the equation!
17000           dblICash = Nz(.ICash, 0)
17010           dblPCash = Nz(.PCash, 0)
17020           dblCost = Nz(.Cost, 0)  ' ** Though we can still check it.
17030           If ((dblICash + dblPCash) <= 0.01) And ((dblICash + dblPCash) >= -0.01) Then
17040             blnCashRound = True
17050           End If
17060           If (dblCost <= 0.01) And (dblCost >= -0.01) Then
17070             blnCostRound = True
17080           End If
17090           Select Case blnCashRound
                Case True
                  ' ** Just get rid of it without telling the user.
17100             .ICash = 0@
17110             .PCash = 0@
17120             If blnCostRound = True Then
17130               .Cost = 0@
17140             End If
17150             cmdSave_Click THIS_PROC  ' ** Procedure: Above.
17160           Case False
17170             Cancel = -1
17180             msgResponse = MsgBox("This account still has values listed in its master record:" & vbCrLf & _
                    "  Income Cash: " & Format(.ICash, "Currency") & vbCrLf & _
                    "  Principal Cash: " & Format(.PCash, "Currency") & vbCrLf & _
                    "  Cost: " & Format(.Cost, "Currency") & vbCrLf & vbCrLf & _
                    "Would you like to view the amounts listed, along with the actual Ledger totals?", _
                    vbQuestion + vbYesNo + vbDefaultButton1, "Warning: Account Not Zeroed Out")
17190             If msgResponse = vbYes Then
17200               gblnMessage = False
17210               DoCmd.OpenForm "frmMasterBalance", , , , , , strCallingForm
17220             End If
17230           End Select
17240         Else
17250           If .Bank_Name.Visible = True Then
17260             .Bank_Name.SetFocus
17270           Else
17280             DoCmd.GoToPage 2, 0, 0
17290             .Contact1_Name.SetFocus
17300           End If
17310         End If
17320       End If
17330     End If
17340   End With

EXITP:
17350   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
          ' ** Ignore.
880     Case 2108  ' ** You must save the field before you execute the GoToControl action, the GoToControl method, or the SetFocus method.
          ' ** Ignore.
890     Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** Ignore.
900     Case 2501  ' ** The '|' action was Canceled.
          ' ** Ignore.
910     Case Else
920       THAT_PROC = THIS_PROC
930       That_Erl = Erl
940       Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
950     End Select

960     Resume EXITP

End Sub

Private Sub dateclosed_KeyDown(KeyCode As Integer, Shift As Integer)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "dateclosed_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17420   intShiftDown = (Shift And acShiftMask) > 0
17430   intAltDown = (Shift And acAltMask) > 0
17440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17470       With Me
17480         intRetVal = 0
17490         If .Bank_Name.Visible = True Then
17500           .Bank_Name.SetFocus
17510         Else
17520           .cmbAssets.SetFocus
17530         End If
17540       End With
17550     End Select
17560   End If

        ' ** Shift keys.
17570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17590       With Me
17600         intRetVal = 0
17610         .adminno.SetFocus
17620       End With
17630     End Select
17640   End If

EXITP:
17650   KeyCode = intRetVal
17660   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub dateclosed_Exit(Cancel As Integer)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "dateclosed_Exit"

17710   ProcessMe  ' ** Procedure: Below.

EXITP:
17720   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Bank_Name_AfterUpdate()

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_Name_AfterUpdate"

        Dim strTmp01 As String

17810   With Me
17820     If IsNull(.Bank_Name) = False Then
17830       strTmp01 = FixQuotes(.Bank_Name)  ' ** Module Function: modStringFuncs.
17840       If Compare_StringA_StringB(.Bank_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
17850         .Bank_Name = strTmp01
17860         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
17870       End If
17880     End If
17890   End With

EXITP:
17900   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Bank_Name_Change()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_Name_Change"

18010   AccountChange  ' ** Procedure: Below.

EXITP:
18020   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Bank_City_AfterUpdate()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_City_AfterUpdate"

        Dim strTmp01 As String

18110   With Me
18120     If IsNull(.Bank_City) = False Then
18130       strTmp01 = FixQuotes(.Bank_City)  ' ** Module Function: modStringFuncs.
18140       If Compare_StringA_StringB(.Bank_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
18150         .Bank_City = strTmp01
18160         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
18170       End If
18180     End If
18190   End With

EXITP:
18200   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Bank_City_Change()

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_City_Change"

18310   AccountChange  ' ** Procedure: Below.

EXITP:
18320   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub opgLoc_AfterUpdate()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_AfterUpdate"

18410   With Me
18420     Select Case .opgLoc
          Case .opgLoc_optUSA.OptionValue
18430       .opgLoc_optUSA_lbl.FontBold = True
18440       .opgLoc_optOther_lbl.FontBold = False
18450       If blnSysAcct = False Then
18460         .Bank_State.Enabled = True
18470         .Bank_State.BorderColor = CLR_LTBLU2
18480         .Bank_State.BackStyle = acBackStyleNormal
18490         .Bank_State_lbl.BackStyle = acBackStyleNormal
18500         .Bank_State_lbl_box.Visible = False
18510         .Bank_Country.Enabled = False
18520         .Bank_Country.BorderColor = WIN_CLR_DISR
18530         .Bank_Country.BackStyle = acBackStyleTransparent
18540         .Bank_Country_lbl.BackStyle = acBackStyleTransparent
18550         .Bank_Country_lbl_box.Visible = True
18560         Select Case blnIsOpen
              Case True
18570           blnIsOpen = False
18580         Case False
18590           .Bank_State.SetFocus
18600         End Select
18610       End If
18620     Case .opgLoc_optOther.OptionValue
18630       .opgLoc_optUSA_lbl.FontBold = False
18640       .opgLoc_optOther_lbl.FontBold = True
18650       If blnSysAcct = False Then
18660         .Bank_State.Enabled = False
18670         .Bank_State.BorderColor = WIN_CLR_DISR
18680         .Bank_State.BackStyle = acBackStyleTransparent
18690         .Bank_State_lbl.BackStyle = acBackStyleTransparent
18700         .Bank_State_lbl_box.Visible = True
18710         .Bank_Country.Enabled = True
18720         .Bank_Country.BorderColor = CLR_LTBLU2
18730         .Bank_Country.BackStyle = acBackStyleNormal
18740         .Bank_Country_lbl.BackStyle = acBackStyleNormal
18750         .Bank_Country_lbl_box.Visible = False
18760         Select Case blnIsOpen
              Case True
18770           blnIsOpen = False
18780         Case False
18790           .Bank_Country.SetFocus
18800         End Select
18810       End If
18820     End Select
18830   End With

EXITP:
18840   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
890     End Select
900     Resume EXITP

End Sub

Private Sub opgLoc_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18920   intShiftDown = (Shift And acShiftMask) > 0
18930   intAltDown = (Shift And acAltMask) > 0
18940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18970       With Me
18980         intRetVal = 0
18990         If .Bank_State.Enabled = True Then
19000           .Bank_State.SetFocus
19010         Else
19020           .Bank_Country.SetFocus
19030         End If
19040       End With
19050     End Select
19060   End If

        ' ** Shift keys.
19070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19090       With Me
19100         intRetVal = 0
19110         .Bank_City.SetFocus
19120       End With
19130     End Select
19140   End If

EXITP:
19150   KeyCode = intRetVal
19160   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub opgLoc_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLoc_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19220   intShiftDown = (Shift And acShiftMask) > 0
19230   intAltDown = (Shift And acAltMask) > 0
19240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19270       With Me
19280         intRetVal = 0
19290         If .Bank_State.Enabled = True Then
19300           .Bank_State.SetFocus
19310         Else
19320           .Bank_Country.SetFocus
19330         End If
19340       End With
19350     End Select
19360   End If

        ' ** Shift keys.
19370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19390       With Me
19400         intRetVal = 0
19410         .Bank_City.SetFocus
19420       End With
19430     End Select
19440   End If

EXITP:
19450   KeyCode = intRetVal
19460   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub Bank_State_Enter()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_State_Enter"

19510   With Me
19520     If .Bank_State.text = vbNullString Then
19530       ExpandCombo .Bank_State  ' ** Module Procedure: modUtilities.
19540     End If
19550   End With

EXITP:
19560   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Bank_State_AfterUpdate()

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_State_AfterUpdate"

19610   With Me
19620     If IsNull(.Bank_State) = False Then
19630       If Trim(.Bank_State) <> vbNullString Then
19640         .Bank_State.Value = FormatUpperCase(Nz(.Bank_State, vbNullString))
19650         Select Case IsNull(.Acct_State_Pref)
              Case True
19660           .Acct_State_Pref = .Bank_State
19670         Case False
19680           If .Acct_State_Pref <> .Bank_State Then
19690             .Acct_State_Pref = .Bank_State
19700           End If
19710         End Select
19720         .Contact1_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
19730         .Contact2_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
19740       End If
19750     End If
19760   End With

EXITP:
19770   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Bank_State_Change()

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_State_Change"

19810   AccountChange  ' ** Procedure: Below.

EXITP:
19820   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Bank_State_NotInList(NewData As String, Response As Integer)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_State_NotInList"

19910   With Me
19920     .Bank_State = Null
19930     MsgBox "Please select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
19940     Response = acDataErrContinue
19950   End With

EXITP:
19960   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Bank_State_Exit(Cancel As Integer)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_State_Exit"

20010   ProcessMe  ' ** Procedure: Below.

EXITP:
20020   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Bank_RoutingNumber_Change()

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_RoutingNumber_Change"

20110   AccountChange  ' ** Procedure: Below.

EXITP:
20120   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Bank_RoutingNumber_AfterUpdate()

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_RoutingNumber_AfterUpdate"

20210   With Me
20220     If IsNull(.Bank_RoutingNumber) = False Then
20230       If Trim(.Bank_RoutingNumber) <> vbNullString Then
20240         If Len(Trim(.Bank_RoutingNumber)) > 9 Then
20250           .Bank_RoutingNumber = Left(Trim(.Bank_RoutingNumber), 9)
20260         End If
20270       End If
20280     End If
20290   End With

EXITP:
20300   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Bank_AccountNumber_Change()

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "Bank_AccountNumber_Change"

20410   AccountChange  ' ** Procedure: Below.

EXITP:
20420   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub FedIFNum1_Change()

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "FedIFNum1_Change"

20510   AccountChange  ' ** Procedure: Below.

EXITP:
20520   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub FedIfNum2_Change()

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "FedIfNum2_Change"

20610   AccountChange  ' ** Procedure: Below.

EXITP:
20620   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub cmbAssets_GotFocus()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_GotFocus"

20710   With Me
20720     .cmbAssets.SelLength = 0
20730     .cmbAssets.SelStart = 99
20740   End With

EXITP:
20750   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub cmbAssets_AfterUpdate()
' ** VGC 11/01/2010: For now, I'm using an unused field for the default asset.

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_AfterUpdate"

20810   With Me
20820     If IsNull(.cmbAssets) = False Then
20830       If .cmbAssets > 0& Then
20840         .taxlot = CStr(.cmbAssets)
20850       Else
20860         .taxlot = "0"
20870       End If
20880     Else
20890       .taxlot = "0"
20900     End If
20910   End With

EXITP:
20920   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub cmbAssets_NotInList(NewData As String, Response As Integer)

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_NotInList"

21010   With Me
21020     .cmbAssets = Null
21030     MsgBox "Please select an Asset from the list.", vbInformation + vbOKOnly, "Invalid Entry"
21040     Response = acDataErrContinue
21050   End With

EXITP:
21060   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub cmbAssets_KeyDown(KeyCode As Integer, Shift As Integer)

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAssets_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21120   intShiftDown = (Shift And acShiftMask) > 0
21130   intAltDown = (Shift And acAltMask) > 0
21140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21170       With Me
21180         intRetVal = 0
21190         .FocusHolder.SetFocus
21200         .tabContacts = .tabContacts_pgContact1.PageIndex
21210         .Contact1_Name.SetFocus
21220         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
21230       End With
21240     End Select
21250   End If

        ' ** Shift keys.
21260   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21270     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21280       With Me
21290         intRetVal = 0
21300         If .FedIFNum2.Visible = True Then
21310           .FedIFNum2.SetFocus
21320         ElseIf .dateclosed.Enabled = True Then
21330           .dateclosed.SetFocus
21340         Else
21350           .adminno.SetFocus
21360         End If
21370       End With
21380     End Select
21390   End If

EXITP:
21400   KeyCode = intRetVal
21410   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub PageBreak_1_to_2_lbl_Click()

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "PageBreak_1_to_2_lbl_Click"

21510   DoCmd.GoToPage 2, 0, 0

EXITP:
21520   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub PageBreak_1_to_2_lbl2_Click()

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "PageBreak_1_to_2_lbl2_Click"

21610   DoCmd.GoToPage 1, 0, 0

EXITP:
21620   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub tabContacts_Change()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "tabContacts_Change"

21710   DoCmd.GoToPage 2, 0, 0

EXITP:
21720   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 2163  ' ** The page number you used as an argument for the GoToPage action or method doesn't exist in the form.
          ' ** Ignore.
880     Case Else
890       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
900     End Select
910     Resume EXITP

End Sub

Private Sub Contact1_Name_AfterUpdate()

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Name_AfterUpdate"

        Dim strTmp01 As String

21810   With Me
21820     If IsNull(.Contact1_Name) = False Then
21830       strTmp01 = FixQuotes(.Contact1_Name)  ' ** Module Function: modStringFuncs.
21840       If Compare_StringA_StringB(.Contact1_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
21850         .Contact1_Name = strTmp01
21860         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
21870       End If
21880     End If
21890   End With

EXITP:
21900   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Name_Change()

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Name_Change"

22010   AccountChange  ' ** Procedure: Below.

EXITP:
22020   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Name_KeyDown(KeyCode As Integer, Shift As Integer)

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22120   intShiftDown = (Shift And acShiftMask) > 0
22130   intAltDown = (Shift And acAltMask) > 0
22140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22170       With Me
22180         intRetVal = 0
22190         DoCmd.GoToPage 2, 0, 0
22200         .FocusHolder.SetFocus
22210         .Contact1_Address1.SetFocus  ' ** This alone scrolls the page up a bit!
22220         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
22230       End With
22240     End Select
22250   End If

        ' ** Shift keys.
22260   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22270     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22280       With Me
22290         intRetVal = 0
22300         DoCmd.GoToPage 1, 0, 0
22310         .cmbAssets.SetFocus
22320         fSetScrollBarPosVT Me, 1&  ' ** Module Function: modScrollBarFuncs.
22330       End With
22340     End Select
22350   End If

EXITP:
22360   KeyCode = intRetVal
22370   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub Contact1_Name_Exit(Cancel As Integer)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Name_Exit"

22410   ProcessMe  ' ** Procedure: Below.

EXITP:
22420   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_Address1_AfterUpdate()

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Address1_AfterUpdate"

        Dim strTmp01 As String

22510   With Me
22520     If IsNull(.Contact1_Address1) = False Then
22530       strTmp01 = FixQuotes(.Contact1_Address1)  ' ** Module Function: modStringFuncs.
22540       If Compare_StringA_StringB(.Contact1_Address1, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
22550         .Contact1_Address1 = strTmp01
22560         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
22570       End If
22580     End If
22590   End With

EXITP:
22600   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Address1_Change()

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Address1_Change"

22710   AccountChange  ' ** Procedure: Below.

EXITP:
22720   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Address1_Exit(Cancel As Integer)

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Address1_Exit"

22810   ProcessMe  ' ** Procedure: Below.

EXITP:
22820   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_Address2_AfterUpdate()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Address2_AfterUpdate"

        Dim strTmp01 As String

22910   With Me
22920     If IsNull(.Contact1_Address2) = False Then
22930       strTmp01 = FixQuotes(.Contact1_Address2)  ' ** Module Function: modStringFuncs.
22940       If Compare_StringA_StringB(.Contact1_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
22950         .Contact1_Address2 = strTmp01
22960         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
22970       End If
22980     End If
22990   End With

EXITP:
23000   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Address2_Change()

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Address2_Change"

23110   AccountChange  ' ** Procedure: Below.

EXITP:
23120   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Address2_Exit(Cancel As Integer)

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Address2_Exit"

23210   ProcessMe  ' ** Procedure: Below.

EXITP:
23220   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_City_AfterUpdate()

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_City_AfterUpdate"

        Dim strTmp01 As String

23310   With Me
23320     If IsNull(.Contact1_City) = False Then
23330       strTmp01 = FixQuotes(.Contact1_City)  ' ** Module Function: modStringFuncs.
23340       If Compare_StringA_StringB(.Contact1_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
23350         .Contact1_City = strTmp01
23360         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
23370       End If
23380     End If
23390   End With

EXITP:
23400   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_City_Change()

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_City_Change"

23510   AccountChange  ' ** Procedure: Below.

EXITP:
23520   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_City_Exit(Cancel As Integer)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_City_Exit"

23610   ProcessMe  ' ** Procedure: Below.

EXITP:
23620   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Public Sub opgLocContact1_AfterUpdate()

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact1_AfterUpdate"

23710   With Me
23720     Select Case .opgLocContact1
          Case .opgLocContact1_optUSA.OptionValue
23730       .FocusHolder.SetFocus
23740       .Contact1_State.Enabled = True
23750       .Contact1_State.BackStyle = acBackStyleNormal
23760       .Contact1_State.BorderColor = CLR_LTBLU2
23770       .Contact1_State_lbl.BackStyle = acBackStyleNormal
23780       .Contact1_Zip.Enabled = True
23790       .Contact1_Zip.BackStyle = acBackStyleNormal
23800       .Contact1_Zip.BorderColor = CLR_LTBLU2
23810       .Contact1_Zip_lbl.BackStyle = acBackStyleNormal
23820       .Contact1_Country.Enabled = False
23830       .Contact1_Country.BackStyle = acBackStyleTransparent
23840       .Contact1_Country.BorderColor = WIN_CLR_DISR
23850       .Contact1_Country_lbl.BackStyle = acBackStyleTransparent
23860       .Contact1_PostalCode.Enabled = False
23870       .Contact1_PostalCode.BackStyle = acBackStyleTransparent
23880       .Contact1_PostalCode.BorderColor = WIN_CLR_DISR
23890       .Contact1_PostalCode_lbl.BackStyle = acBackStyleTransparent
23900       Select Case blnIsOpen
            Case True
23910         blnIsOpen = False
23920       Case False
23930         .Contact1_State.SetFocus
23940         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
23950       End Select
23960     Case .opgLocContact1_optOther.OptionValue
23970       .FocusHolder.SetFocus
23980       .Contact1_Country.Enabled = True
23990       .Contact1_Country.BackStyle = acBackStyleNormal
24000       .Contact1_Country.BorderColor = CLR_LTBLU2
24010       .Contact1_Country_lbl.BackStyle = acBackStyleNormal
24020       .Contact1_PostalCode.Enabled = True
24030       .Contact1_PostalCode.BackStyle = acBackStyleNormal
24040       .Contact1_PostalCode.BorderColor = CLR_LTBLU2
24050       .Contact1_PostalCode_lbl.BackStyle = acBackStyleNormal
24060       .Contact1_State.Enabled = False
24070       .Contact1_State.BackStyle = acBackStyleTransparent
24080       .Contact1_State.BorderColor = WIN_CLR_DISR
24090       .Contact1_State_lbl.BackStyle = acBackStyleTransparent
24100       .Contact1_Zip.Enabled = False
24110       .Contact1_Zip.BackStyle = acBackStyleTransparent
24120       .Contact1_Zip.BorderColor = WIN_CLR_DISR
24130       .Contact1_Zip_lbl.BackStyle = acBackStyleTransparent
24140       Select Case blnIsOpen
            Case True
24150         blnIsOpen = False
24160       Case False
24170         .Contact1_Country.SetFocus
24180         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
24190       End Select
24200     End Select
24210   End With

EXITP:
24220   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub opgLocContact1_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact1_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24320   intShiftDown = (Shift And acShiftMask) > 0
24330   intAltDown = (Shift And acAltMask) > 0
24340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24370       With Me
24380         intRetVal = 0
24390         If .Contact1_State.Enabled = True Then
24400           .Contact1_State.SetFocus
24410         Else
24420           .Contact1_Country.SetFocus
24430         End If
24440       End With
24450     End Select
24460   End If

        ' ** Shift keys.
24470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24490       With Me
24500         intRetVal = 0
24510         .Contact1_City.SetFocus
24520       End With
24530     End Select
24540   End If

EXITP:
24550   KeyCode = intRetVal
24560   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub opgLocContact1_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact1_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24620   intShiftDown = (Shift And acShiftMask) > 0
24630   intAltDown = (Shift And acAltMask) > 0
24640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24670       With Me
24680         intRetVal = 0
24690         If .Contact1_State.Enabled = True Then
24700           .Contact1_State.SetFocus
24710         Else
24720           .Contact1_Country.SetFocus
24730         End If
24740       End With
24750     End Select
24760   End If

        ' ** Shift keys.
24770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24790       With Me
24800         intRetVal = 0
24810         .Contact1_City.SetFocus
24820       End With
24830     End Select
24840   End If

EXITP:
24850   KeyCode = intRetVal
24860   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub Contact1_State_Enter()

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_State_Enter"

24910   With Me
24920     If .Contact1_State.text = vbNullString Then
24930       ExpandCombo .Contact1_State  ' ** Module Procedure: modUtilities.
24940     End If
24950   End With

EXITP:
24960   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_State_AfterUpdate()

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_State_AfterUpdate"

25010   With Me
25020     If IsNull(.Contact1_State) = False Then
25030       If Trim(.Contact1_State) <> vbNullString Then
25040         .Contact1_State.Value = FormatUpperCase(Nz(.Contact1_State, vbNullString))
25050         Select Case IsNull(.Acct_State_Pref)
              Case True
25060           .Acct_State_Pref = .Contact1_State
25070         Case False
25080           If .Acct_State_Pref <> .Contact1_State Then
25090             .Acct_State_Pref = .Contact1_State
25100           End If
25110         End Select
25120         .Bank_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
25130         .Contact2_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
25140       End If
25150     End If
25160   End With

EXITP:
25170   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_State_Change()

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_State_Change"

25210   AccountChange  ' ** Procedure: Below.

EXITP:
25220   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_State_NotInList(NewData As String, Response As Integer)

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_State_NotInList"

25310   With Me
25320     .Contact1_State = Null
25330     MsgBox "Please select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
25340     Response = acDataErrContinue
25350   End With

EXITP:
25360   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_State_Exit(Cancel As Integer)

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_State_Exit"

25410   ProcessMe  ' ** Procedure: Below.

EXITP:
25420   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_Zip_Exit(Cancel As Integer)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Zip_Exit"

25510   With Me
25520     If .Contact1_Zip <> "" Then
25530       If Right(.Contact1_Zip.text, 4) <> "____" And Right(.Contact1_Zip.text, 4) <> "    " And _
                (Not IsNumeric(Mid(.Contact1_Zip.text, 7, 1)) Or Not IsNumeric(Mid(.Contact1_Zip.text, 8, 1)) Or _
                Not IsNumeric(Mid(.Contact1_Zip.text, 9, 1)) Or Not IsNumeric(Mid(.Contact1_Zip.text, 10, 1))) Then
25540         MsgBox "Please enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
25550         Cancel = -1
25560       Else
25570         ProcessMe  ' ** Procedure: Below.
25580       End If
25590     End If
25600   End With

EXITP:
25610   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_Country_Change()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Country_Change"

25710   AccountChange  ' ** Procedure: Below.

EXITP:
25720   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Country_Exit(Cancel As Integer)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Country_Exit"

25810   ProcessMe  ' ** Procedure: Below.

EXITP:
25820   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_PostalCode_Change()

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_PostalCode_Change"

25910   AccountChange  ' ** Procedure: Below.

EXITP:
25920   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_PostalCode_Exit(Cancel As Integer)

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_PostalCode_Exit"

26010   ProcessMe  ' ** Procedure: Below.

EXITP:
26020   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_Phone1_Change()

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Phone1_Change"

26110   AccountChange  ' ** Procedure: Below.

EXITP:
26120   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Phone1_Exit(Cancel As Integer)

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Phone1_Exit"

26210   ProcessMe  ' ** Procedure: Below.

EXITP:
26220   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_Phone2_Change()

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Phone2_Change"

26310   AccountChange  ' ** Procedure: Below.

EXITP:
26320   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Phone2_Exit(Cancel As Integer)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Phone2_Exit"

26410   ProcessMe  ' ** Procedure: Below.

EXITP:
26420   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_Fax_Change()

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Fax_Change"

26510   AccountChange  ' ** Procedure: Below.

EXITP:
26520   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Fax_Exit(Cancel As Integer)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Fax_Exit"

26610   ProcessMe  ' ** Procedure: Below.

EXITP:
26620   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact1_Email_KeyDown(KeyCode As Integer, Shift As Integer)

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Email_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26720   intShiftDown = (Shift And acShiftMask) > 0
26730   intAltDown = (Shift And acAltMask) > 0
26740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26770       With Me
26780         intRetVal = 0
26790         .GoToPage 1, 0, 0
26800         .accountno.SetFocus  ' ** To make sure the Detail gets scrolled up.
26810         fSetScrollBarPosVT Me, 1&  ' ** Module Function: modScrollBarFuncs.
26820       End With
26830     End Select
26840   End If

        ' ** Shift keys.
26850   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26870       With Me
26880         DoCmd.GoToPage 2, 0, 0
26890         .FocusHolder.SetFocus
26900         .Contact1_Fax.SetFocus  ' ** This alone scrolls the page up a bit!
26910         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
26920       End With
26930     End Select
26940   End If

EXITP:
26950   KeyCode = intRetVal
26960   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub Contact1_Email_Change()

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Email_Change"

27010   AccountChange  ' ** Procedure: Below.

EXITP:
27020   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact1_Email_Exit(Cancel As Integer)

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact1_Email_Exit"

27110   ProcessMe  ' ** Procedure: Below.

EXITP:
27120   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_Name_AfterUpdate()

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Name_AfterUpdate"

        Dim strTmp01 As String

27210   With Me
27220     If IsNull(.Contact2_Name) = False Then
27230       strTmp01 = FixQuotes(.Contact2_Name)  ' ** Module Function: modStringFuncs.
27240       If Compare_StringA_StringB(.Contact2_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
27250         .Contact2_Name = strTmp01
27260         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
27270       End If
27280     End If
27290   End With

EXITP:
27300   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Name_KeyDown(KeyCode As Integer, Shift As Integer)

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Name_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27420   intShiftDown = (Shift And acShiftMask) > 0
27430   intAltDown = (Shift And acAltMask) > 0
27440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27470       With Me
27480         intRetVal = 0
27490         .Contact2_Address1.SetFocus
27500       End With
27510     End Select
27520   End If

        ' ** Shift keys.
27530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27550       With Me
27560         intRetVal = 0
27570         .Contact1_Email.SetFocus
27580       End With
27590     End Select
27600   End If

EXITP:
27610   KeyCode = intRetVal
27620   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub Contact2_Name_Change()

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Name_Change"

27710   AccountChange  ' ** Procedure: Below.

EXITP:
27720   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Name_Exit(Cancel As Integer)

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Name_Exit"

27810   ProcessMe  ' ** Procedure: Below.

EXITP:
27820   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_Address1_AfterUpdate()

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Address1_AfterUpdate"

        Dim strTmp01 As String

27910   With Me
27920     If IsNull(.Contact2_Address1) = False Then
27930       strTmp01 = FixQuotes(.Contact2_Address1)  ' ** Module Function: modStringFuncs.
27940       If Compare_StringA_StringB(.Contact2_Address1, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
27950         .Contact2_Address1 = strTmp01
27960         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
27970       End If
27980     End If
27990   End With

EXITP:
28000   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Address1_Change()

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Address1_Change"

28110   AccountChange  ' ** Procedure: Below.

EXITP:
28120   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Address1_Exit(Cancel As Integer)

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Address1_Exit"

28210   ProcessMe  ' ** Procedure: Below.

EXITP:
28220   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_Address2_AfterUpdate()

28300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Address2_AfterUpdate"

        Dim strTmp01 As String

28310   With Me
28320     If IsNull(.Contact2_Address2) = False Then
28330       strTmp01 = FixQuotes(.Contact2_Address2)  ' ** Module Function: modStringFuncs.
28340       If Compare_StringA_StringB(.Contact2_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
28350         .Contact2_Address2 = strTmp01
28360         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
28370       End If
28380     End If
28390   End With

EXITP:
28400   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Address2_Change()

28500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Address2_Change"

28510   AccountChange  ' ** Procedure: Below.

EXITP:
28520   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Address2_Exit(Cancel As Integer)

28600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Address2_Exit"

28610   ProcessMe  ' ** Procedure: Below.

EXITP:
28620   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_City_AfterUpdate()

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_City_AfterUpdate"

        Dim strTmp01 As String

28710   With Me
28720     If IsNull(.Contact2_City) = False Then
28730       strTmp01 = FixQuotes(.Contact2_City)  ' ** Module Function: modStringFuncs.
28740       If Compare_StringA_StringB(.Contact2_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
28750         .Contact2_City = strTmp01
28760         cmdSave_Click THIS_PROC  ' ** Procedure: Above.
28770       End If
28780     End If
28790   End With

EXITP:
28800   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_City_Change()

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_City_Change"

28910   AccountChange  ' ** Procedure: Below.

EXITP:
28920   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_City_Exit(Cancel As Integer)

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_City_Exit"

29010   ProcessMe  ' ** Procedure: Below.

EXITP:
29020   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Public Sub opgLocContact2_AfterUpdate()

29100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact2_AfterUpdate"

29110   With Me
29120     Select Case .opgLocContact2
          Case .opgLocContact2_optUSA.OptionValue
29130       .FocusHolder.SetFocus
29140       .Contact2_State.Enabled = True
29150       .Contact2_State.BackStyle = acBackStyleNormal
29160       .Contact2_State.BorderColor = CLR_LTBLU2
29170       .Contact2_State_lbl.BackStyle = acBackStyleNormal
29180       .Contact2_Zip.Enabled = True
29190       .Contact2_Zip.BackStyle = acBackStyleNormal
29200       .Contact2_Zip.BorderColor = CLR_LTBLU2
29210       .Contact2_Zip_lbl.BackStyle = acBackStyleNormal
29220       .Contact2_Country.Enabled = False
29230       .Contact2_Country.BackStyle = acBackStyleTransparent
29240       .Contact2_Country.BorderColor = WIN_CLR_DISR
29250       .Contact2_Country_lbl.BackStyle = acBackStyleTransparent
29260       .Contact2_PostalCode.Enabled = False
29270       .Contact2_PostalCode.BackStyle = acBackStyleTransparent
29280       .Contact2_PostalCode.BorderColor = WIN_CLR_DISR
29290       .Contact2_PostalCode_lbl.BackStyle = acBackStyleTransparent
29300       Select Case blnIsOpen
            Case True
29310         blnIsOpen = False
29320       Case False
29330         .Contact2_State.SetFocus
29340         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
29350       End Select
29360     Case .opgLocContact2_optOther.OptionValue
29370       .FocusHolder.SetFocus
29380       .Contact2_Country.Enabled = True
29390       .Contact2_Country.BackStyle = acBackStyleNormal
29400       .Contact2_Country.BorderColor = CLR_LTBLU2
29410       .Contact2_Country_lbl.BackStyle = acBackStyleNormal
29420       .Contact2_PostalCode.Enabled = True
29430       .Contact2_PostalCode.BackStyle = acBackStyleNormal
29440       .Contact2_PostalCode.BorderColor = CLR_LTBLU2
29450       .Contact2_PostalCode_lbl.BackStyle = acBackStyleNormal
29460       .Contact2_State.Enabled = False
29470       .Contact2_State.BackStyle = acBackStyleTransparent
29480       .Contact2_State.BorderColor = WIN_CLR_DISR
29490       .Contact2_State_lbl.BackStyle = acBackStyleTransparent
29500       .Contact2_Zip.Enabled = False
29510       .Contact2_Zip.BackStyle = acBackStyleTransparent
29520       .Contact2_Zip.BorderColor = WIN_CLR_DISR
29530       .Contact2_Zip_lbl.BackStyle = acBackStyleTransparent
29540       Select Case blnIsOpen
            Case True
29550         blnIsOpen = False
29560       Case False
29570         .Contact2_Country.SetFocus
29580         fSetScrollBarPosVT Me, 999&  ' ** Module Function: modScrollBarFuncs.
29590       End Select
29600     End Select
29610   End With

EXITP:
29620   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub opgLocContact2_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

29700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact2_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29720   intShiftDown = (Shift And acShiftMask) > 0
29730   intAltDown = (Shift And acAltMask) > 0
29740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29770       With Me
29780         intRetVal = 0
29790         If .Contact2_State.Enabled = True Then
29800           .Contact2_State.SetFocus
29810         Else
29820           .Contact2_Country.SetFocus
29830         End If
29840       End With
29850     End Select
29860   End If

        ' ** Shift keys.
29870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29880     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29890       With Me
29900         intRetVal = 0
29910         .Contact2_City.SetFocus
29920       End With
29930     End Select
29940   End If

EXITP:
29950   KeyCode = intRetVal
29960   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub opgLocContact2_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact2_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30020   intShiftDown = (Shift And acShiftMask) > 0
30030   intAltDown = (Shift And acAltMask) > 0
30040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30070       With Me
30080         intRetVal = 0
30090         If .Contact2_State.Enabled = True Then
30100           .Contact2_State.SetFocus
30110         Else
30120           .Contact2_Country.SetFocus
30130         End If
30140       End With
30150     End Select
30160   End If

        ' ** Shift keys.
30170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30190       With Me
30200         intRetVal = 0
30210         .Contact2_City.SetFocus
30220       End With
30230     End Select
30240   End If

EXITP:
30250   KeyCode = intRetVal
30260   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub Contact2_State_Enter()

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_State_Enter"

30310   With Me
30320     If .Contact2_State.text = vbNullString Then
30330       ExpandCombo .Contact2_State  ' ** Module Procedure: modUtilities.
30340     End If
30350   End With

EXITP:
30360   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_State_NotInList(NewData As String, Response As Integer)

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_State_NotInList"

30410   With Me
30420     .Contact2_State = Null
30430     MsgBox "Please select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
30440     Response = acDataErrContinue
30450   End With

EXITP:
30460   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_State_AfterUpdate()

30500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_State_AfterUpdate"

30510   With Me
30520     If IsNull(.Contact2_State) = False Then
30530       If Trim(.Contact2_State) <> vbNullString Then
30540         .Contact2_State.Value = FormatUpperCase(Nz(.Contact2_State, vbNullString))
30550         Select Case IsNull(.Acct_State_Pref)
              Case True
30560           .Acct_State_Pref = .Contact2_State
30570         Case False
30580           If .Acct_State_Pref <> .Contact2_State Then
30590             .Acct_State_Pref = .Contact2_State
30600           End If
30610         End Select
30620         .Bank_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
30630         .Contact1_State.DefaultValue = Chr(34) & .Acct_State_Pref & Chr(34)
30640       End If
30650     End If
30660   End With

EXITP:
30670   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_State_Change()

30700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_State_Change"

30710   AccountChange  ' ** Procedure: Below.

EXITP:
30720   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_State_Exit(Cancel As Integer)

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_State_Exit"

30810   ProcessMe  ' ** Procedure: Below.

EXITP:
30820   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_Zip_Change()

30900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Zip_Change"

30910   AccountChange  ' ** Procedure: Below.

EXITP:
30920   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Zip_Exit(Cancel As Integer)

31000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Zip_Exit"

31010   With Me
31020     If .Contact2_Zip <> "" Then
31030       If Right(.Contact2_Zip.text, 4) <> "____" And Right(.Contact2_Zip.text, 4) <> "    " And _
                (Not IsNumeric(Mid(.Contact2_Zip.text, 7, 1)) Or Not IsNumeric(Mid(.Contact2_Zip.text, 8, 1)) Or _
                Not IsNumeric(Mid(.Contact2_Zip.text, 9, 1)) Or Not IsNumeric(Mid(.Contact2_Zip.text, 10, 1))) Then
31040         MsgBox "Please enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
31050         Cancel = -1
31060       Else
31070         ProcessMe  ' ** Procedure: Below.
31080       End If
31090     End If
31100   End With

EXITP:
31110   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Country_Change()

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Country_Change"

31210   AccountChange  ' ** Procedure: Below.

EXITP:
31220   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Country_Exit(Cancel As Integer)

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Country_Exit"

31310   ProcessMe  ' ** Procedure: Below.

EXITP:
31320   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_PostalCode_Change()

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_PostalCode_Change"

31410   AccountChange  ' ** Procedure: Below.

EXITP:
31420   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_PostalCode_Exit(Cancel As Integer)

31500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_PostalCode_Exit"

31510   ProcessMe  ' ** Procedure: Below.

EXITP:
31520   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_Phone1_Change()

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Phone1_Change"

31610   AccountChange  ' ** Procedure: Below.

EXITP:
31620   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Phone1_Exit(Cancel As Integer)

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Phone1_Exit"

31710   ProcessMe  ' ** Procedure: Below.

EXITP:
31720   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_Phone2_Change()

31800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Phone2_Change"

31810   AccountChange  ' ** Procedure: Below.

EXITP:
31820   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Phone2_Exit(Cancel As Integer)

31900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Phone2_Exit"

31910   ProcessMe  ' ** Procedure: Below.

EXITP:
31920   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_Fax_Change()

32000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Fax_Change"

32010   AccountChange  ' ** Procedure: Below.

EXITP:
32020   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Fax_Exit(Cancel As Integer)

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Fax_Exit"

32110   ProcessMe  ' ** Procedure: Below.

EXITP:
32120   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Contact2_Email_KeyDown(KeyCode As Integer, Shift As Integer)

32200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Email_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
32220   intShiftDown = (Shift And acShiftMask) > 0
32230   intAltDown = (Shift And acAltMask) > 0
32240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32270       With Me
32280         intRetVal = 0
32290         .Notes.SetFocus
32300       End With
32310     End Select
32320   End If

        ' ** Shift keys.
32330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32350       With Me
32360         intRetVal = 0
32370         .Contact2_Fax.SetFocus
32380       End With
32390     End Select
32400   End If

EXITP:
32410   KeyCode = intRetVal
32420   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub Contact2_Email_Change()

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Email_Change"

32510   AccountChange  ' ** Procedure: Below.

EXITP:
32520   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Contact2_Email_Exit(Cancel As Integer)

32600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact2_Email_Exit"

32610   ProcessMe  ' ** Procedure: Below.

EXITP:
32620   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub Notes_GotFocus()

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "Notes_GotFocus"

32710   With Me
32720     .Notes.SelLength = 0
32730     .Notes.SelStart = 999
32740   End With

EXITP:
32750   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Notes_Change()

32800 On Error GoTo ERRH

        Const THIS_PROC As String = "Notes_Change"

32810   AccountChange  ' ** Procedure: Below.

EXITP:
32820   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Private Sub Notes_KeyDown(KeyCode As Integer, Shift As Integer)

32900 On Error GoTo ERRH

        Const THIS_PROC As String = "Notes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
32920   intShiftDown = (Shift And acShiftMask) > 0
32930   intAltDown = (Shift And acAltMask) > 0
32940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32960     Select Case intRetVal
          Case vbKeyTab
32970       With Me
32980         intRetVal = 0
32990         Select Case blnSysAcct
              Case True
33000           DoCmd.SelectObject acForm, .Parent.Name, False
33010           .Parent.cmdClose.SetFocus
33020         Case False
33030           DoCmd.SelectObject acForm, .Parent.Name, False
33040           .Parent.cmdClose.SetFocus
33050         End Select
33060       End With
33070     End Select
33080   End If

        ' ** Shift keys.
33090   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33100     Select Case intRetVal
          Case vbKeyTab
33110       With Me
33120         intRetVal = 0
33130         Select Case blnSysAcct
              Case True
33140           DoCmd.GoToPage 1, 0, 0
33150           .cmdStatementFreq.SetFocus
33160           fSetScrollBarPosVT Me, 1&  ' ** Module Function: modScrollBarFuncs.
33170         Case False
33180           .Contact2_Email.SetFocus
33190         End Select
33200       End With
33210     End Select
33220   End If

EXITP:
33230   KeyCode = intRetVal
33240   Exit Sub

ERRH:
870     intRetVal = 0
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Sub

Private Sub Notes_Exit(Cancel As Integer)

33300 On Error GoTo ERRH

        Const THIS_PROC As String = "Notes_Exit"

33310   ProcessMe  ' ** Procedure: Below.

EXITP:
33320   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Public Sub AccountChange()

33400 On Error GoTo ERRH

        Const THIS_PROC As String = "AccountChange"

        Dim strControl As String

33410 On Error Resume Next
33420   strControl = Screen.ActiveControl.Name
33430 On Error GoTo ERRH

33440   With Me
33450     blnAccountHasChanged = True
33460     .Parent.AcctChanged_Set blnAccountHasChanged  ' ** Form Procedure: frmAccountProfile.
33470     If IsNull(.accountno) And strControl <> "accountno" Then
33480       MsgBox "Please enter a valid account number.", vbInformation + vbOKOnly, "Entry Required"
33490       .shortname.SetFocus
33500     End If
33510   End With

EXITP:
33520   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Public Function ProcessMe() As Boolean

33600 On Error GoTo ERRH

        Const THIS_PROC As String = "ProcessMe"

        Dim blnRetVal As Boolean

33610   blnRetVal = True

33620   With Me
33630     If IsNull(.accounttype) Then
33640       blnRetVal = False
33650       MsgBox "Please enter an account type to continue", vbInformation + vbOKOnly, "Entry Required"
33660       DoCmd.CancelEvent
33670       .GoToPage 1, 0, 0
33680       .accounttype.SetFocus
33690     End If
33700   End With

EXITP:
33710   ProcessMe = blnRetVal
33720   Exit Function

ERRH:
870     blnRetVal = False
880     THAT_PROC = THIS_PROC
890     That_Erl = Erl
900     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
910     Resume EXITP

End Function

Public Sub Exit_Set(blnLeaving As Boolean)

33800 On Error GoTo ERRH

        Const THIS_PROC As String = "Exit_Set"

33810   blnExit = blnLeaving

EXITP:
33820   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Public Sub EscPressed_Set(blnEscape As Boolean)

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "EscPressed_Set"

33910   blnEscPressed = blnEscape

EXITP:
33920   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Public Sub NewRec_Get()

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "NewRec_Get"

34010   With Me
34020     If .NewRecord = False Then
34030       DoCmd.GoToRecord , , acNewRec
34040     End If
34050   End With

EXITP:
34060   Exit Sub

ERRH:
870     THAT_PROC = THIS_PROC
880     That_Erl = Erl
890     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
900     Resume EXITP

End Sub

Public Sub ViewOnly_Set(blnView As Boolean)

34100 On Error GoTo ERRH

        Const THIS_PROC As String = "ViewOnly_Set"

34110   With Me
34120     blnViewOnly = blnView
34130     .ViewOnly = blnView
34140   End With

EXITP:
34150   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Function RelAcct_Chk() As Boolean

34200 On Error GoTo ERRH

        Const THIS_PROC As String = "RelAcct_Chk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim strAcct As String, strRelAccts As String
        Dim lngRels As Long, arr_varRel() As Variant
        Dim lngAccts As Long, arr_varAcct() As Variant
        Dim lngRelAlls As Long, arr_varRelAll() As Variant
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, arr_varTmp01 As Variant, strTmp02 As String, strTmp03 As String, lngTmp04 As Long, blnTmp05 As Boolean
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

        Const A_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const A_ACCT As Integer = 0
        Const A_RELS As Integer = 1
        Const A_RCNT As Integer = 2
        Const A_ARR  As Integer = 3
        Const A_NEW  As Integer = 4
        Const A_FND  As Integer = 5

        Const RA_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const RA_ACCT As Integer = 0
        Const RA_FND  As Integer = 1
        Const RA_DUPE As Integer = 2

        Const R_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const R_ACCT As Integer = 0
        Const R_FND  As Integer = 1
        Const R_DUPE As Integer = 2

34210   blnRetVal = True

34220   lngAccts = 0&
34230   ReDim arr_varRel(A_ELEMS, 0)
34240   lngRels = 0&
34250   ReDim arr_varRel(R_ELEMS, 0)
34260   lngRelAlls = 0&
34270   ReDim arr_varRelAll(RA_ELEMS, 0)

34280   Set dbs = CurrentDb
34290   With dbs

          ' ** Update Account, set related_accountno = Null where vbNullString.
34300     Set qdf = .QueryDefs("qryAccountProfile_08")
34310     qdf.Execute

          ' ** Account, all non-Null related_accountno.
34320     Set qdf = .QueryDefs("qryAccountProfile_09")
34330     Set rst = qdf.OpenRecordset
34340     With rst
34350       If .BOF = True And .EOF = True Then
              ' ** No related accounts specified.
34360       Else
34370         .MoveLast
34380         lngRecs = .RecordCount
34390         .MoveFirst
34400         For lngX = 1& To lngRecs
34410           lngRels = 0&
34420           ReDim arr_varRel(R_ELEMS, 0)
34430           strAcct = ![accountno]
34440           strRelAccts = Trim(![related_accountno])
34450           lngAccts = lngAccts + 1&
34460           lngE = lngAccts - 1&
34470           ReDim Preserve arr_varAcct(A_ELEMS, lngE)
34480           arr_varAcct(A_ACCT, lngE) = strAcct
34490           arr_varAcct(A_RELS, lngE) = strRelAccts
34500           arr_varAcct(A_RCNT, lngE) = CLng(0)
34510           arr_varAcct(A_NEW, lngE) = vbNullString
34520           strTmp02 = strRelAccts
34530           intPos01 = InStr(strTmp02, " ")  ' ** Check for space, used in previous TA versions.
34540           Do While intPos01 > 0
34550             strTmp02 = Trim(Left(strTmp02, (intPos01 - 1))) & "," & Trim(Mid(strTmp02, (intPos01 + 1)))
34560             intPos01 = InStr(strTmp02, " ")
34570           Loop
34580           intPos01 = InStr(strTmp02, ",")
34590           If intPos01 > 0 Then
34600             strTmp03 = strTmp02
34610             Do While intPos01 > 0
34620               lngRels = lngRels + 1&
34630               lngE = lngRels - 1&
34640               ReDim Preserve arr_varRel(R_ELEMS, lngE)
34650               arr_varRel(R_ACCT, lngE) = Trim(Left(strTmp03, (intPos01 - 1)))
34660               arr_varRel(R_FND, lngE) = CBool(False)
34670               strTmp03 = Trim(Mid(strTmp03, (intPos01 + 1)))
34680               intPos01 = InStr(strTmp03, ",")
34690             Loop
34700             lngRels = lngRels + 1&
34710             lngE = lngRels - 1&
34720             ReDim Preserve arr_varRel(R_ELEMS, lngE)
34730             arr_varRel(R_ACCT, lngE) = strTmp03
34740             arr_varRel(R_FND, lngE) = CBool(False)
34750           Else
                  ' ** Only 1 account; it may just list the other.
34760             lngRels = lngRels + 1&
34770             lngE = lngRels - 1&
34780             ReDim Preserve arr_varRel(R_ELEMS, lngE)
34790             arr_varRel(R_ACCT, lngE) = strTmp02
34800             arr_varRel(R_FND, lngE) = CBool(False)
34810           End If
                ' ** Now check the accounts.
34820           For lngY = 0& To (lngRels - 1&)
34830             varTmp00 = Null
34840             varTmp00 = DLookup("[accountno]", "account", "[accountno] = '" & arr_varRel(R_ACCT, lngY) & "'")
34850             If IsNull(varTmp00) = False Then
34860               arr_varRel(R_FND, lngY) = True
34870             End If
34880             If lngY > 0& Then
                    ' ** Check for dupes.
34890               For lngZ = 0& To (lngY - 1&)
34900                 If arr_varRel(R_ACCT, lngZ) = arr_varRel(R_ACCT, lngY) Then
34910                   arr_varRel(R_DUPE, lngY) = True
34920                   Exit For
34930                 End If
34940               Next
34950             End If
34960             lngRelAlls = lngRelAlls + 1&
34970             lngE = lngRelAlls - 1&
34980             ReDim Preserve arr_varRelAll(RA_ELEMS, lngE)
34990             arr_varRelAll(RA_ACCT, lngE) = arr_varRel(R_ACCT, lngY)
35000             arr_varRelAll(RA_FND, lngE) = arr_varRel(R_FND, lngY)
35010             arr_varRelAll(RA_DUPE, lngE) = arr_varRel(R_DUPE, lngY)
35020           Next
35030           arr_varAcct(A_ARR, (lngAccts - 1&)) = arr_varRel
35040           arr_varAcct(A_RCNT, (lngAccts - 1&)) = lngRels
35050           If lngX < lngRecs Then .MoveNext
35060         Next
35070       End If
35080       .Close
35090     End With

35100     If lngAccts > 0& Then

            ' ** OK, we've collected all the related accounts, and their details.
            ' ** 1. Remove non-existant accounts from related_accountno fields.
            ' ** 2. Make sure each pair (or triple, etc.) have identical related_accountno field contents.
            ' ** 3. Remove any dupes found within a single related_accountno field.
            ' ** 4. Assemble proper account sets.
35110       For lngX = 0& To (lngAccts - 1&)
35120         strAcct = arr_varAcct(A_ACCT, lngX)
35130         arr_varTmp01 = arr_varAcct(A_ARR, lngX)
35140         If arr_varAcct(A_RCNT, lngX) = 1& Then
                ' ** It had better be a different account.
35150           If arr_varTmp01(R_ACCT, 0) = strAcct Then
                  ' ** Same account; delete it!
35160             arr_varAcct(A_NEW, lngX) = vbNullString
35170             arr_varAcct(A_RCNT, lngX) = 0&
35180           Else
35190             If arr_varTmp01(R_FND, 0) = False Then
                    ' ** It's a non-existant account; delete it!
35200               arr_varAcct(A_NEW, lngX) = vbNullString
35210               arr_varAcct(A_RCNT, lngX) = 0&
35220             Else
                    ' ** This should be the pair.
35230               arr_varAcct(A_NEW, lngX) = arr_varAcct(A_ACCT, lngX) & "," & arr_varTmp01(R_ACCT, 0)
                    ' ** Find it's counterpart.
35240               arr_varAcct(A_FND, lngX) = CBool(False)
35250               For lngY = 0& To (lngAccts - 1&)
35260                 If arr_varAcct(A_ACCT, lngY) <> arr_varAcct(A_ACCT, lngX) Then
35270                   arr_varAcct(A_FND, lngX) = True
                        ' ** OK, now how complicated do we want to make this?
                        ' ** If this counterpart also has related accounts, and its list is different,
                        ' ** should we combine them all, or just make it match the first one found?
                        ' ** Report this finding to the user, then give the first priority.
35280                   If arr_varAcct(A_NEW, lngY) <> vbNullString And arr_varAcct(A_NEW, lngY) <> arr_varAcct(A_NEW, lngX) Then
35290                     MsgBox "A routine check of the Related Accounts field has found some anomalies." & vbCrLf & vbCrLf & _
                            "Account: " & arr_varAcct(A_ACCT, lngY) & vbCrLf & _
                            "Previous Related Accounts: " & arr_varAcct(A_NEW, lngY) & vbCrLf & _
                            "New Related Accounts: " & arr_varAcct(A_NEW, lngX) & vbCrLf & vbCrLf & _
                            "Write these down, then check and rechoose Related Accounts as necessary.", _
                            vbExclamation + vbOKOnly, "Related Accounts Error"
35300                   End If
35310                   arr_varAcct(A_NEW, lngY) = arr_varAcct(A_NEW, lngX)
35320                   Exit For
35330                 End If
35340               Next
35350               If arr_varAcct(A_FND, lngX) = False Then
                      ' ** Only 1 of the pair was found to have a value in related_accountno; put the other one in.
35360                 lngAccts = lngAccts + 1&  ' ** Will this work OK while we're within its For/Next loop?
35370                 lngE = lngAccts - 1&
35380                 ReDim Preserve arr_varAcct(A_ELEMS, lngE)
35390                 arr_varAcct(A_ACCT, lngE) = arr_varTmp01(R_ACCT, 0)
35400                 arr_varAcct(A_RELS, lngE) = vbNullString  ' ** Yes, it's existing value is blank.
35410                 arr_varAcct(A_RCNT, lngE) = 2&
35420                 arr_varAcct(A_ARR, lngE) = Null
35430                 arr_varAcct(A_NEW, lngE) = arr_varAcct(A_NEW, lngX)
35440                 arr_varAcct(A_FND, lngE) = CBool(True)
35450                 arr_varAcct(A_RCNT, lngX) = 2&  ' ** Set this as now corrected.
35460                 arr_varAcct(A_FND, lngX) = CBool(True)  ' ** Set this as now found.
35470               End If
35480             End If
35490           End If
35500         Else
                ' ** More than 1 account is listed.
35510           lngRels = (UBound(arr_varTmp01, 2) + 1&)
35520           strTmp02 = vbNullString: lngTmp04 = 0&: blnTmp05 = False
35530           For lngY = 0& To (lngRels - 1&)
                  ' ** Assemble its proper list.
35540             If arr_varTmp01(R_FND, lngY) = True And arr_varTmp01(R_DUPE, lngY) = False Then
35550               lngTmp04 = lngTmp04 + 1&
35560               If lngTmp04 = 1& Then
35570                 strTmp02 = arr_varTmp01(R_ACCT, lngY)
                      ' ** If the array doesn't include this account, start with it.
35580                 For lngZ = 0& To (lngRels - 1&)
35590                   If arr_varTmp01(R_FND, lngZ) = True And arr_varTmp01(R_DUPE, lngZ) = False Then
35600                     If arr_varTmp01(R_ACCT, lngZ) = arr_varAcct(A_ACCT, lngX) Then
35610                       blnTmp05 = True
35620                       Exit For
35630                     End If
35640                   End If
35650                 Next  ' ** lngZ.
35660                 If blnTmp05 = False Then
35670                   lngTmp04 = lngTmp04 + 1&
35680                   strTmp02 = arr_varAcct(A_ACCT, lngX) & "," & strTmp02
35690                 End If
35700               Else
35710                 strTmp02 = strTmp02 & "," & arr_varTmp01(R_ACCT, lngY)
35720               End If
35730             End If
35740           Next  ' ** lngY.
35750           If lngTmp04 = 0& Then
                  ' ** No legitimate related accounts are listed; delete whatever's there.
35760             arr_varAcct(A_NEW, lngX) = vbNullString
35770             arr_varAcct(A_RCNT, lngX) = 0&
35780           Else
35790             If lngTmp04 = 1& Then  ' ** Remember, this is after checking validity and dupes.
35800               If strTmp02 = arr_varAcct(A_ACCT, lngX) Then
                      ' ** Only this one is listed, so delete it.
35810                 arr_varAcct(A_NEW, lngX) = vbNullString
35820                 arr_varAcct(A_RCNT, lngX) = 0&
35830               Else
                      ' ** This shouldn't happen! Only 1 account that isn't this, should have prompted this one to be added, above!
35840                 Beep
35850                 Debug.Print "'REL ACCTS? 1: " & arr_varAcct(A_ACCT, lngX) & " : " & strTmp02
35860               End If
35870             Else
                    ' ** More than 1, so should be OK.
35880               arr_varAcct(A_RCNT, lngX) = lngTmp04
35890               arr_varAcct(A_NEW, lngX) = strTmp02
                    ' ** Put this grouping in all the counterparts.
35900               For lngY = 0& To (lngRels - 1&)
35910                 If arr_varTmp01(R_FND, lngY) = True And arr_varTmp01(R_DUPE, lngY) = False Then
35920                   If arr_varTmp01(R_ACCT, lngY) <> arr_varAcct(A_ACCT, lngX) Then
35930                     blnTmp05 = False
35940                     For lngZ = 0& To (lngAccts - 1&)
35950                       If arr_varAcct(A_ACCT, lngZ) = arr_varTmp01(R_ACCT, lngY) Then
35960                         blnTmp05 = True
35970                         If arr_varAcct(A_NEW, lngZ) <> vbNullString Then
35980                           If arr_varAcct(A_NEW, lngZ) <> strTmp02 Then
                                  ' ** Oh boy! A royal mess!
35990                             If arr_varAcct(A_NEW, lngY) <> vbNullString And arr_varAcct(A_NEW, lngY) <> arr_varAcct(A_NEW, lngX) Then
36000                               MsgBox "A routine check of the Related Accounts field has found some anomalies." & vbCrLf & vbCrLf & _
                                      "Account: " & arr_varAcct(A_ACCT, lngZ) & vbCrLf & _
                                      "Previous Related Accounts: " & arr_varAcct(A_NEW, lngZ) & vbCrLf & _
                                      "New Related Accounts: " & arr_varAcct(A_NEW, lngX) & vbCrLf & vbCrLf & _
                                      "Write these down, then check and rechoose Related Accounts as necessary.", _
                                      vbExclamation + vbOKOnly, "Related Accounts Error"
36010                             End If
36020                           End If
36030                         Else
                                ' ** The loop has gotten here yet, so just put it in.
                                ' ** Any conflicts will be found when the loop does get here.
36040                         End If
36050                         arr_varAcct(A_RCNT, lngZ) = lngTmp04
36060                         arr_varAcct(A_NEW, lngZ) = strTmp02
36070                         Exit For
36080                       End If
36090                     Next
36100                     If blnTmp05 = False Then
                            ' ** At lease 1 of the group didn't have a value in related_accountno; put it into the master array.
36110                       lngAccts = lngAccts + 1&  ' ** Will this work OK while we're within its For/Next loop?
36120                       lngE = lngAccts - 1&
36130                       ReDim Preserve arr_varAcct(A_ELEMS, lngE)
36140                       arr_varAcct(A_ACCT, lngE) = arr_varTmp01(R_ACCT, lngY)
36150                       arr_varAcct(A_RELS, lngE) = vbNullString  ' ** Yes, it's existing value is blank.
36160                       arr_varAcct(A_RCNT, lngE) = lngTmp04
36170                       arr_varAcct(A_ARR, lngE) = Null
36180                       arr_varAcct(A_NEW, lngE) = strTmp02
36190                       arr_varAcct(A_FND, lngE) = CBool(True)
36200                     End If
36210                   End If
36220                 End If
36230               Next  ' ** lngY.
36240             End If
36250           End If
36260         End If
36270       Next  ' ** lngX.

            ' ** Yoikes! This has turned into a very complex procedure!

            ' ** Now update everybody for real.
36280       For lngX = 0& To (lngAccts - 1&)
36290         If arr_varAcct(A_RCNT, lngX) > 0& Then
36300           If arr_varAcct(A_NEW, lngX) <> vbNullString Then
                  ' ** Update Account for related accounts, by specified [actno], [relacct].
36310             If arr_varAcct(A_NEW, lngX) <> arr_varAcct(A_RELS, lngX) Then  ' ** Only update if different.
36320               Set qdf = .QueryDefs("qryAccountProfile_03")
36330               With qdf.Parameters
36340                 ![actno] = arr_varAcct(A_ACCT, lngX)
36350                 ![relacct] = arr_varAcct(A_NEW, lngX)
36360               End With
36370               qdf.Execute
36380             End If
36390           Else
                  ' ** What gives?!
36400             Beep
36410             Debug.Print "'REL ACCTS? 2: " & arr_varAcct(A_ACCT, lngX) & " : " & CStr(arr_varAcct(A_RCNT, lngX))
36420           End If
36430         Else
                ' ** Null out related_accountno in Account, by specified [actno].
36440           Set qdf = .QueryDefs("qryAccountProfile_04")
36450           With qdf.Parameters
36460             ![actno] = arr_varAcct(A_ACCT, lngX)
36470           End With
36480           qdf.Execute
36490         End If
36500       Next

36510     End If  ' ** lngAccts > 0&

36520     .Close
36530   End With

EXITP:
36540   Set rst = Nothing
36550   Set qdf = Nothing
36560   Set qdf = Nothing
36570   Exit Function

ERRH:
870     blnRetVal = False
880     Select Case ERR.Number
        Case Else
890       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
900     End Select
910     Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

36610   Select Case intWhere
        Case acCmdRecordsGoToFirst
36620     DoCmd.RunCommand acCmdRecordsGoToFirst
36630   Case acCmdRecordsGoToLast
36640     DoCmd.RunCommand acCmdRecordsGoToLast
36650   Case acCmdRecordsGoToPrevious
36660     DoCmd.RunCommand acCmdRecordsGoToPrevious
36670   Case acCmdRecordsGoToNext
36680     DoCmd.RunCommand acCmdRecordsGoToNext
36690   Case Else
36700     If IsMissing(varID) = False Then
36710       With Me
36720         Set rst = .RecordsetClone
36730         With rst
36740           .FindFirst "[accountno] = '" & varID & "'"
36750           If .NoMatch = False Then
36760             Me.Bookmark = .Bookmark
36770           End If
36780           .Close
36790         End With
36800       End With
36810     End If
36820   End Select

EXITP:
36830   Set rst = Nothing
36840   Exit Sub

ERRH:
870     Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
880     Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
890     Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub
