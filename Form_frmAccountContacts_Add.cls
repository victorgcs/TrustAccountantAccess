VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountContacts_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountContacts_Add"

'VGC 07/23/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Cancel:           C {cmdCancel}
' **   Email:            E {Contact_Email}
' **   Fax:              F {Contact_Fax}
' **   Other:            H {opgLocContact_optOther}
' **   Name:             M {opgAccountSource_optName}
' **   Contact Name:     N {Contact_Name}
' **   OK:               O {cmdOK}
' **   Phone 1:          P {Contact_Phone1}
' **   USA:              U {opgLocContact_optUSA}
' **   Contact 1:        1 {opgConNum_optOne}
' **   Contact 2:        2 {opgConNum_optTwo}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** cmbAccounts combo box constants:
Private Const CBX_A_ACTNO As Integer = 0  ' ** accountno
'Private Const CBX_A_DESC  As Integer = 1  ' ** Desc
Private Const CBX_A_SNAM  As Integer = 2  ' ** shortname
'Private Const CBX_A_LNAM  As Integer = 3  ' ** legalname
Private Const CBX_A_ONE   As Integer = 4  ' ** Has_One
Private Const CBX_A_TWO   As Integer = 5  ' ** Has_Two

Private frm As Access.Form
Private lngOptOneLbl_Top As Long, lngOptOneLbl_Left As Long, lngOptTwoLbl_Top As Long, lngOptTwoLbl_Left As Long
Private blnIsOpen As Boolean, blnEmpty As Boolean, blnHasOne As Boolean, blnHasTwo As Boolean
Private strAccountNo As String, lngTpp As Long, blnCountry As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = Trim(.OpenArgs)
150           intPos01 = InStr(strTmp01, "~")
160           If intPos01 > 0 Then

170             blnIsOpen = True
180             blnHasOne = False: blnHasTwo = False
190             blnEmpty = False: blnCountry = False

200             strCallingForm = Left(strTmp01, (intPos01 - 1))
210             strTmp01 = Mid(strTmp01, (intPos01 + 1))
220             intPos01 = InStr(strTmp01, "~")
230             strAccountNo = Left(strTmp01, (intPos01 - 1))
240             strTmp01 = Mid(strTmp01, (intPos01 + 1))
250             intPos01 = InStr(strTmp01, "~")
260             If Left(strTmp01, (intPos01 - 1)) = "True" Then blnHasOne = True
270             If Mid(strTmp01, (intPos01 + 1)) = "True" Then blnHasTwo = True

280             lngOptOneLbl_Top = .opgConNum_optOne_lbl.Top
290             lngOptOneLbl_Left = .opgConNum_optOne_lbl.Left
300             lngOptTwoLbl_Top = .opgConNum_optTwo_lbl.Top
310             lngOptTwoLbl_Left = .opgConNum_optTwo_lbl.Left
320             .Contact_Msg_lbl1.Visible = False
330             .Contact_Msg_lbl2.Visible = False

340             Set frm = Forms(strCallingForm)

                ' ** tblAccount_Contact_Staging, linked to Account.
350             varTmp00 = DCount("*", "qryAccount_Contacts_01_01")
360             If IsNull(varTmp00) = True Then
370               blnEmpty = True
380             Else
390               If varTmp00 = 0 Then
400                 blnEmpty = True
410               Else
420                 blnEmpty = False
430               End If
440             End If

450             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

460             StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

470             CLR_DISABLED_FG = CLR_DKGRY
480             CLR_DISABLED_BG = CLR_LTTEAL

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
490             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

500             DoCmd.RunCommand acCmdRecordsGoToNew

510             opgAccountSource_AfterUpdate  ' ** Procedure: Below.
520             chkRememberMe_AfterUpdate  ' ** Procedure: Below.
530             opgConNum_AfterUpdate  ' ** Procedure: Below.
540             opgLocContact_AfterUpdate  ' ** Procedure: Below.

550             Select Case frm.opgView
                Case frm.opgView_optWide.OptionValue
560               blnCountry = frm.chkEnableCountry_Wide
570             Case frm.opgView_optCompact.OptionValue
580               blnCountry = frm.chkEnableCountry_Compact
590             End Select
600             If blnCountry = False Then
610               If .opgLocContact = .opgLocContact_optOther.OptionValue Then
620                 .opgLocContact = .opgLocContact_optUSA.OptionValue
630                 opgLocContact_AfterUpdate  ' ** Procedure: Below.
640               End If
650               .opgLocContact.Enabled = False
660             End If

                ' ** EVENT CHECK: chkRememberMe!
670             cmbAccounts_AfterUpdate  ' ** Procedure: Below.

680             If strAccountNo = "{none}" Then
690               .cmbAccounts = Null
700               .cmbAccounts.SetFocus
710             Else
720               If blnHasOne = True Or blnHasTwo = True Then
730                 If blnHasOne = True And blnHasTwo = True Then
                      ' ** They'll have to choose another account.
740                   .cmbAccounts = Null
750                   .Contact_Msg_lbl1.Caption = "Account # " & strAccountNo
760                   .Contact_Msg_lbl1.Visible = True
770                   .Contact_Msg_lbl2.Visible = True
780                   DoEvents
790                   DisableAll True  ' ** Procedure: Below.
800                   .cmbAccounts.SetFocus
810                 Else
                      ' ** Only allow the missing contact.
820                   Select Case blnHasOne
                      Case True
830                     .opgConNum = .opgConNum_optTwo.OptionValue
840                     opgConNum_AfterUpdate  ' ** Procedure: Below.
850                     .opgConNum_optOne.Enabled = False
860                     .opgConNum_optOne_lbl.Top = (.opgConNum_optOne.Top - (3& * lngTpp))
870                     .opgConNum_optOne_lbl.Left = (.opgConNum_optOne.Left + (13& * lngTpp))
880                     .opgConNum_optOne_lbl_dim_hi.Visible = True
890                   Case False
900                     .opgConNum = .opgConNum_optOne.OptionValue
910                     opgConNum_AfterUpdate  ' ** Procedure: Below.
920                     .opgConNum_optTwo.Enabled = False
930                     .opgConNum_optTwo_lbl.Top = (.opgConNum_optTwo.Top - (3& * lngTpp))
940                     .opgConNum_optTwo_lbl.Left = (.opgConNum_optTwo.Left + (13& * lngTpp))
950                     .opgConNum_optTwo_lbl_dim_hi.Visible = True
960                   End Select
970                   .cmbAccounts = strAccountNo  ' ** Don't put it into accountno yet.
980                   .Contact_Name.SetFocus
990                 End If
1000              End If
1010            End If

1020          Else
1030            Cancel = -1
1040          End If
1050        Else
1060          Cancel = -1
1070        End If
1080      Else
1090        Cancel = -1
1100      End If
1110    End With

1120    If Cancel = -1 Then
1130      Beep
1140      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1150    End If

1160    DoCmd.Hourglass False

EXITP:
1170    Exit Sub

ERRH:
1180    DoCmd.Hourglass False
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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Other:            H {opgLocContact_optOther}
        ' **   USA:              U {opgLocContact_optUSA}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyEscape
1370        With Me
1380          intRetVal = 0
1390          cmdCancel_Click  ' ** Procedure: Below.
1400        End With
1410      End Select
1420    End If

        ' ** Alt keys.
1430    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1440      Select Case intRetVal
          Case vbKeyH
1450        With Me
1460          intRetVal = 0
1470          If .opgLocContact.Enabled = True Then
1480            .opgLocContact.SetFocus
1490            .opgLocContact = .opgLocContact_optOther.OptionValue
1500            opgLocContact_AfterUpdate  ' ** Procedure: Below
1510          End If
1520        End With
1530      Case vbKeyU
1540        With Me
1550          intRetVal = 0
1560          If .opgLocContact.Enabled = True Then
1570            .opgLocContact.SetFocus
1580            .opgLocContact = .opgLocContact_optUSA.OptionValue
1590            opgLocContact_AfterUpdate  ' ** Procedure: Below
1600          End If
1610        End With
1620      End Select
1630    End If

        ' ** Ctrl keys.
1640    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1650      Select Case intRetVal
          Case vbKeyS
1660        intRetVal = 0
1670        cmdSave_Click  ' ** Procedure: Below.
1680      End Select
1690    End If

        ' ** Ctrl-Shift keys.
1700    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1710      Select Case intRetVal
          Case vbKeyF
1720        With Me
1730          intRetVal = 0
1740          .FocusHolder.SetFocus
1750        End With
1760      End Select
1770    End If

EXITP:
1780    KeyCode = intRetVal
1790    Exit Sub

ERRH:
1800    intRetVal = 0
1810    Select Case ERR.Number
        Case Else
1820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1830    End Select
1840    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1910    With Me
1920      If .chkRememberMe = False Then
1930        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
1940      End If
1950    End With

1960    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
1970    Set frm = Nothing
1980    Exit Sub

ERRH:
1990    Select Case ERR.Number
        Case Else
2000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2010    End Select
2020    Resume EXITP

End Sub

Private Sub cmdOK_Click()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim frmSub1 As Access.Form, frmSub2 As Access.Form
        Dim lngContactID As Long
        Dim blnContinue As Boolean

2110    blnContinue = True

2120    With Me
2130      If IsNull(.cmbAccounts) = True Then
2140        blnContinue = False
2150        MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
2160      Else
2170        If Trim(.cmbAccounts) = vbNullString Then
2180          blnContinue = False
2190          MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
2200        Else
2210          If IsNull(.Contact_Name) = True Then
2220            blnContinue = False
2230            MsgBox "You must enter a Contact Name to continue.", vbInformation + vbOKOnly, "Entry Required"
2240          Else
2250            If Trim(.Contact_Name) = vbNullString Then
2260              blnContinue = False
2270              MsgBox "You must enter a Contact Name to continue.", vbInformation + vbOKOnly, "Entry Required"
2280            Else

2290              If blnHasOne = False And blnHasTwo = False Then
                    ' ** Now they have one.
2300                .Contact_ConNum1 = True
2310                .Contact_ConNum2 = False
2320              ElseIf blnHasOne = True Or blnHasTwo = True Then  ' ** It wouldn't get here if both were True.
                    ' ** Now they have two.
2330                .Contact_ConNum1 = True
2340                .Contact_ConNum2 = True
2350              End If

2360              .accountno = .cmbAccounts
2370              .Contact_Changed = True
2380              .Contact_Number = .opgConNum
2390              .Username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
2400              .Contact_DateModified = Now()

2410              cmdSave_Click  ' ** Procedure: Above.
2420              lngContactID = .Contact_ID

2430              If blnEmpty = True Then
                    ' ** Turn everything on!
2440                With frm
2450                  Set frmSub1 = .frmAccountContacts_Sub.Form
2460                  Set frmSub2 = .frmAccountContacts_Sub2.Form
2470                  .cmdPreview.Enabled = True
2480                  .cmdPrint.Enabled = True
2490                  .cmdWord.Enabled = True
      #If NoExcel Then
2500                  .cmdExcel.Enabled = False
      #Else
2510                  .cmdExcel.Enabled = True
      #End If
2520                  .cmdDelete.Enabled = True
2530                  .cmdUpdate.Enabled = True
2540                  .opgFormat_lbl.ForeColor = CLR_VDKGRY
2550                  .opgFormat_lbl_dim_hi.Visible = False
2560                  .opgFormatName.Enabled = True
2570                  .opgFormatZip.Enabled = True
2580                  .opgFormatPhone.Enabled = True
2590                  .opgPrint_lbl.ForeColor = CLR_VDKGRY
2600                  .opgPrint_lbl_dim_hi.Visible = False
2610                  .opgPrintAddress.Enabled = True
2620                  .opgPrintCSZ.Enabled = True
2630                  .opgPrintCSZCP.Enabled = True
2640                  .opgPrintPhone.Enabled = True
2650                  .opgPrintFax.Enabled = True
2660                  .opgPrintEmail.Enabled = True
2670                  If frm.opgView = frm.opgView_optWide.OptionValue Then
2680                    .chkShowAcctNum.Enabled = True
2690                    .chkShowShortName.Enabled = True
2700                  End If
2710                  .chkEnableCountry_Wide.Enabled = True
2720                  .chkEnableCountry_Compact.Enabled = True
2730                  .chkPageOf.Enabled = True
2740                  .chkAlwaysUpdate.Enabled = True
2750                  .chkAlwaysUpdate_lbl2.ForeColor = CLR_VDKGRY
2760                  .chkAlwaysUpdate_lbl2_dim_hi.Visible = False
2770                  With frmSub1
                        ' ** Whether or not a particular column is visible should make no difference here.
2780                    .accountno_lbl.ForeColor = CLR_DKGRY2
2790                    .accountno_lbl_dim_hi.Visible = False
2800                    .accountno_lbl_line.BorderColor = CLR_DKGRY
2810                    .accountno_lbl_line_dim_hi.Visible = False
2820                    .shortname_lbl.ForeColor = CLR_DKGRY2
2830                    .shortname_lbl_dim_hi.Visible = False
2840                    .shortname_lbl_line.BorderColor = CLR_DKGRY
2850                    .shortname_lbl_Line_dim_hi.Visible = False
2860                    .Contact_Name_lbl.ForeColor = CLR_DKGRY2
2870                    .Contact_Name_lbl_dim_hi.Visible = False
2880                    .Contact_Name_lbl_line.BorderColor = CLR_DKGRY
2890                    .Contact_Name_lbl_line_dim_hi.Visible = False
2900                    .Contact_Address1_lbl.ForeColor = CLR_DKGRY2
2910                    .Contact_Address1_lbl_dim_hi.Visible = False
2920                    .Contact_Address1_lbl_line.BorderColor = CLR_DKGRY
2930                    .Contact_Address1_lbl_line_dim_hi.Visible = False
2940                    .Contact_Address2_lbl.ForeColor = CLR_DKGRY2
2950                    .Contact_Address2_lbl_dim_hi.Visible = False
2960                    .Contact_Address2_lbl_line.BorderColor = CLR_DKGRY
2970                    .Contact_Address2_lbl_line_dim_hi.Visible = False
2980                    .Contact_City_lbl.ForeColor = CLR_DKGRY2
2990                    .Contact_City_lbl_dim_hi.Visible = False
3000                    .Contact_City_lbl_line.BorderColor = CLR_DKGRY
3010                    .Contact_City_lbl_line_dim_hi.Visible = False
3020                    .Contact_State_lbl.ForeColor = CLR_DKGRY2
3030                    .Contact_State_lbl_dim_hi.Visible = False
3040                    .Contact_State_lbl_line.BorderColor = CLR_DKGRY
3050                    .Contact_State_lbl_line_dim_hi.Visible = False
3060                    .Contact_Zip_lbl.ForeColor = CLR_DKGRY2
3070                    .Contact_Zip_lbl_dim_hi.Visible = False
3080                    .Contact_Zip_lbl_line.BorderColor = CLR_DKGRY
3090                    .Contact_Zip_lbl_line_dim_hi.Visible = False
3100                    .Contact_Country_lbl.ForeColor = CLR_DKGRY2
3110                    .Contact_Country_lbl_dim_hi.Visible = False
3120                    .Contact_Country_lbl_line.BorderColor = CLR_DKGRY
3130                    .Contact_Country_lbl_line_dim_hi.Visible = False
3140                    .Contact_PostalCode_lbl.ForeColor = CLR_DKGRY2
3150                    .Contact_PostalCode_lbl_dim_hi.Visible = False
3160                    .Contact_PostalCode_lbl_line.BorderColor = CLR_DKGRY
3170                    .Contact_PostalCode_lbl_line_dim_hi.Visible = False
3180                    .Contact_Phone1_lbl.ForeColor = CLR_DKGRY2
3190                    .Contact_Phone1_lbl_dim_hi.Visible = False
3200                    .Contact_Phone1_lbl_line.BorderColor = CLR_DKGRY
3210                    .Contact_Phone1_lbl_line_dim_hi.Visible = False
3220                    .Contact_Phone2_lbl.ForeColor = CLR_DKGRY2
3230                    .Contact_Phone2_lbl_dim_hi.Visible = False
3240                    .Contact_Phone2_lbl_line.BorderColor = CLR_DKGRY
3250                    .Contact_Phone2_lbl_line_dim_hi.Visible = False
3260                    .Contact_Fax_lbl.ForeColor = CLR_DKGRY2
3270                    .Contact_Fax_lbl_dim_hi.Visible = False
3280                    .Contact_Fax_lbl_line.BorderColor = CLR_DKGRY
3290                    .Contact_Fax_lbl_line_dim_hi.Visible = False
3300                    .Contact_Email_lbl.ForeColor = CLR_DKGRY2
3310                    .Contact_Email_lbl_dim_hi.Visible = False
3320                    .Contact_Email_lbl_line.BorderColor = CLR_DKGRY
3330                    .Contact_Email_lbl_line_dim_hi.Visible = False
3340                  End With  ' ** frmSub1.
3350                  With frmSub2
                        ' ** Whether or not a particular column is visible should make no difference here.
3360                    .accountno1_lbl.ForeColor = CLR_DKGRY2
3370                    .accountno1_lbl_dim_hi.Visible = False
3380                    .accountno1_lbl_line.BorderColor = CLR_DKGRY
3390                    .accountno1_lbl_line_dim_hi.Visible = False
3400                    .Contact_Number_lbl.ForeColor = CLR_DKGRY2
3410                    .Contact_Number_lbl_dim_hi.Visible = False
3420                    .Contact_Number_lbl_line.BorderColor = CLR_DKGRY
3430                    .Contact_Number_lbl_line_dim_hi.Visible = False
3440                    .Contact_Name_lbl.ForeColor = CLR_DKGRY2
3450                    .Contact_Name_lbl_dim_hi.Visible = False
3460                    .Contact_Name_lbl_line.BorderColor = CLR_DKGRY
3470                    .Contact_Name_lbl_line_dim_hi.Visible = False
3480                    .Contact_Address1_lbl.ForeColor = CLR_DKGRY2
3490                    .Contact_Address1_lbl_dim_hi.Visible = False
3500                    .Contact_Address1_lbl_line.BorderColor = CLR_DKGRY
3510                    .Contact_Address1_lbl_line_dim_hi.Visible = False
3520                    .opgLocContact_lbl.ForeColor = CLR_DKGRY2
3530                    .opgLocContact_lbl_dim_hi.Visible = False
3540                    .opgLocContact_lbl2.ForeColor = CLR_DKGRY2
3550                    .opgLocContact_lbl2_dim_hi.Visible = False
3560                    .opgLocContact_lbl_line.BorderColor = CLR_DKGRY
3570                    .opgLocContact_lbl_line_dim_hi.Visible = False
3580                    .Contact_City_lbl.ForeColor = CLR_DKGRY2
3590                    .Contact_City_lbl_dim_hi.Visible = False
3600                    .Contact_City_lbl_line.BorderColor = CLR_DKGRY
3610                    .Contact_City_lbl_line_dim_hi.Visible = False
3620                    .Contact_Country_lbl.ForeColor = CLR_DKGRY2
3630                    .Contact_Country_lbl_dim_hi.Visible = False
3640                    .Contact_Country_lbl_line.BorderColor = CLR_DKGRY
3650                    .Contact_Country_lbl_line_dim_hi.Visible = False
3660                    .Contact_Phone1_lbl.ForeColor = CLR_DKGRY2
3670                    .Contact_Phone1_lbl_dim_hi.Visible = False
3680                    .Contact_Phone1_lbl_line.BorderColor = CLR_DKGRY
3690                    .Contact_Phone1_lbl_line_dim_hi.Visible = False
3700                    .Contact_Email_lbl.ForeColor = CLR_DKGRY2
3710                    .Contact_Email_lbl_dim_hi.Visible = False
3720                    .Contact_Email_lbl_line.BorderColor = CLR_DKGRY
3730                    .Contact_Email_lbl_line_dim_hi.Visible = False
3740                  End With  ' ** frmSub2.
3750                End With  ' ** frm.
3760              End If  ' ** blnEmpty.

3770              frmSub1.Requery
3780              frmSub1.MoveRec 0, lngContactID  ' ** Form Procedure: frmAccountContacts_Sub.
3790              frmSub2.Requery
3800              frmSub2.MoveRec 0, lngContactID  ' ** Form Procedure: frmAccountContacts_Sub.

3810              DoCmd.Close acForm, THIS_NAME

                  'Select Case frm.opgView
                  'Case frm.opgView_optWide.OptionValue
                  'Case frm.opgView_optCompact.OptionValue
                  'End Select

3820            End If
3830          End If
3840        End If
3850      End If
3860    End With

EXITP:
3870    Set frmSub1 = Nothing
3880    Set frmSub2 = Nothing
3890    Exit Sub

ERRH:
3900    Select Case ERR.Number
        Case Else
3910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3920    End Select
3930    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4020    intShiftDown = (Shift And acShiftMask) > 0
4030    intAltDown = (Shift And acAltMask) > 0
4040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4060      Select Case intRetVal
          Case vbKeyTab
4070        With Me
4080          intRetVal = 0
4090          .cmdCancel.SetFocus
4100        End With
4110      End Select
4120    End If

        ' ** Shift keys.
4130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4140      Select Case intRetVal
          Case vbKeyTab
4150        With Me
4160          intRetVal = 0
4170          If .Contact_Email.Enabled = True Then
4180            .Contact_Email.SetFocus
4190          Else
4200            .cmbAccounts.SetFocus
4210          End If
4220        End With
4230      End Select
4240    End If

EXITP:
4250    KeyCode = intRetVal
4260    Exit Sub

ERRH:
4270    intRetVal = 0
4280    Select Case ERR.Number
        Case Else
4290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4300    End Select
4310    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

4410    With Me
4420      Select Case IsNull(.Contact_ID)
          Case True
            ' ** Just close, no entry yet made.
4430      Case False
4440        DoCmd.RunCommand acCmdDeleteRecord
4450      End Select

4460      DoCmd.Close acForm, THIS_NAME

4470    End With

EXITP:
4480    Exit Sub

ERRH:
4490    Select Case ERR.Number
        Case Else
4500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4510    End Select
4520    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4620    intShiftDown = (Shift And acShiftMask) > 0
4630    intAltDown = (Shift And acAltMask) > 0
4640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab
4670        With Me
4680          intRetVal = 0
4690          .cmbAccounts.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab
4750        With Me
4760          intRetVal = 0
4770          If .cmdOK.Enabled = True Then
4780            .cmdOK.SetFocus
4790          Else
4800            .cmbAccounts.SetFocus
4810          End If
4820        End With
4830      End Select
4840    End If

        ' ** Ctrl keys.
4850    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4860      Select Case intRetVal
          Case vbKeyTab
4870        With Me
4880          intRetVal = 0
4890          .opgAccountSource.SetFocus
4900        End With
4910      End Select
4920    End If

EXITP:
4930    KeyCode = intRetVal
4940    Exit Sub

ERRH:
4950    intRetVal = 0
4960    Select Case ERR.Number
        Case Else
4970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4980    End Select
4990    Resume EXITP

End Sub

Public Sub cmdSave_Click()

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

Private Sub opgAccountSource_AfterUpdate()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

5110    strAccountNo = vbNullString

5120    With Me
5130      If IsNull(.cmbAccounts) = False Then
5140        If Len(.cmbAccounts.Column(CBX_A_ACTNO)) > 0 Then
5150          strAccountNo = .cmbAccounts.Column(CBX_A_ACTNO)
5160        End If
5170      End If
5180      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
5190        .cmbAccounts.RowSource = "qryAccount_Contacts_18a"
5200        .opgAccountSource_optNumber_lbl.FontBold = True
5210        .opgAccountSource_optName_lbl.FontBold = False
5220      Case .opgAccountSource_optName.OptionValue
5230        .cmbAccounts.RowSource = "qryAccount_Contacts_18b"
5240        .opgAccountSource_optNumber_lbl.FontBold = False
5250        .opgAccountSource_optName_lbl.FontBold = True
5260      End Select
5270      DoEvents
5280      If strAccountNo <> vbNullString Then
5290        .cmbAccounts = strAccountNo
5300      End If
5310    End With

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

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
5490          .chkRememberMe.SetFocus
5500        End With
5510      End Select
5520    End If

        ' ** Shift keys.
5530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5550        With Me
5560          intRetVal = 0
5570          .cmdCancel.SetFocus
5580        End With
5590      End Select
5600    End If

EXITP:
5610    KeyCode = intRetVal
5620    Exit Sub

ERRH:
5630    intRetVal = 0
5640    Select Case ERR.Number
        Case Else
5650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5660    End Select
5670    Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5720    intShiftDown = (Shift And acShiftMask) > 0
5730    intAltDown = (Shift And acAltMask) > 0
5740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5770        With Me
5780          intRetVal = 0
5790          .chkRememberMe.SetFocus
5800        End With
5810      End Select
5820    End If

        ' ** Shift keys.
5830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5850        With Me
5860          intRetVal = 0
5870          .cmdCancel.SetFocus
5880        End With
5890      End Select
5900    End If

EXITP:
5910    KeyCode = intRetVal
5920    Exit Sub

ERRH:
5930    intRetVal = 0
5940    Select Case ERR.Number
        Case Else
5950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5960    End Select
5970    Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

6010    With Me
6020      Select Case .chkRememberMe
          Case True
6030        .chkRememberMe_lbl.FontBold = True
6040      Case False
6050        .chkRememberMe_lbl.FontBold = False
6060      End Select
6070    End With

EXITP:
6080    Exit Sub

ERRH:
6090    Select Case ERR.Number
        Case Else
6100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6110    End Select
6120    Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

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
6290          .cmbAccounts.SetFocus
6300        End With
6310      End Select
6320    End If

        ' ** Shift keys.
6330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6350        With Me
6360          intRetVal = 0
6370          .opgAccountSource.SetFocus
6380        End With
6390      End Select
6400    End If

EXITP:
6410    KeyCode = intRetVal
6420    Exit Sub

ERRH:
6430    intRetVal = 0
6440    Select Case ERR.Number
        Case Else
6450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6460    End Select
6470    Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

6510    With Me
6520      If .cmbAccounts.text = vbNullString Then
6530        ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
6540      End If
6550    End With

EXITP:
6560    Exit Sub

ERRH:
6570    Select Case ERR.Number
        Case Else
6580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6590    End Select
6600    Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

6710    With Me
6720      If IsNull(.cmbAccounts) = False Then
6730        If Trim(.cmbAccounts) <> vbNullString Then
6740          If IsNull(.cmbAccounts.Column(CBX_A_SNAM)) = False Then  ' ** Shortname.
6750            blnHasOne = CBool(.cmbAccounts.Column(CBX_A_ONE))
6760            blnHasTwo = CBool(.cmbAccounts.Column(CBX_A_TWO))
6770            If blnHasOne = True And blnHasTwo = True Then
6780              .Contact_Msg_lbl1.Caption = "Account # " & .cmbAccounts.Column(CBX_A_ACTNO)
6790              .Contact_Msg_lbl1.Visible = True
6800              .Contact_Msg_lbl2.Visible = True
6810              DisableAll True  ' ** Procedure: Below.
6820            Else
6830              .Contact_Msg_lbl1.Visible = False
6840              .Contact_Msg_lbl2.Visible = False
6850              DisableAll False  ' ** Procedure: Below.
6860              DoEvents
6870              If blnHasOne = False And blnHasTwo = False Then
6880                .opgConNum = .opgConNum_optOne.OptionValue
6890                opgConNum_AfterUpdate  ' ** Procedure: Below.
6900              Else
6910                Select Case blnHasOne
                    Case True
6920                  .opgConNum = .opgConNum_optTwo.OptionValue
6930                  opgConNum_AfterUpdate  ' ** Procedure: Below.
6940                  DoEvents
6950                  .opgConNum_optOne.Enabled = False
6960                  .opgConNum_optOne_lbl.Top = (.opgConNum_optOne.Top - (3& * lngTpp))
6970                  .opgConNum_optOne_lbl.Left = (.opgConNum_optOne.Left + (13& * lngTpp))
6980                  .opgConNum_optOne_lbl_dim_hi.Visible = True
6990                Case False
7000                  .opgConNum = .opgConNum_optOne.OptionValue
7010                  opgConNum_AfterUpdate  ' ** Procedure: Below.
7020                  DoEvents
7030                  .opgConNum_optTwo.Enabled = False
7040                  .opgConNum_optTwo_lbl.Top = (.opgConNum_optTwo.Top - (3& * lngTpp))
7050                  .opgConNum_optTwo_lbl.Left = (.opgConNum_optTwo.Left + (13& * lngTpp))
7060                  .opgConNum_optTwo_lbl_dim_hi.Visible = True
7070                End Select
7080              End If
7090              If IsNull(.Contact_ID) = False Then
7100  On Error Resume Next
7110                DoCmd.RunCommand acCmdDeleteRecord
7120                DoCmd.RunCommand acCmdRecordsGoToNew
7130  On Error GoTo ERRH
7140              End If
7150            End If
7160          End If
7170        End If
7180      End If
7190    End With

EXITP:
7200    Exit Sub

ERRH:
7210    Select Case ERR.Number
        Case Else
7220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7230    End Select
7240    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

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
7390          If .Contact_Name.Enabled = True Then
7400            .Contact_Name.SetFocus
7410          Else
7420            .cmdCancel.SetFocus
7430          End If
7440        End With
7450      End Select
7460    End If

        ' ** Shift keys.
7470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7480      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7490        With Me
7500          intRetVal = 0
7510          .cmdCancel.SetFocus
7520        End With
7530      End Select
7540    End If

        ' ** Ctrl-Shift keys.
7550    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7570        With Me
7580          intRetVal = 0
7590          .chkRememberMe.SetFocus
7600        End With
7610      End Select
7620    End If

EXITP:
7630    KeyCode = intRetVal
7640    Exit Sub

ERRH:
7650    intRetVal = 0
7660    Select Case ERR.Number
        Case Else
7670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7680    End Select
7690    Resume EXITP

End Sub

Private Sub opgConNum_AfterUpdate()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgConNum_AfterUpdate"

7710    With Me
7720      Select Case .opgConNum
          Case .opgConNum_optOne.OptionValue
7730        .opgConNum_optOne_lbl.FontBold = True
7740        .opgConNum_optTwo_lbl.FontBold = False
7750        .opgConNum_optOne_box.Visible = True
7760        .opgConNum_optTwo_box.Visible = False
7770      Case .opgConNum_optTwo.OptionValue
7780        .opgConNum_optOne_lbl.FontBold = False
7790        .opgConNum_optTwo_lbl.FontBold = True
7800        .opgConNum_optOne_box.Visible = False
7810        .opgConNum_optTwo_box.Visible = True
7820      Case Else
7830        .opgConNum_optOne_lbl.FontBold = False
7840        .opgConNum_optTwo_lbl.FontBold = False
7850        .opgConNum_optOne_box.Visible = False
7860        .opgConNum_optTwo_box.Visible = False
7870      End Select
7880    End With

EXITP:
7890    Exit Sub

ERRH:
7900    Select Case ERR.Number
        Case Else
7910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7920    End Select
7930    Resume EXITP

End Sub

Private Sub opgConNum_optOne_KeyDown(KeyCode As Integer, Shift As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgConNum_optOne_KeyDown"

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
8090          .Contact_Name.SetFocus
8100        End With
8110      End Select
8120    End If

        ' ** Shift keys.
8130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8150        With Me
8160          intRetVal = 0
8170          .cmbAccounts.SetFocus
8180        End With
8190      End Select
8200    End If

        ' ** Ctrl-Shift keys.
8210    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8220      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8230        With Me
8240          intRetVal = 0
8250          .chkRememberMe.SetFocus
8260        End With
8270      End Select
8280    End If

EXITP:
8290    KeyCode = intRetVal
8300    Exit Sub

ERRH:
8310    intRetVal = 0
8320    Select Case ERR.Number
        Case Else
8330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8340    End Select
8350    Resume EXITP

End Sub

Private Sub opgConNum_optTwo_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgConNum_optTwo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8420    intShiftDown = (Shift And acShiftMask) > 0
8430    intAltDown = (Shift And acAltMask) > 0
8440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8470        With Me
8480          intRetVal = 0
8490          .Contact_Name.SetFocus
8500        End With
8510      End Select
8520    End If

        ' ** Shift keys.
8530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8550        With Me
8560          intRetVal = 0
8570          .cmbAccounts.SetFocus
8580        End With
8590      End Select
8600    End If

        ' ** Ctrl-Shift keys.
8610    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8620      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8630        With Me
8640          intRetVal = 0
8650          .chkRememberMe.SetFocus
8660        End With
8670      End Select
8680    End If

EXITP:
8690    KeyCode = intRetVal
8700    Exit Sub

ERRH:
8710    intRetVal = 0
8720    Select Case ERR.Number
        Case Else
8730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8740    End Select
8750    Resume EXITP

End Sub

Private Sub Contact_Name_AfterUpdate()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_AfterUpdate"

        Dim strTmp01 As String

8810    With Me
8820      If IsNull(.Contact_Name) = False Then
8830        strTmp01 = FixQuotes(.Contact_Name)  ' ** Module Function: modStringFuncs.
8840        If Compare_StringA_StringB(.Contact_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
8850          .Contact_Name = strTmp01
8860          cmdSave_Click ' ** Procedure: Above.
8870        End If
8880      End If
8890    End With

EXITP:
8900    Exit Sub

ERRH:
8910    Select Case ERR.Number
        Case Else
8920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8930    End Select
8940    Resume EXITP

End Sub

Private Sub Contact_Address1_AfterUpdate()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_AfterUpdate"

        Dim strTmp01 As String

9010    With Me
9020      If IsNull(.Contact_Address1) = False Then
9030        strTmp01 = FixQuotes(.Contact_Address1)  ' ** Module Function: modStringFuncs.
9040        If Compare_StringA_StringB(.Contact_Address1, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
9050          .Contact_Address1 = strTmp01
9060          cmdSave_Click  ' ** Procedure: Above.
9070        End If
9080      End If
9090    End With

EXITP:
9100    Exit Sub

ERRH:
9110    Select Case ERR.Number
        Case Else
9120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9130    End Select
9140    Resume EXITP

End Sub

Private Sub Contact_Address2_AfterUpdate()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_AfterUpdate"

        Dim strTmp01 As String

9210    With Me
9220      If IsNull(.Contact_Address2) = False Then
9230        strTmp01 = FixQuotes(.Contact_Address2)  ' ** Module Function: modStringFuncs.
9240        If Compare_StringA_StringB(.Contact_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
9250          .Contact_Address2 = strTmp01
9260          cmdSave_Click  ' ** Procedure: Above.
9270        End If
9280      End If
9290    End With

EXITP:
9300    Exit Sub

ERRH:
9310    Select Case ERR.Number
        Case Else
9320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9330    End Select
9340    Resume EXITP

End Sub

Private Sub Contact_City_AfterUpdate()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_AfterUpdate"

        Dim strTmp01 As String

9410    With Me
9420      If IsNull(.Contact_City) = False Then
9430        strTmp01 = FixQuotes(.Contact_City)  ' ** Module Function: modStringFuncs.
9440        If Compare_StringA_StringB(.Contact_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
9450          .Contact_City = strTmp01
9460          cmdSave_Click  ' ** Procedure: Above.
9470        End If
9480      End If
9490    End With

EXITP:
9500    Exit Sub

ERRH:
9510    Select Case ERR.Number
        Case Else
9520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9530    End Select
9540    Resume EXITP

End Sub

Private Sub opgLocContact_AfterUpdate()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_AfterUpdate"

9610    With Me
9620      Select Case .opgLocContact
          Case .opgLocContact_optUSA.OptionValue
9630        .Contact_State.Enabled = True
9640        .Contact_State.BackStyle = acBackStyleNormal
9650        .Contact_State.BorderColor = CLR_LTBLU2
9660        .Contact_Zip.Enabled = True
9670        .Contact_Zip.BackStyle = acBackStyleNormal
9680        .Contact_Zip.BorderColor = CLR_LTBLU2
9690        .Contact_Country.Enabled = False
9700        .Contact_Country.BackStyle = acBackStyleTransparent
9710        .Contact_Country.BorderColor = WIN_CLR_DISR
9720        .Contact_PostalCode.Enabled = False
9730        .Contact_PostalCode.BackStyle = acBackStyleTransparent
9740        .Contact_PostalCode.BorderColor = WIN_CLR_DISR
9750        Select Case blnIsOpen
            Case True
9760          blnIsOpen = False
9770        Case False
9780          .Contact_State.SetFocus
9790        End Select
9800      Case .opgLocContact_optOther.OptionValue
9810        .Contact_Country.Enabled = True
9820        .Contact_Country.BackStyle = acBackStyleNormal
9830        .Contact_Country.BorderColor = CLR_LTBLU2
9840        .Contact_PostalCode.Enabled = True
9850        .Contact_PostalCode.BackStyle = acBackStyleNormal
9860        .Contact_PostalCode.BorderColor = CLR_LTBLU2
9870        .Contact_State.Enabled = False
9880        .Contact_State.BackStyle = acBackStyleTransparent
9890        .Contact_State.BorderColor = WIN_CLR_DISR
9900        .Contact_Zip.Enabled = False
9910        .Contact_Zip.BackStyle = acBackStyleTransparent
9920        .Contact_Zip.BorderColor = WIN_CLR_DISR
9930        Select Case blnIsOpen
            Case True
9940          blnIsOpen = False
9950        Case False
9960          .Contact_Country.SetFocus
9970        End Select
9980      End Select
9990    End With

EXITP:
10000   Exit Sub

ERRH:
10010   Select Case ERR.Number
        Case Else
10020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10030   End Select
10040   Resume EXITP

End Sub

Private Sub opgLocContact_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10120   intShiftDown = (Shift And acShiftMask) > 0
10130   intAltDown = (Shift And acAltMask) > 0
10140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10170       With Me
10180         intRetVal = 0
10190         If .Contact_State.Enabled = True Then
10200           .Contact_State.SetFocus
10210         Else
10220           .Contact_Country.SetFocus
10230         End If
10240       End With
10250     End Select
10260   End If

        ' ** Shift keys.
10270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10290       With Me
10300         intRetVal = 0
10310         .Contact_City.SetFocus
10320       End With
10330     End Select
10340   End If

EXITP:
10350   KeyCode = intRetVal
10360   Exit Sub

ERRH:
10370   intRetVal = 0
10380   Select Case ERR.Number
        Case Else
10390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10400   End Select
10410   Resume EXITP

End Sub

Private Sub opgLocContact_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10520   intShiftDown = (Shift And acShiftMask) > 0
10530   intAltDown = (Shift And acAltMask) > 0
10540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10570       With Me
10580         intRetVal = 0
10590         If .Contact_State.Enabled = True Then
10600           .Contact_State.SetFocus
10610         Else
10620           .Contact_Country.SetFocus
10630         End If
10640       End With
10650     End Select
10660   End If

        ' ** Shift keys.
10670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10690       With Me
10700         intRetVal = 0
10710         .Contact_City.SetFocus
10720       End With
10730     End Select
10740   End If

EXITP:
10750   KeyCode = intRetVal
10760   Exit Sub

ERRH:
10770   intRetVal = 0
10780   Select Case ERR.Number
        Case Else
10790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10800   End Select
10810   Resume EXITP

End Sub

Private Sub Contact_State_Enter()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_Enter"

10910   With Me
10920     If .Contact_State.text = vbNullString Then
10930       ExpandCombo .Contact_State  ' ** Module Procedure: modUtilities.
10940     End If
10950   End With

EXITP:
10960   Exit Sub

ERRH:
10970   Select Case ERR.Number
        Case Else
10980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10990   End Select
11000   Resume EXITP

End Sub

Private Sub Contact_State_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_AfterUpdate"

11110   With Me
11120     If IsNull(.Contact_State) = False Then
11130       If Trim(.Contact_State) <> vbNullString Then
11140         .Contact_State.Value = FormatUpperCase(Nz(.Contact_State, vbNullString))
11150         Select Case IsNull(.Contact_State_Pref)
              Case True
11160           .Contact_State_Pref = .Contact_State
11170         Case False
11180           If .Contact_State_Pref <> .Contact_State Then
11190             .Contact_State_Pref = .Contact_State
11200           End If
11210         End Select
11220       End If
11230     End If
11240   End With

EXITP:
11250   Exit Sub

ERRH:
11260   Select Case ERR.Number
        Case Else
11270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11280   End Select
11290   Resume EXITP

End Sub

Private Sub Contact_State_NotInList(NewData As String, Response As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_NotInList"

11310   With Me
11320     .Contact_State = Null
11330     MsgBox "Please select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
11340     Response = acDataErrContinue
11350   End With

EXITP:
11360   Exit Sub

ERRH:
11370   Select Case ERR.Number
        Case Else
11380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11390   End Select
11400   Resume EXITP

End Sub

Private Sub Contact_Zip_Exit(Cancel As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Exit"

11510   With Me
11520     If .Contact_Zip <> "" Then
11530       If Right(.Contact_Zip.text, 4) <> "____" And Right(.Contact_Zip.text, 4) <> "    " And _
                (Not IsNumeric(Mid(.Contact_Zip.text, 7, 1)) Or Not IsNumeric(Mid(.Contact_Zip.text, 8, 1)) Or _
                Not IsNumeric(Mid(.Contact_Zip.text, 9, 1)) Or Not IsNumeric(Mid(.Contact_Zip.text, 10, 1))) Then
11540         MsgBox "Please enter a valid zip code.", vbInformation + vbOKOnly, "Invalid Entry"
11550         Cancel = -1
11560       End If
11570     End If
11580   End With

EXITP:
11590   Exit Sub

ERRH:
11600   Select Case ERR.Number
        Case Else
11610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11620   End Select
11630   Resume EXITP

End Sub

Private Sub Contact_Email_AfterUpdate()
' ** Make sure it's a valid email address.

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String

11710   With Me
11720     If IsNull(.Contact_Email) = False Then
11730       If Trim(.Contact_Email) <> vbNullString Then
11740         strTmp01 = Trim(.Contact_Email)
11750         intPos01 = InStr(strTmp01, "@")
11760         If intPos01 > 0 Then
11770           intPos01 = InStr(intPos01, strTmp01, ".")
11780           If intPos01 > 0 Then
11790             .Contact_MailTo = "MailTo:" & strTmp01
11800             cmdSave_Click  ' ** Procedure: Above.
11810           Else
11820             Beep
11830             MsgBox "Entry must be a valid email address.", vbInformation + vbOKOnly, "Invalid Entry"
11840             .Contact_Email.SetFocus
11850           End If
11860         Else
11870           Beep
11880           MsgBox "Entry must be a valid email address.", vbInformation + vbOKOnly, "Invalid Entry"
11890           .Contact_Email.SetFocus
11900         End If
11910       End If
11920     End If
11930   End With

EXITP:
11940   Exit Sub

ERRH:
11950   Select Case ERR.Number
        Case Else
11960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11970   End Select
11980   Resume EXITP

End Sub

Private Sub Contact_Email_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12020   intShiftDown = (Shift And acShiftMask) > 0
12030   intAltDown = (Shift And acAltMask) > 0
12040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12070       With Me
12080         intRetVal = 0
12090         If .cmdOK.Enabled = True Then
12100           .cmdOK.SetFocus
12110         Else
12120           .cmdCancel.SetFocus
12130         End If
12140       End With
12150     End Select
12160   End If

        ' ** Shift keys.
12170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12190       With Me
12200         intRetVal = 0
12210         .Contact_Fax.SetFocus
12220       End With
12230     End Select
12240   End If

EXITP:
12250   KeyCode = intRetVal
12260   Exit Sub

ERRH:
12270   intRetVal = 0
12280   Select Case ERR.Number
        Case Else
12290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12300   End Select
12310   Resume EXITP

End Sub

Private Sub DisableAll(blnDisable As Boolean)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "DisableAll"

12410   With Me
12420     Select Case blnDisable
          Case True
12430       .opgConNum = 0
12440       opgConNum_AfterUpdate  ' ** Procedure: Above.
12450       .cmdOK.Enabled = False
12460       DoEvents
12470       .Contact_ID.ForeColor = MY_CLR_LTBGE
12480       .Contact_lbl.ForeColor = WIN_CLR_DISF
12490       .Contact_lbl_dim_hi.Visible = True
12500       .opgConNum_optOne.Enabled = False
12510       .opgConNum_optTwo.Enabled = False
12520       .opgConNum.Enabled = False
12530       .opgConNum_optOne_lbl.Top = (lngOptOneLbl_Top - lngTpp)
12540       .opgConNum_optOne_lbl.Left = (lngOptOneLbl_Left - lngTpp)
12550       .opgConNum_optOne_lbl_dim_hi.Visible = True
12560       .opgConNum_optTwo_lbl.Top = (lngOptTwoLbl_Top - lngTpp)
12570       .opgConNum_optTwo_lbl.Left = (lngOptTwoLbl_Left - lngTpp)
12580       .opgConNum_optTwo_lbl_dim_hi.Visible = True
12590       .Contact_Name.Enabled = False
12600       .Contact_Name.BackStyle = acBackStyleTransparent
12610       .Contact_Name.BorderColor = WIN_CLR_DISR
12620       .Contact_Address1.Enabled = False
12630       .Contact_Address1.BackStyle = acBackStyleTransparent
12640       .Contact_Address1.BorderColor = WIN_CLR_DISR
12650       .Contact_Address2.Enabled = False
12660       .Contact_Address2.BackStyle = acBackStyleTransparent
12670       .Contact_Address2.BorderColor = WIN_CLR_DISR
12680       .Contact_City.Enabled = False
12690       .Contact_City.BackStyle = acBackStyleTransparent
12700       .Contact_City.BorderColor = WIN_CLR_DISR
12710       .opgLocContact.Enabled = False
12720       .Contact_State.Enabled = False
12730       .Contact_State.BackStyle = acBackStyleTransparent
12740       .Contact_State.BorderColor = WIN_CLR_DISR
12750       .Contact_Zip.Enabled = False
12760       .Contact_Zip.BackStyle = acBackStyleTransparent
12770       .Contact_Zip.BorderColor = WIN_CLR_DISR
12780       .Contact_Country.Enabled = False
12790       .Contact_Country.BackStyle = acBackStyleTransparent
12800       .Contact_Country.BorderColor = WIN_CLR_DISR
12810       .Contact_PostalCode.Enabled = False
12820       .Contact_PostalCode.BackStyle = acBackStyleTransparent
12830       .Contact_PostalCode.BorderColor = WIN_CLR_DISR
12840       .Contact_Phone1.Enabled = False
12850       .Contact_Phone1.BackStyle = acBackStyleTransparent
12860       .Contact_Phone1.BorderColor = WIN_CLR_DISR
12870       .Contact_Phone2.Enabled = False
12880       .Contact_Phone2.BackStyle = acBackStyleTransparent
12890       .Contact_Phone2.BorderColor = WIN_CLR_DISR
12900       .Contact_Fax.Enabled = False
12910       .Contact_Fax.BackStyle = acBackStyleTransparent
12920       .Contact_Fax.BorderColor = WIN_CLR_DISR
12930       .Contact_Email.Enabled = False
12940       .Contact_Email.BackStyle = acBackStyleTransparent
12950       .Contact_Email.BorderColor = WIN_CLR_DISR
12960     Case False
12970       .Contact_ID.ForeColor = CLR_GRY3
12980       .Contact_lbl.ForeColor = CLR_DKGRY2
12990       .Contact_lbl_dim_hi.Visible = False
13000       .opgConNum.Enabled = True
13010       .opgConNum_optOne.Enabled = True
13020       .opgConNum_optTwo.Enabled = True
13030       .opgConNum_optOne_lbl.Top = lngOptOneLbl_Top
13040       .opgConNum_optOne_lbl.Left = lngOptOneLbl_Left
13050       .opgConNum_optOne_lbl_dim_hi.Visible = False
13060       .opgConNum_optTwo_lbl.Top = lngOptTwoLbl_Top
13070       .opgConNum_optTwo_lbl.Left = lngOptTwoLbl_Left
13080       .opgConNum_optTwo_lbl_dim_hi.Visible = False
13090       .Contact_Name.Enabled = True
13100       .Contact_Name.BackStyle = acBackStyleNormal
13110       .Contact_Name.BorderColor = CLR_LTBLU2
13120       .Contact_Address1.Enabled = True
13130       .Contact_Address1.BackStyle = acBackStyleNormal
13140       .Contact_Address1.BorderColor = CLR_LTBLU2
13150       .Contact_Address2.Enabled = True
13160       .Contact_Address2.BackStyle = acBackStyleNormal
13170       .Contact_Address2.BorderColor = CLR_LTBLU2
13180       .Contact_City.Enabled = True
13190       .Contact_City.BackStyle = acBackStyleNormal
13200       .Contact_City.BorderColor = CLR_LTBLU2
13210       .opgLocContact.Enabled = True
13220       Select Case .opgLocContact
            Case .opgLocContact_optUSA.OptionValue
13230         .Contact_State.Enabled = True
13240         .Contact_State.BackStyle = acBackStyleNormal
13250         .Contact_State.BorderColor = CLR_LTBLU2
13260         .Contact_Zip.Enabled = True
13270         .Contact_Zip.BackStyle = acBackStyleNormal
13280         .Contact_Zip.BorderColor = CLR_LTBLU2
13290         .Contact_Country.Enabled = False
13300         .Contact_Country.BackStyle = acBackStyleTransparent
13310         .Contact_Country.BorderColor = WIN_CLR_DISR
13320         .Contact_PostalCode.Enabled = False
13330         .Contact_PostalCode.BackStyle = acBackStyleTransparent
13340         .Contact_PostalCode.BorderColor = WIN_CLR_DISR
13350       Case .opgLocContact_optOther.OptionValue
13360         .Contact_Country.Enabled = True
13370         .Contact_Country.BackStyle = acBackStyleNormal
13380         .Contact_Country.BorderColor = CLR_LTBLU2
13390         .Contact_PostalCode.Enabled = True
13400         .Contact_PostalCode.BackStyle = acBackStyleNormal
13410         .Contact_PostalCode.BorderColor = CLR_LTBLU2
13420         .Contact_State.Enabled = False
13430         .Contact_State.BackStyle = acBackStyleTransparent
13440         .Contact_State.BorderColor = WIN_CLR_DISR
13450         .Contact_Zip.Enabled = False
13460         .Contact_Zip.BackStyle = acBackStyleTransparent
13470         .Contact_Zip.BorderColor = WIN_CLR_DISR
13480       End Select
13490       .Contact_Phone1.Enabled = True
13500       .Contact_Phone1.BackStyle = acBackStyleNormal
13510       .Contact_Phone1.BorderColor = CLR_LTBLU2
13520       .Contact_Phone2.Enabled = True
13530       .Contact_Phone2.BackStyle = acBackStyleNormal
13540       .Contact_Phone2.BorderColor = CLR_LTBLU2
13550       .Contact_Fax.Enabled = True
13560       .Contact_Fax.BackStyle = acBackStyleNormal
13570       .Contact_Fax.BorderColor = CLR_LTBLU2
13580       .Contact_Email.Enabled = True
13590       .Contact_Email.BackStyle = acBackStyleNormal
13600       .Contact_Email.BorderColor = CLR_LTBLU2
13610       .cmdOK.Enabled = True
13620     End Select
13630   End With

EXITP:
13640   Exit Sub

ERRH:
13650   Select Case ERR.Number
        Case Else
13660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13670   End Select
13680   Resume EXITP

End Sub
