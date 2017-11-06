VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountContacts_Sub2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountContacts_Sub2"

'VGC 07/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd on frmAccountContacts}
' **   Enable Country:   B {chkEnableCountry_Compact on frmAccountContacts}
' **   Close:            C {cmdClose on frmAccountContacts}
' **   View Wide:        D {opgView_optWide on frmAccountContacts}
' **   Reset:            E {cmdReset on frmAccountContacts}
' **   Excel:            L {cmdExcel on frmAccountContacts}
' **   View Compact:     M {opgView_optCompact on frmAccountContacts}
' **   Account Num:      N {chkShowAcctNum on frmAccountContacts}
' **   Word:             O {cmdWord on frmAccountContacts}
' **   Print:            P {cmdPrint on frmAccountContacts}
' **   Short Name:       S {chkShowShortName on frmAccountContacts}
' **   Delete:           T {cmdDelete on frmAccountContacts}
' **   Update:           U {cmdUpdate on frmAccountContacts}
' **   Preview:          W {cmdPreview on frmAccountContacts}
' **   Exit:             X {cmdClose on frmAccountContacts}
' **   Always Update:    Y {chkAlwaysUpdate on frmAccountContacts}

' ** Shortcut Ctrl keys responsive from this form:
' **   Send Email:       E {Contact_Email_cmd}
' **   Page Count:       G {chkPageOf on frmAccountContacts}
' **   Other:            O {opgLocContact_optOther}
' **   Save:             S {cmdSave}
' **   USA:              U {opgLocContact_optUSA}

' ******************************************
' ** opgFormat and opgPrint shortcut keys:
' ******************************************

' ** Shortcut Ctrl keys responsive from this form:
' **   Zip Unfrmat:      D {opgFormatZip_optUnformatted on frmAccountContacts}
' **   Zip Frmat:        F {opgFormatZip_optFormatted on frmAccountContacts}
' **   Name Last/First:  I {opgFormatName_optLastFirst on frmAccountContacts}
' **   Phone Unfrmat:    M {opgFormatPhone_optUnformatted on frmAccountContacts}
' **   Phone Frmat:      R {opgFormatPhone_optFormatted on frmAccountContacts}
' **   Name As Written:  T {opgFormatName_optAsWritten on frmAccountContacts}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   CSZ Separate:     A {opgPrintCSZ_optSeparate, opgPrintCSZCP_optSeparate on frmAccountContacts}
' **   CSZ Combined:     B {opgPrintCSZ_optCombined, opgPrintCSZCP_optCombined on frmAccountContacts}
' **   Addr Combined:    C {opgPrintAddress_optCombined on frmAccountContacts}
' **   Addr None:        N {opgPrintAddress_optNone on frmAccountContacts}
' **   CSZ None:         O {opgPrintCSZ_optNone, opgPrintCSZCP_optNone on frmAccountContacts}
' **   Addr Separate:    S {opgPrintAddress_optSeparate on frmAccountContacts}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Phone Combined:   C {opgPrintPhone_optCombined on frmAccountContacts}
' **   Email None:       E {opgPrintEmail_optNone on frmAccountContacts}
' **   Phone None:       N {opgPrintPhone_optNone on frmAccountContacts}
' **   Fax None:         O {opgPrintFax_optNone on frmAccountContacts}
' **   Fax Print:        P {opgPrintFax_optPrint on frmAccountContacts}
' **   Phone Separate:   S {opgPrintPhone_optSeparate on frmAccountContacts}
' **   Email Print:      T {opgPrintEmail_optPrint on frmAccountContacts}
' **   Phone 1 Only:     Y {opgPrintPhone_opt1Only on frmAccountContacts}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[alphasort], [Contact_Number]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine As String = "Sort_line"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private lngForm_Width As Long, lngCity_Left As Long, lngCountry_Left As Long, lngCountry_Width As Long
Private lngPhone_Left As Long, lngEmail_Left As Long, lngEmailCmd_Left As Long, lngLocContact_Width As Long

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Left As Long, lngSortLine_Width As Long
Private varOldValue As Variant, blnAddComma As Boolean
Private lngRecsCur As Long, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmAccountContacts" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
190         StateCodeQrySet Me  ' ** Module Function: modPreferenceFuncs.

            ' ** Just get these out of the way.
200         .form_width_line.Left = 0&

210         lngForm_Width = .Width
220         lngSortLbl_Width = .Sort_lbl.Width
230         lngCity_Left = .Contact_City.Left
240         lngCountry_Left = .Contact_Country.Left
250         lngCountry_Width = (.Contact_Phone1.Left - .Contact_Country.Left)  ' ** Includes separation.
260         lngPhone_Left = .Contact_Phone1.Left
270         lngEmail_Left = .Contact_Email.Left
280         lngEmailCmd_Left = .Contact_Email_cmd.Left
290         lngLocContact_Width = (.Contact_City.Left - .opgLocContact.Left)  ' ** Includes separation.

300       Else
310         Cancel = -1
320       End If
330     End With

340     If Cancel = -1 Then
350       Beep
360       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
370     End If

EXITP:
380     Exit Sub

ERRH:
390     Select Case ERR.Number
        Case Else
400       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
410     End Select
420     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me

520       .accountno1.Enabled = False
530       .accountno1.Locked = True
540       .shortname1.Enabled = False
550       .shortname1.Locked = True
560       .accountno1.ForeColor = CLR_DISABLED_FG
570       .accountno1.BackColor = CLR_DISABLED_BG
580       .shortname1.ForeColor = CLR_DISABLED_FG
590       .shortname1.BackColor = CLR_DISABLED_BG
600       .Contact_Number.ForeColor = CLR_DISABLED_FG
610       .Contact_Number.BackColor = CLR_DISABLED_BG

620       SortNow THIS_PROC  ' ** Procedure: Below.

630     End With

EXITP:
640     Exit Sub

ERRH:
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub Form_Current()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

710     With Me
720       Select Case .Parent.chkEnableCountry_Compact
          Case True
730         Select Case .opgLocContact
            Case .opgLocContact_optUSA.OptionValue
740           .Contact_State.Enabled = True
750           .Contact_State.Locked = False
760           .Contact_Zip.Enabled = True
770           .Contact_Zip.Locked = False
780           .Contact_Zip_Format.Enabled = True
790           .Contact_Zip_Format.Locked = False
800           .Contact_Country.Enabled = False
810           .Contact_Country.Locked = True
820           .Contact_PostalCode.Enabled = False
830           .Contact_PostalCode.Locked = True
840         Case .opgLocContact_optOther.OptionValue
850           .Contact_State.Enabled = False
860           .Contact_State.Locked = True
870           .Contact_Zip.Enabled = False
880           .Contact_Zip.Locked = True
890           .Contact_Zip_Format.Enabled = False
900           .Contact_Zip_Format.Locked = True
910           .Contact_Country.Enabled = True
920           .Contact_Country.Locked = False
930           .Contact_PostalCode.Enabled = True
940           .Contact_PostalCode.Locked = False
950         End Select
960       Case False
970         .Contact_State.Enabled = True
980         .Contact_State.Locked = False
990         .Contact_Zip.Enabled = True
1000        .Contact_Zip.Locked = False
1010        .Contact_Zip_Format.Enabled = True
1020        .Contact_Zip_Format.Locked = False
1030        .Contact_Country.Enabled = False
1040        .Contact_Country.Locked = True
1050        .Contact_PostalCode.Enabled = False
1060        .Contact_PostalCode.Locked = True
1070      End Select
1080      Select Case IsNull(.Contact_Email)
          Case True
1090        .Contact_Email_cmd.Enabled = False
1100      Case False
1110        .Contact_Email_cmd.Enabled = True
1120      End Select
1130    End With

EXITP:
1140    Exit Sub

ERRH:
1150    Select Case ERR.Number
        Case Else
1160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1170    End Select
1180    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Exit:             X {cmdClose on frmAccountContacts}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Zip Unfrmat:      D {opgFormatZip_optUnformatted on frmAccountContacts}
        ' **   Send Email:       E {Contact_Email_cmd}
        ' **   Zip Frmat:        F {opgFormatZip_optFormatted on frmAccountContacts}
        ' **   Page Count:       G {chkPageOf on frmAccountContacts}
        ' **   Name Last/First:  I {opgFormatName_optLastFirst on frmAccountContacts}
        ' **   Phone Unfrmat:    M {opgFormatPhone_optUnformatted on frmAccountContacts}
        ' **   Other:            O {opgLocContact_optOther}
        ' **   Phone Frmat:      R {opgFormatPhone_optFormatted on frmAccountContacts}
        ' **   Save:             S {cmdSave}
        ' **   Name As Written:  T {opgFormatName_optAsWritten on frmAccountContacts}
        ' **   USA:              U {opgLocContact_optUSA}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   CSZ Separate:     A {opgPrintCSZ_optSeparate, opgPrintCSZCP_optSeparate on frmAccountContacts}
        ' **   CSZ Combined:     B {opgPrintCSZ_optCombined, opgPrintCSZCP_optCombined on frmAccountContacts}
        ' **   Addr Combined:    C {opgPrintAddress_optCombined on frmAccountContacts}
        ' **   Addr None:        N {opgPrintAddress_optNone on frmAccountContacts}
        ' **   CSZ None:         O {opgPrintCSZ_optNone, opgPrintCSZCP_optNone on frmAccountContacts}
        ' **   Addr Separate:    S {opgPrintAddress_optSeparate on frmAccountContacts}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Phone Combined:   C {opgPrintPhone_optCombined on frmAccountContacts}
        ' **   Email None:       E {opgPrintEmail_optNone on frmAccountContacts}
        ' **   Phone None:       N {opgPrintPhone_optNone on frmAccountContacts}
        ' **   Fax None:         O {opgPrintFax_optNone on frmAccountContacts}
        ' **   Fax Print:        P {opgPrintFax_optPrint on frmAccountContacts}
        ' **   Phone Separate:   S {opgPrintPhone_optSeparate on frmAccountContacts}
        ' **   Email Print:      T {opgPrintEmail_optPrint on frmAccountContacts}
        ' **   Phone 1 Only:     Y {opgPrintPhone_opt1Only on frmAccountContacts}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyUp
1270        intRetVal = 0
1280        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1290      Case vbKeyDown
1300        intRetVal = 0
1310        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1320      Case vbKeyEscape
1330        With Me
1340          intRetVal = 0
1350          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountContacts.
1360        End With
1370      End Select
1380    End If

        ' ** Alt keys.
1390    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1400      Select Case intRetVal
          Case vbKeyB
1410        With Me
1420          intRetVal = 0
1430          Select Case .Parent.opgView
              Case .Parent.opgView_optWide.OptionValue
1440            If .Parent.chkEnableCountry_Wide.Enabled = True Then
1450              DoCmd.SelectObject acForm, .Parent.Name, False
1460              .Parent.chkEnableCountry_Wide.SetFocus
1470            Else
1480              Beep
1490            End If
1500          Case .Parent.opgView_optCompact.OptionValue
1510            If .Parent.chkEnableCountry_Compact.Enabled = True Then
1520              DoCmd.SelectObject acForm, .Parent.Name, False
1530              .Parent.chkEnableCountry_Compact.SetFocus
1540            Else
1550              Beep
1560            End If
1570          End Select
1580        End With
1590      Case vbKeyX
1600        With Me
1610          intRetVal = 0
1620          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountContacts.
1630        End With
1640      End Select
1650    End If

        ' ** Ctrl keys.
1660    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1670      Select Case intRetVal
          Case vbKeyD
            ' **   Zip Unfrmat:      D {opgFormatZip_optUnformatted}
1680        With Me
1690          intRetVal = 0
1700          DoCmd.SelectObject acForm, .Parent.Name, False
1710          .Parent.opgFormatZip = .Parent.opgFormatZip_optUnformatted.OptionValue
1720          .Parent.opgFormatZip.SetFocus
1730          .Parent.opgFormatZip_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
1740        End With
1750      Case vbKeyE
1760        With Me
1770          intRetVal = 0
1780          If .Contact_Email_cmd.Visible = True And .Contact_Email_cmd.Enabled = True Then
1790            .Contact_Email_cmd.SetFocus
1800            Contact_Email_cmd_Click  ' ** Procedure: Below.
1810          End If
1820        End With
1830      Case vbKeyF
            ' **   Zip Frmat:        F {opgFormatZip_optFormatted}
1840        With Me
1850          intRetVal = 0
1860          DoCmd.SelectObject acForm, .Parent.Name, False
1870          .Parent.opgFormatZip = .Parent.opgFormatZip_optFormatted.OptionValue
1880          .Parent.opgFormatZip.SetFocus
1890          .Parent.opgFormatZip_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
1900        End With
1910      Case vbKeyG
1920        With Me
1930          intRetVal = 0
1940          If .Parent.chkPageOf.Enabled = True Then
1950            DoCmd.SelectObject acForm, .Parent.Name, False
1960            .Parent.chkPageOf.SetFocus
1970            .Parent.chkPageOf = (Not .Parent.chkPageOf)
1980            .Parent.chkPageOf_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
1990          End If
2000        End With
2010      Case vbKeyI
            ' **   Name Last/First:  I {opgFormatName_optLastFirst}
2020        With Me
2030          intRetVal = 0
2040          DoCmd.SelectObject acForm, .Parent.Name, False
2050          .Parent.opgFormatName = .Parent.opgFormatName_optLastFirst.OptionValue
2060          .Parent.opgFormatName.SetFocus
2070          .Parent.opgFormatName_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2080        End With
2090      Case vbKeyM
            ' **   Phone Unfrmat:    M {opgFormatPhone_optUnformatted}
2100        With Me
2110          intRetVal = 0
2120          DoCmd.SelectObject acForm, .Parent.Name, False
2130          .Parent.opgFormatPhone = .Parent.opgFormatPhone_optUnformatted.OptionValue
2140          .Parent.opgFormatPhone.SetFocus
2150          .Parent.opgFormatPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2160        End With
2170      Case vbKeyO
2180        With Me
2190          intRetVal = 0
2200          Select Case .Parent.opgView
              Case .Parent.opgView_optWide.OptionValue
2210            Beep
2220          Case .Parent.opgView_optCompact.OptionValue
2230            If .opgLocContact.Visible = True And .opgLocContact.Enabled = True Then
2240              .opgLocContact.SetFocus
2250              .opgLocContact = .opgLocContact_optOther.OptionValue
2260              opgLocContact_AfterUpdate  ' ** Procedure: Below.
2270            Else
2280              Beep
2290            End If
2300          End Select
2310        End With
2320      Case vbKeyR
            ' **   Phone Frmat:      R {opgFormatPhone_optFormatted}
2330        With Me
2340          intRetVal = 0
2350          DoCmd.SelectObject acForm, .Parent.Name, False
2360          .Parent.opgFormatPhone = .Parent.opgFormatPhone_optFormatted.OptionValue
2370          .Parent.opgFormatPhone.SetFocus
2380          .Parent.opgFormatPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2390        End With
2400      Case vbKeyS
2410        intRetVal = 0
2420        cmdSave_Click  ' ** Procedure: Below.
2430      Case vbKeyT
            ' **   Name As Written:  T {opgFormatName_optAsWritten}
2440        With Me
2450          intRetVal = 0
2460          DoCmd.SelectObject acForm, .Parent.Name, False
2470          .Parent.opgFormatName = .Parent.opgFormatName_optAsWritten.OptionValue
2480          .Parent.opgFormatName.SetFocus
2490          .Parent.opgFormatName_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2500        End With
2510      Case vbKeyU
2520        With Me
2530          intRetVal = 0
2540          Select Case .Parent.opgView
              Case .Parent.opgView_optWide.OptionValue
2550            Beep
2560          Case .Parent.opgView_optCompact.OptionValue
2570            If .opgLocContact.Visible = True And .opgLocContact.Enabled = True Then
2580              .opgLocContact.SetFocus
2590              .opgLocContact = .opgLocContact_optUSA.OptionValue
2600              opgLocContact_AfterUpdate  ' ** Procedure: Below.
2610            Else
2620              Beep
2630            End If
2640          End Select
2650        End With
2660      Case vbKeyTab, vbKeyReturn
2670        With Me
2680          intRetVal = 0
2690          DoCmd.SelectObject acForm, .Parent.Name, False
2700          .Parent.opgView.SetFocus
2710        End With
2720      Case vbKeyPageUp
2730        intRetVal = 0
2740        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2750      Case vbKeyPageDown
2760        intRetVal = 0
2770        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2780      Case vbKeyHome
2790        With Me
2800          intRetVal = 0
2810          If .Contact_Name.Visible = True And .Contact_Name.Enabled = True Then
2820            .Contact_Name.SetFocus
2830          Else
2840            .Contact_Name_LastFirst.SetFocus
2850          End If
2860        End With
2870      Case vbKeyEnd
2880        With Me
2890          intRetVal = 0
2900          .Contact_Email.SetFocus
2910        End With
2920      End Select
2930    End If

        ' ** Ctrl-Alt keys.
2940    If intCtrlDown And intAltDown And (Not intShiftDown) Then
2950      Select Case intRetVal
          Case vbKeyA
            ' **   CSZ Separate:     A {opgPrintCSZ_optSeparate, opgPrintCSZCP_optSeparate}
2960        With Me
2970          intRetVal = 0
2980          DoCmd.SelectObject acForm, .Parent.Name, False
2990          If .Parent.opgPrintCSZ.Visible = True And .Parent.opgPrintCSZ.Enabled = True Then
3000            .Parent.opgPrintCSZ = .Parent.opgPrintCSZ_optSeparate.OptionValue
3010            .Parent.opgPrintCSZ.SetFocus
3020            .Parent.opgPrintCSZ_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3030          ElseIf .Parent.opgPrintCSZCP.Visible = True And .Parent.opgPrintCSZCP.Enabled = True Then
3040            .Parent.opgPrintCSZCP = .Parent.opgPrintCSZCP_optSeparate.OptionValue
3050            .Parent.opgPrintCSZCP.SetFocus
3060            .Parent.opgPrintCSZCP_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3070          End If
3080        End With
3090      Case vbKeyB
            ' **   CSZ Combined:     B {opgPrintCSZ_optCombined, opgPrintCSZCP_optCombined}
3100        With Me
3110          intRetVal = 0
3120          DoCmd.SelectObject acForm, .Parent.Name, False
3130          If .Parent.opgPrintCSZ.Visible = True And .Parent.opgPrintCSZ.Enabled = True Then
3140            .Parent.opgPrintCSZ = .Parent.opgPrintCSZ_optCombined.OptionValue
3150            .Parent.opgPrintCSZ.SetFocus
3160            .Parent.opgPrintCSZ_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3170          ElseIf .Parent.opgPrintCSZCP.Visible = True And .Parent.opgPrintCSZCP.Enabled = True Then
3180            .Parent.opgPrintCSZCP = .Parent.opgPrintCSZCP_optCombined.OptionValue
3190            .Parent.opgPrintCSZCP.SetFocus
3200            .Parent.opgPrintCSZCP_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3210          End If
3220        End With
3230      Case vbKeyC
            ' **   Addr Combined:    C {opgPrintAddress_optCombined}
3240        With Me
3250          intRetVal = 0
3260          DoCmd.SelectObject acForm, .Parent.Name, False
3270          .Parent.opgPrintAddress = .Parent.opgPrintAddress_optCombined.OptionValue
3280          .Parent.opgPrintAddress.SetFocus
3290          .Parent.opgPrintAddress_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3300        End With
3310      Case vbKeyN
            ' **   Addr None:        N {opgPrintAddress_optNone}
3320        With Me
3330          intRetVal = 0
3340          DoCmd.SelectObject acForm, .Parent.Name, False
3350          .Parent.opgPrintAddress = .Parent.opgPrintAddress_optNone.OptionValue
3360          .Parent.opgPrintAddress.SetFocus
3370          .Parent.opgPrintAddress_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3380        End With
3390      Case vbKeyO
            ' **   CSZ None:         O {opgPrintCSZ_optNone, opgPrintCSZCP_optNone}
3400        With Me
3410          intRetVal = 0
3420          DoCmd.SelectObject acForm, .Parent.Name, False
3430          If .Parent.opgPrintCSZ.Visible = True And .Parent.opgPrintCSZ.Enabled = True Then
3440            .Parent.opgPrintCSZ = .Parent.opgPrintCSZ_optNone.OptionValue
3450            .Parent.opgPrintCSZ.SetFocus
3460            .Parent.opgPrintCSZ_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3470          ElseIf .Parent.opgPrintCSZCP.Visible = True And .Parent.opgPrintCSZCP.Enabled = True Then
3480            .Parent.opgPrintCSZCP = .Parent.opgPrintCSZCP_optNone.OptionValue
3490            .Parent.opgPrintCSZCP.SetFocus
3500            .Parent.opgPrintCSZCP_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3510          End If
3520        End With
3530      Case vbKeyS
            ' **   Addr Separate:    S {opgPrintAddress_optSeparate}
3540        With Me
3550          intRetVal = 0
3560          DoCmd.SelectObject acForm, .Parent.Name, False
3570          .Parent.opgPrintAddress = .Parent.opgPrintAddress_optSeparate.OptionValue
3580          .Parent.opgPrintAddress.SetFocus
3590          .Parent.opgPrintAddress_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3600        End With
3610      End Select
3620    End If

        ' ** Ctrl-Shift keys.
3630    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3640      Select Case intRetVal
          Case vbKeyC
            ' **   Phone Combined:   C {opgPrintPhone_optCombined}
3650        With Me
3660          intRetVal = 0
3670          DoCmd.SelectObject acForm, .Parent.Name, False
3680          .Parent.opgPrintPhone = .Parent.opgPrintPhone_optCombined.OptionValue
3690          .Parent.opgPrintPhone.SetFocus
3700          .Parent.opgPrintPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3710        End With
3720      Case vbKeyE
            ' **   Email None:       E {opgPrintEmail_optNone}
3730        With Me
3740          intRetVal = 0
3750          DoCmd.SelectObject acForm, .Parent.Name, False
3760          .Parent.opgPrintEmail = .Parent.opgPrintEmail_optNone.OptionValue
3770          .Parent.opgPrintEmail.SetFocus
3780          .Parent.opgPrintEmail_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3790        End With
3800      Case vbKeyF
3810        With Me
3820          intRetVal = 0
3830          DoCmd.SelectObject acForm, .Parent.Name, False
3840          .Parent.FocusHolder.SetFocus
3850        End With
3860      Case vbKeyN
            ' **   Phone None:       N {opgPrintPhone_optNone}
3870        With Me
3880          intRetVal = 0
3890          DoCmd.SelectObject acForm, .Parent.Name, False
3900          .Parent.opgPrintPhone = .Parent.opgPrintPhone_optNone.OptionValue
3910          .Parent.opgPrintPhone.SetFocus
3920          .Parent.opgPrintPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3930        End With
3940      Case vbKeyO
            ' **   Fax None:         O {opgPrintFax_optNone}
3950        With Me
3960          intRetVal = 0
3970          DoCmd.SelectObject acForm, .Parent.Name, False
3980          .Parent.opgPrintFax = .Parent.opgPrintFax_optNone.OptionValue
3990          .Parent.opgPrintFax.SetFocus
4000          .Parent.opgPrintFax_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
4010        End With
4020      Case vbKeyP
            ' **   Fax Print:        P {opgPrintFax_optPrint}
4030        With Me
4040          intRetVal = 0
4050          DoCmd.SelectObject acForm, .Parent.Name, False
4060          .Parent.opgPrintFax = .Parent.opgPrintFax_optPrint.OptionValue
4070          .Parent.opgPrintFax.SetFocus
4080          .Parent.opgPrintFax_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
4090        End With
4100      Case vbKeyS
            ' **   Phone Separate:   S {opgPrintPhone_optSeparate}
4110        With Me
4120          intRetVal = 0
4130          DoCmd.SelectObject acForm, .Parent.Name, False
4140          .Parent.opgPrintPhone = .Parent.opgPrintPhone_optSeparate.OptionValue
4150          .Parent.opgPrintPhone.SetFocus
4160          .Parent.opgPrintPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
4170        End With
4180      Case vbKeyT
            ' **   Email Print:      T {opgPrintEmail_optPrint}
4190        With Me
4200          intRetVal = 0
4210          DoCmd.SelectObject acForm, .Parent.Name, False
4220          .Parent.opgPrintEmail = .Parent.opgPrintEmail_optPrint.OptionValue
4230          .Parent.opgPrintEmail.SetFocus
4240          .Parent.opgPrintEmail_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
4250        End With
4260      Case vbKeyY
            ' **   Phone 1 Only:     Y {opgPrintPhone_opt1Only}
4270        With Me
4280          intRetVal = 0
4290          DoCmd.SelectObject acForm, .Parent.Name, False
4300          .Parent.opgPrintPhone = .Parent.opgPrintPhone_opt1Only.OptionValue
4310          .Parent.opgPrintPhone.SetFocus
4320          .Parent.opgPrintPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
4330        End With
4340      Case vbKeyTab, vbKeyReturn
4350        With Me
4360          intRetVal = 0
4370          DoCmd.SelectObject acForm, .Parent.Name, False
4380          .Parent.cmdClose.SetFocus
4390        End With
4400      End Select
4410    End If

        ' **   Phone Combined:   C {opgPrintPhone_optCombined}
        ' **   Email None:       E {opgPrintEmail_optNone}
        ' **   Phone None:       N {opgPrintPhone_optNone}
        ' **   Fax None:         O {opgPrintFax_optNone}
        ' **   Fax Print:        P {opgPrintFax_optPrint}
        ' **   Phone Separate:   S {opgPrintPhone_optSeparate}
        ' **   Email Print:      T {opgPrintEmail_optPrint}
        ' **   Phone 1 Only:     Y {opgPrintPhone_opt1Only}

EXITP:
4420    KeyCode = intRetVal
4430    Exit Sub

ERRH:
4440    intRetVal = 0
4450    Select Case ERR.Number
        Case Else
4460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4470    End Select
4480    Resume EXITP

End Sub

Public Sub cmdSave_Click()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

4510    DoCmd.RunCommand acCmdSaveRecord
4520    DoEvents

EXITP:
4530    Exit Sub

ERRH:
4540    Select Case ERR.Number
        Case Else
4550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4560    End Select
4570    Resume EXITP

End Sub

Private Sub accountno1_lbl_cmd_accountno_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno1_lbl_cmd_accountno_DblClick"

4610    SortNow "accountno_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub accountno1_lbl_cmd_shortname_DblClick(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno1_lbl_cmd_shortname_DblClick"

4710    SortNow "shortname_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case Else
4740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4750    End Select
4760    Resume EXITP

End Sub

Private Sub Contact_Number_lbl_DblClick(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Number_lbl_DblClick"

4810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4820    Exit Sub

ERRH:
4830    Select Case ERR.Number
        Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Private Sub Contact_Name_lbl_DblClick(Cancel As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_lbl_DblClick"

4910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub Contact_Name_Enter()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_Enter"

5010    With Me
5020      varOldValue = .Contact_Name
5030    End With

EXITP:
5040    Exit Sub

ERRH:
5050    Select Case ERR.Number
        Case 2427  ' ** You entered an expression that has no value.
          ' ** No records.
5060    Case Else
5070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5080    End Select
5090    Resume EXITP

End Sub

Private Sub Contact_Name_GotFocus()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_GotFocus"

5110    With Me
5120      .Contact_Name.SelLength = 0
5130      .Contact_Name.SelStart = 99
5140    End With

EXITP:
5150    Exit Sub

ERRH:
5160    Select Case ERR.Number
        Case Else
5170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5180    End Select
5190    Resume EXITP

End Sub

Private Sub Contact_Name_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_AfterUpdate"

        Dim varLast As Variant, varFirst As Variant, varSfx As Variant, varLastFirst As Variant
        Dim strTmp01 As String

5210    With Me
5220      If IsNull(.Contact_Name) = False Then
5230        If Trim(.Contact_Name) <> vbNullString Then
5240          strTmp01 = Trim(FixQuotes(.Contact_Name))  ' ** Module Function: modStringFuncs.
5250          If Compare_StringA_StringB(.Contact_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
5260            .Contact_Name = strTmp01
5270          End If
5280          .Contact_Changed = True

5290          cmdSave_Click  ' ** Procedure: Above.

5300          varLast = Null: varFirst = Null: varSfx = Null: varLastFirst = Null

5310          If InStr(strTmp01, "Regional Center") > 0 Or InStr(strTmp01, "(") > 0 Or InStr(strTmp01, "Cnty") > 0 Then
5320            varLast = strTmp01
5330          Else
5340            If InStr(GetLastWord(strTmp01, True), "~") > 0 Then  ' ** Module Function: modStringFuncs.
5350              varLast = Left(GetLastWord(strTmp01, True), (InStr(GetLastWord(strTmp01, True), "~") - 1))  ' ** Module Function: modStringFuncs.
5360            Else
5370              varLast = GetLastWord(strTmp01, True)  ' ** Module Function: modStringFuncs.
5380            End If
5390          End If

5400          varSfx = IIf(InStr(strTmp01, "Regional Center") > 0 Or InStr(strTmp01, "(") > 0 Or _
                InStr(strTmp01, "Cnty") > 0, Null, IIf(InStr(GetLastWord(strTmp01, True), "~") > 0, _
                Mid(GetLastWord(strTmp01, True), (InStr(GetLastWord(strTmp01, True), "~") + 1)), Null))  ' ** Module Function: modStringFuncs.
5410          varFirst = IIf(IsNull(strTmp01) = True, Null, IIf(varLast = strTmp01, Null, _
                Trim(Left(strTmp01, (InStr(strTmp01, varLast) - 1))) & IIf(IsNull(varSfx) = True, "", _
                IIf(Left(varSfx, 1) = ",", varSfx, " " & varSfx))))
5420          varLastFirst = IIf(IsNull(strTmp01) = True, Null, varLast & IIf(IsNull(varFirst) = True, "", ", " & varFirst))

5430          Select Case IsNull(varFirst)
              Case True
5440            If IsNull(.Contact_Name_First) = False Then
5450              .Contact_Name_First = Null
5460            End If
5470          Case False
5480            If Trim(varFirst) <> vbNullString Then
5490              .Contact_Name_First = varFirst
5500            Else
5510              .Contact_Name_First = Null
5520            End If
5530          End Select
5540          Select Case IsNull(varLast)
              Case True
5550            If IsNull(.Contact_Name_Last) = False Then
5560              .Contact_Name_Last = Null
5570            End If
5580          Case False
5590            If Trim(varLast) <> vbNullString Then
5600              .Contact_Name_Last = varLast
5610            Else
5620              .Contact_Name_Last = Null
5630            End If
5640          End Select
5650          Select Case IsNull(varLastFirst)
              Case True
5660            If IsNull(.Contact_Name_LastFirst) = False Then
5670              .Contact_Name_LastFirst = Null
5680            End If
5690          Case False
5700            If Trim(varLastFirst) <> vbNullString Then
5710              .Contact_Name_LastFirst = varLastFirst
5720            Else
5730              .Contact_Name_LastFirst = Null
5740            End If
5750          End Select

5760          cmdSave_Click  ' ** Procedure: Above.

5770        Else
5780          If IsNull(varOldValue) = False Then
5790            .Contact_Name = Null
5800            .Contact_Name_First = Null
5810            .Contact_Name_Last = Null
5820            .Contact_Name_LastFirst = Null
5830            .Contact_Changed = True
5840            cmdSave_Click  ' ** Procedure: Above.
5850          End If
5860        End If  ' ** vbNullString.
5870      Else
5880        If IsNull(varOldValue) = False Then
5890          .Contact_Name_First = Null
5900          .Contact_Name_Last = Null
5910          .Contact_Name_LastFirst = Null
5920          .Contact_Changed = True
5930          cmdSave_Click  ' ** Procedure: Above.
5940        End If
5950      End If  ' ** IsNull.
5960    End With

EXITP:
5970    Exit Sub

ERRH:
5980    Select Case ERR.Number
        Case Else
5990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6000    End Select
6010    Resume EXITP

End Sub

Private Sub Contact_Name_KeyDown(KeyCode As Integer, Shift As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_KeyDown"

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
6190          .Contact_Address1.SetFocus
6200        End With
6210      End Select
6220    End If

        ' ** Shift keys.
6230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6250        With Me
6260          intRetVal = 0
6270          If .CurrentRecord > 1 Then
6280            MoveRec acCmdRecordsGoToPrevious
6290            .Contact_Email.SetFocus
6300          Else
6310            DoCmd.SelectObject acForm, .Parent.Name, False
6320            .Parent.cmdClose.SetFocus
6330          End If
6340        End With
6350      End Select
6360    End If

EXITP:
6370    KeyCode = intRetVal
6380    Exit Sub

ERRH:
6390    intRetVal = 0
6400    Select Case ERR.Number
        Case Else
6410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6420    End Select
6430    Resume EXITP

End Sub

Private Sub Contact_Name_Exit(Cancel As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_Exit"

6510    varOldValue = Null

EXITP:
6520    Exit Sub

ERRH:
6530    Select Case ERR.Number
        Case Else
6540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6550    End Select
6560    Resume EXITP

End Sub

Private Sub Contact_Name_LastFirst_Enter()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_LastFirst_Enter"

6610    With Me
6620      varOldValue = .Contact_Name_LastFirst
6630    End With

EXITP:
6640    Exit Sub

ERRH:
6650    Select Case ERR.Number
        Case Else
6660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6670    End Select
6680    Resume EXITP

End Sub

Private Sub Contact_Name_LastFirst_GotFocus()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_LastFirst_GotFocus"

6710    With Me
6720      .Contact_Name_LastFirst.SelLength = 0
6730      .Contact_Name_LastFirst.SelStart = 99
6740    End With

EXITP:
6750    Exit Sub

ERRH:
6760    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
6770    Case Else
6780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6790    End Select
6800    Resume EXITP

End Sub

Private Sub Contact_Name_LastFirst_BeforeUpdate(Cancel As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_LastFirst_BeforeUpdate"

        Dim msgResponse As VbMsgBoxResult
        Dim intPos01 As Integer

        'CHECK WHETHER THIS HAS TO USE .TEXT!
6910    With Me
6920      blnAddComma = False
6930      If IsNull(.Contact_Name_LastFirst) = False Then
6940        If Trim(.Contact_Name_LastFirst) <> vbNullString Then
6950          intPos01 = InStr(.Contact_Name_LastFirst, ",")
6960          If intPos01 = 0 Then
6970            msgResponse = MsgBox("There is no comma in this Contact Name." & vbCrLf & _
                  "Should this contact's name not be split 'Last, First'?" & vbCrLf & vbCrLf & _
                  "Click 'Yes' to continue with no comma." & vbCrLf & _
                  "Click 'No' to have a comma automatically added." & vbCrLf & _
                  "Click 'Cancel' to make changes before continuing.", _
                  vbInformation + vbYesNoCancel, "No Comma In Last, First")
6980            Select Case msgResponse
                Case vbYes
                  ' ** Continue as-is.
6990            Case vbNo
7000              blnAddComma = True
7010            Case Else
7020              Cancel = -1
7030            End Select
7040          End If
7050        End If
7060      End If
7070    End With

EXITP:
7080    Exit Sub

ERRH:
7090    Select Case ERR.Number
        Case Else
7100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7110    End Select
7120    Resume EXITP

End Sub

Private Sub Contact_Name_LastFirst_AfterUpdate()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_LastFirst_AfterUpdate"

        Dim varLast As Variant, varFirst As Variant, varSfx As Variant, varLastFirst As Variant, varName As Variant
        Dim intPos01 As Integer
        Dim strTmp01 As String

7210    With Me
7220      If IsNull(.Contact_Name_LastFirst) = False Then
7230        If Trim(.Contact_Name_LastFirst) <> vbNullString Then
7240          strTmp01 = Trim(FixQuotes(.Contact_Name_LastFirst))  ' ** Module Function: modStringFuncs.
7250          If Compare_StringA_StringB(.Contact_Name_LastFirst, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
7260            .Contact_Name_LastFirst = strTmp01
7270          End If
7280          .Contact_Changed = True

7290          cmdSave_Click  ' ** Procedure: Above.

7300          varLast = Null: varFirst = Null: varSfx = Null: varLastFirst = Null: varName = Null

7310          intPos01 = InStr(strTmp01, ",")
7320          If intPos01 > 0 Then
7330            If intPos01 > 1 Then
7340              varLast = Trim(Left(strTmp01, (intPos01 - 1)))
7350              If Len(strTmp01) > intPos01 Then
7360                varFirst = Trim(Mid(strTmp01, (intPos01 + 1)))
7370                intPos01 = InStr(varFirst, ",")
7380                If intPos01 > 0 Then
                      ' ** varLast stands.
7390                  varSfx = Trim(Mid(varFirst, intPos01))
7400                  varFirst = Trim(Left(varFirst, (intPos01 - 1)))
7410                  varName = varFirst & " " & varLast & varSfx
7420                Else
                      ' ** Check for other suffixes.
7430                  If Right(varFirst, 3) = " Jr" Or Right(varFirst, 3) = " Sr" Or _
                          Right(varFirst, 3) = " II" Then
7440                    varSfx = Right(varFirst, 2)
7450                    varFirst = Trim(Left(varFirst, (Len(varFirst) - 2)))
7460                    varName = varFirst & " " & varLast & " " & varSfx
7470                  ElseIf Right(varFirst, 4) = " Jr." Or Right(varFirst, 4) = " Sr." Or _
                          Right(varFirst, 4) = " III" Or Right(varFirst, 4) = " Esq" Then
7480                    varSfx = Right(varFirst, 3)
7490                    varFirst = Trim(Left(varFirst, (Len(varFirst) - 3)))
7500                    varName = varFirst & " " & varLast & " " & varSfx
7510                  ElseIf Right(varFirst, 5) = " Esq." Then
7520                    varSfx = Right(varFirst, 4)
7530                    varFirst = Trim(Left(varFirst, (Len(varFirst) - 4)))
7540                    varName = varFirst & " " & varLast & " " & varSfx
7550                  Else
7560                    varName = varFirst & " " & varLast
7570                  End If
7580                End If
7590              Else
                    ' ** Comma at the end.
7600                .Contact_Name_LastFirst = varLast
7610                varFirst = Null
7620                varName = varLast
7630              End If
7640            Else
                  ' ** Comma at the beginning.
7650              If Len(strTmp01) > 1 Then
7660                strTmp01 = Trim(Mid(strTmp01, 2))
7670                intPos01 = InStr(strTmp01, ",")
7680                If intPos01 > 0 Then
                      ' ** What is this?
7690                  If intPos01 = 1 Or intPos01 = Len(strTmp01) Then
                        ' ** Oh, I give up!
7700                    .Contact_Name_LastFirst = Null
7710                    varLast = Null
7720                    varFirst = Null
7730                    varName = Null
7740                  Else
7750                    varLast = Trim(Left(strTmp01, (intPos01 - 1)))
7760                    varFirst = Trim(Mid(strTmp01, (intPos01 + 1)))
7770                    varName = varFirst & " " & varLast
7780                  End If
7790                Else
                      ' ** No comma.
7800                  .Contact_Name_LastFirst = strTmp01
7810                  varLast = strTmp01
7820                  varFirst = Null
7830                  varName = strTmp01
7840                End If
7850              Else
                    ' ** Only a comma.
7860                .Contact_Name_LastFirst = Null
7870                varFirst = Null
7880                varLast = Null
7890                varName = Null
7900              End If
7910            End If
7920          Else
7930            If blnAddComma = True Then
                  ' ** These should be the same expressions as in the queries.
7940              varLast = IIf(InStr(strTmp01, "Regional Center") > 0 Or InStr(strTmp01, "(") > 0 Or _
                    InStr(strTmp01, "Cnty") > 0, strTmp01, IIf(InStr(GetLastWord(strTmp01, True), "~") > 0, _
                    Left(GetLastWord(strTmp01, True), (InStr(GetLastWord(strTmp01, True), "~") - 1)), _
                    GetLastWord(strTmp01, True)))  ' ** Module Function: modStringFuncs.
7950              varSfx = IIf(InStr(strTmp01, "Regional Center") > 0 Or InStr(strTmp01, "(") > 0 Or _
                    InStr(strTmp01, "Cnty") > 0, Null, IIf(InStr(GetLastWord(strTmp01, True), "~") > 0, _
                    Mid(GetLastWord(strTmp01, True), (InStr(GetLastWord(strTmp01, True), "~") + 1)), Null))  ' ** Module Function: modStringFuncs.
7960              varFirst = IIf(IsNull(strTmp01) = True, Null, IIf(varLast = strTmp01, Null, _
                    Trim(Left(strTmp01, (InStr(strTmp01, varLast) - 1))) & IIf(IsNull(varSfx) = True, "", _
                    IIf(Left(varSfx, 1) = ",", varSfx, " " & varSfx))))
7970              varLastFirst = IIf(IsNull(strTmp01) = True, Null, varLast & IIf(IsNull(varFirst) = True, "", ", " & varFirst))
7980              .Contact_Name_LastFirst = varLastFirst
7990              varName = strTmp01
8000            Else
                  ' ** No comma.
8010              varLast = strTmp01
8020              varFirst = Null
8030              varName = strTmp01
8040            End If
8050          End If
8060          .Contact_Name = varName  ' ** Any or all of these may be null.
8070          .Contact_Name_Last = varLast
8080          .Contact_Name_First = varFirst
8090          cmdSave_Click  ' ** Procedure: Above.
8100        Else
8110          If IsNull(varOldValue) = False Then
8120            .Contact_Name_LastFirst = Null
8130            .Contact_Name = Null
8140            .Contact_Name_Last = Null
8150            .Contact_Name_First = Null
8160            .Contact_Changed = True
8170            cmdSave_Click  ' ** Procedure: Above.
8180          End If
8190        End If  ' ** vbNullString.
8200      Else
8210        If IsNull(varOldValue) = False Then
8220          .Contact_Name = Null
8230          .Contact_Name_Last = Null
8240          .Contact_Name_First = Null
8250          .Contact_Changed = True
8260          cmdSave_Click  ' ** Procedure: Above.
8270        End If
8280      End If  ' ** IsNull.
8290    End With

EXITP:
8300    Exit Sub

ERRH:
8310    Select Case ERR.Number
        Case Else
8320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8330    End Select
8340    Resume EXITP

End Sub

Private Sub Contact_Name_LastFirst_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_LastFirst_KeyDown"

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
8490          .Contact_Address1.SetFocus
8500        End With
8510      End Select
8520    End If

        ' ** Shift keys.
8530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8550        With Me
8560          intRetVal = 0
8570          If .CurrentRecord > 1 Then
8580            MoveRec acCmdRecordsGoToPrevious
8590            .Contact_Email.SetFocus
8600          Else
8610            DoCmd.SelectObject acForm, .Parent.Name, False
8620            .Parent.cmdClose.SetFocus
8630          End If
8640        End With
8650      End Select
8660    End If

EXITP:
8670    KeyCode = intRetVal
8680    Exit Sub

ERRH:
8690    intRetVal = 0
8700    Select Case ERR.Number
        Case Else
8710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8720    End Select
8730    Resume EXITP

End Sub

Private Sub Contact_Name_LastFirst_Exit(Cancel As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_LastFirst_Exit"

8810    varOldValue = Null

EXITP:
8820    Exit Sub

ERRH:
8830    Select Case ERR.Number
        Case Else
8840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8850    End Select
8860    Resume EXITP

End Sub

Private Sub Contact_Address1_lbl_cmd_Address1_DblClick(Cancel As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_lbl_cmd_Address1_DblClick"

8910    SortNow "Contact_Address1_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
8920    Exit Sub

ERRH:
8930    Select Case ERR.Number
        Case Else
8940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8950    End Select
8960    Resume EXITP

End Sub

Private Sub Contact_Address1_Enter()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_Enter"

9010    With Me
9020      varOldValue = .Contact_Address1
9030    End With

EXITP:
9040    Exit Sub

ERRH:
9050    Select Case ERR.Number
        Case Else
9060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9070    End Select
9080    Resume EXITP

End Sub

Private Sub Contact_Address1_GotFocus()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_GotFocus"

9110    With Me
9120      .Contact_Address1.SelLength = 0
9130      .Contact_Address1.SelStart = 99
9140    End With

EXITP:
9150    Exit Sub

ERRH:
9160    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
9170    Case Else
9180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9190    End Select
9200    Resume EXITP

End Sub

Private Sub Contact_Address1_AfterUpdate()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_AfterUpdate"

        Dim strTmp01 As String

9310    With Me
9320      If IsNull(.Contact_Address1) = False Then
9330        If Trim(.Contact_Address1) <> vbNullString Then
9340          strTmp01 = Trim(FixQuotes(.Contact_Address1))  ' ** Module Function: modStringFuncs.
9350          If Compare_StringA_StringB(.Contact_Address1, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
9360            .Contact_Address1 = strTmp01
9370            .Contact_Changed = True
9380            cmdSave_Click  ' ** Procedure: Above.
9390          End If
9400        Else
9410          If IsNull(varOldValue) = False Then
9420            .Contact_Address1 = Null
9430            .Contact_Changed = True
9440            cmdSave_Click  ' ** Procedure: Above.
9450          End If
9460        End If
9470      Else
9480        If IsNull(varOldValue) = False Then
9490          .Contact_Changed = True
9500          cmdSave_Click  ' ** Procedure: Above.
9510        End If
9520      End If
9530    End With

EXITP:
9540    Exit Sub

ERRH:
9550    Select Case ERR.Number
        Case Else
9560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9570    End Select
9580    Resume EXITP

End Sub

Private Sub Contact_Address1_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9670        With Me
9680          intRetVal = 0
9690          .Contact_Address2.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9750        With Me
9760          intRetVal = 0
9770          If .Contact_Name.Visible = True And .Contact_Name.Enabled = True Then
9780            .Contact_Name.SetFocus
9790          ElseIf .Contact_Name_LastFirst.Visible = True And .Contact_Name_LastFirst.Enabled = True Then
9800            .Contact_Name_LastFirst.SetFocus
9810          End If
9820        End With
9830      End Select
9840    End If

EXITP:
9850    KeyCode = intRetVal
9860    Exit Sub

ERRH:
9870    intRetVal = 0
9880    Select Case ERR.Number
        Case Else
9890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9900    End Select
9910    Resume EXITP

End Sub

Private Sub Contact_Address1_Exit(Cancel As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_Exit"

10010   varOldValue = Null

EXITP:
10020   Exit Sub

ERRH:
10030   Select Case ERR.Number
        Case Else
10040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10050   End Select
10060   Resume EXITP

End Sub

Private Sub Contact_Address1_lbl_cmd_Address2_DblClick(Cancel As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_lbl_cmd_Address2_DblClick"

10110   SortNow "Contact_Address2_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
10120   Exit Sub

ERRH:
10130   Select Case ERR.Number
        Case Else
10140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10150   End Select
10160   Resume EXITP

End Sub

Private Sub Contact_Address2_Enter()

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_Enter"

10210   With Me
10220     varOldValue = .Contact_Address2
10230   End With

EXITP:
10240   Exit Sub

ERRH:
10250   Select Case ERR.Number
        Case Else
10260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10270   End Select
10280   Resume EXITP

End Sub

Private Sub Contact_Address2_GotFocus()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_GotFocus"

10310   With Me
10320     .Contact_Address2.SelLength = 0
10330     .Contact_Address2.SelStart = 99
10340   End With

EXITP:
10350   Exit Sub

ERRH:
10360   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
10370   Case Else
10380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10390   End Select
10400   Resume EXITP

End Sub

Private Sub Contact_Address2_AfterUpdate()

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_AfterUpdate"

        Dim strTmp01 As String

10510   With Me
10520     If IsNull(.Contact_Address2) = False Then
10530       If Trim(.Contact_Address2) <> vbNullString Then
10540         strTmp01 = Trim(FixQuotes(.Contact_Address2))  ' ** Module Function: modStringFuncs.
10550         If Compare_StringA_StringB(.Contact_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
10560           .Contact_Address2 = strTmp01
10570           .Contact_Changed = True
10580           cmdSave_Click  ' ** Procedure: Above.
10590         End If
10600       Else
10610         If IsNull(varOldValue) = False Then
10620           .Contact_Address2 = Null
10630           .Contact_Changed = True
10640           cmdSave_Click  ' ** Procedure: Above.
10650         End If
10660       End If
10670     Else
10680       If IsNull(varOldValue) = False Then
10690         .Contact_Changed = True
10700         cmdSave_Click  ' ** Procedure: Above.
10710       End If
10720     End If
10730   End With

EXITP:
10740   Exit Sub

ERRH:
10750   Select Case ERR.Number
        Case Else
10760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10770   End Select
10780   Resume EXITP

End Sub

Private Sub Contact_Address2_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10820   intShiftDown = (Shift And acShiftMask) > 0
10830   intAltDown = (Shift And acAltMask) > 0
10840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         intRetVal = 0
10890         If .opgLocContact.Visible = True Then
10900           .opgLocContact.SetFocus
10910         Else
10920           .Contact_City.SetFocus
10930         End If
10940       End With
10950     End Select
10960   End If

        ' ** Shift keys.
10970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10990       With Me
11000         intRetVal = 0
11010         .Contact_Address1.SetFocus
11020       End With
11030     End Select
11040   End If

EXITP:
11050   KeyCode = intRetVal
11060   Exit Sub

ERRH:
11070   intRetVal = 0
11080   Select Case ERR.Number
        Case Else
11090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11100   End Select
11110   Resume EXITP

End Sub

Private Sub Contact_Address2_Exit(Cancel As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_Exit"

11210   varOldValue = Null

EXITP:
11220   Exit Sub

ERRH:
11230   Select Case ERR.Number
        Case Else
11240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11250   End Select
11260   Resume EXITP

End Sub

Private Sub opgLocContact_lbl_DblClick(Cancel As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_lbl_DblClick"

11310   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
11320   Exit Sub

ERRH:
11330   Select Case ERR.Number
        Case Else
11340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11350   End Select
11360   Resume EXITP

End Sub

Private Sub opgLocContact_Enter()

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_Enter"

11410   With Me
11420     varOldValue = .opgLocContact
11430   End With

EXITP:
11440   Exit Sub

ERRH:
11450   Select Case ERR.Number
        Case Else
11460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11470   End Select
11480   Resume EXITP

End Sub

Public Sub opgLocContact_AfterUpdate()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_AfterUpdate"

11510   With Me
11520     If .opgLocContact <> varOldValue Then
11530       Select Case .opgLocContact
            Case .opgLocContact_optUSA.OptionValue
11540         .opgLocContact_optUSA_lbl.FontBold = True
11550         .opgLocContact_optOther_lbl.FontBold = False
11560         If IsNull(.Contact_Country) = False Then
11570           .Contact_Country = Null
11580         End If
11590         If IsNull(.Contact_PostalCode) = False Then
11600           .Contact_PostalCode = Null
11610         End If
11620         varOldValue = .opgLocContact
11630         .Contact_Country.Enabled = False
11640         .Contact_Country.Locked = True
11650         .Contact_PostalCode.Enabled = False
11660         .Contact_PostalCode.Locked = True
11670         .Contact_State.Enabled = True
11680         .Contact_State.Locked = False
11690         .Contact_Zip.Enabled = True
11700         .Contact_Zip.Locked = False
11710         .Contact_Zip_Format.Enabled = True
11720         .Contact_Zip_Format.Locked = False
11730       Case .opgLocContact_optOther.OptionValue
11740         .opgLocContact_optUSA_lbl.FontBold = False
11750         .opgLocContact_optOther_lbl.FontBold = True
11760         If IsNull(.Contact_State) = False Then
11770           .Contact_State = Null
11780         End If
11790         If IsNull(.Contact_Zip) = False Then
11800           .Contact_Zip = Null
11810         End If
11820         If IsNull(.Contact_Zip_Format) = False Then
11830           .Contact_Zip_Format = Null
11840         End If
11850         varOldValue = .opgLocContact
11860         .Contact_Country.Enabled = True
11870         .Contact_Country.Locked = False
11880         .Contact_PostalCode.Enabled = True
11890         .Contact_PostalCode.Locked = False
11900         .Contact_State.Enabled = False
11910         .Contact_State.Locked = True
11920         .Contact_Zip.Enabled = False
11930         .Contact_Zip.Locked = True
11940         .Contact_Zip_Format.Enabled = False
11950         .Contact_Zip_Format.Locked = True
11960       End Select
11970       .Contact_Changed = True
11980       cmdSave_Click  ' ** Procedure: Above.
11990     End If
12000   End With

EXITP:
12010   Exit Sub

ERRH:
12020   Select Case ERR.Number
        Case Else
12030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12040   End Select
12050   Resume EXITP

End Sub

Private Sub opgLocContact_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12120   intShiftDown = (Shift And acShiftMask) > 0
12130   intAltDown = (Shift And acAltMask) > 0
12140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12170       With Me
12180         intRetVal = 0
12190         .Contact_City.SetFocus
12200       End With
12210     End Select
12220   End If

        ' ** Shift keys.
12230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12250       With Me
12260         intRetVal = 0
12270         .Contact_Address2.SetFocus
12280       End With
12290     End Select
12300   End If

EXITP:
12310   KeyCode = intRetVal
12320   Exit Sub

ERRH:
12330   intRetVal = 0
12340   Select Case ERR.Number
        Case Else
12350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12360   End Select
12370   Resume EXITP

End Sub

Private Sub opgLocContact_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12420   intShiftDown = (Shift And acShiftMask) > 0
12430   intAltDown = (Shift And acAltMask) > 0
12440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12470       With Me
12480         intRetVal = 0
12490         .Contact_City.SetFocus
12500       End With
12510     End Select
12520   End If

        ' ** Shift keys.
12530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12550       With Me
12560         intRetVal = 0
12570         .Contact_Address2.SetFocus
12580       End With
12590     End Select
12600   End If

EXITP:
12610   KeyCode = intRetVal
12620   Exit Sub

ERRH:
12630   intRetVal = 0
12640   Select Case ERR.Number
        Case Else
12650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12660   End Select
12670   Resume EXITP

End Sub

Private Sub opgLocContact_Exit(Cancel As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_Exit"

12710   varOldValue = Null

EXITP:
12720   Exit Sub

ERRH:
12730   Select Case ERR.Number
        Case Else
12740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12750   End Select
12760   Resume EXITP

End Sub

Private Sub Contact_City_lbl_cmd_City_DblClick(Cancel As Integer)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_lbl_cmd_City_DblClick"

12810   SortNow "Contact_City_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
12820   Exit Sub

ERRH:
12830   Select Case ERR.Number
        Case Else
12840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12850   End Select
12860   Resume EXITP

End Sub

Private Sub Contact_City_Enter()

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_Enter"

12910   With Me
12920     varOldValue = .Contact_City
12930   End With

EXITP:
12940   Exit Sub

ERRH:
12950   Select Case ERR.Number
        Case Else
12960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12970   End Select
12980   Resume EXITP

End Sub

Private Sub Contact_City_GotFocus()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_GotFocus"

13010   With Me
13020     .Contact_City.SelLength = 0
13030     .Contact_City.SelStart = 99
13040   End With

EXITP:
13050   Exit Sub

ERRH:
13060   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
13070   Case Else
13080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13090   End Select
13100   Resume EXITP

End Sub

Private Sub Contact_City_AfterUpdate()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_AfterUpdate"

        Dim strTmp01 As String

13210   With Me
13220     If IsNull(.Contact_City) = False Then
13230       If Trim(.Contact_City) <> vbNullString Then
13240         strTmp01 = Trim(FixQuotes(.Contact_City))  ' ** Module Function: modStringFuncs.
13250         If Compare_StringA_StringB(.Contact_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
13260           .Contact_City = strTmp01
13270           .Contact_Changed = True
13280           cmdSave_Click  ' ** Procedure: Above.
13290         End If
13300       Else
13310         If IsNull(varOldValue) = False Then
13320           .Contact_City = Null
13330           .Contact_Changed = True
13340           cmdSave_Click  ' ** Procedure: Above.
13350         End If
13360       End If
13370     Else
13380       If IsNull(varOldValue) = False Then
13390         .Contact_Changed = True
13400         cmdSave_Click  ' ** Procedure: Above.
13410       End If
13420     End If
13430   End With

EXITP:
13440   Exit Sub

ERRH:
13450   Select Case ERR.Number
        Case Else
13460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13470   End Select
13480   Resume EXITP

End Sub

Private Sub Contact_City_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13520   intShiftDown = (Shift And acShiftMask) > 0
13530   intAltDown = (Shift And acAltMask) > 0
13540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13570       With Me
13580         intRetVal = 0
13590         If .opgLocContact.Visible = True Then
13600           Select Case .opgLocContact
                Case .opgLocContact_optUSA.OptionValue
13610             .Contact_State.SetFocus
13620           Case .opgLocContact_optOther.OptionValue
13630             .Contact_Country.SetFocus
13640           End Select
13650         Else
13660           .Contact_State.SetFocus
13670         End If
13680       End With
13690     End Select
13700   End If

        ' ** Shift keys.
13710   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13720     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13730       With Me
13740         intRetVal = 0
13750         If .opgLocContact.Visible = True Then
13760           .opgLocContact.SetFocus
13770         Else
13780           .Contact_Address2.SetFocus
13790         End If
13800       End With
13810     End Select
13820   End If

EXITP:
13830   KeyCode = intRetVal
13840   Exit Sub

ERRH:
13850   intRetVal = 0
13860   Select Case ERR.Number
        Case Else
13870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13880   End Select
13890   Resume EXITP

End Sub

Private Sub Contact_City_Exit(Cancel As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_Exit"

13910   varOldValue = Null

EXITP:
13920   Exit Sub

ERRH:
13930   Select Case ERR.Number
        Case Else
13940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13950   End Select
13960   Resume EXITP

End Sub

Private Sub Contact_City_lbl_cmd_State_DblClick(Cancel As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_lbl_cmd_State_DblClick"

14010   SortNow "Contact_State_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
14020   Exit Sub

ERRH:
14030   Select Case ERR.Number
        Case Else
14040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14050   End Select
14060   Resume EXITP

End Sub

Private Sub Contact_State_Enter()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_Enter"

14110   With Me
14120     varOldValue = .Contact_State
14130     If .Contact_State.text = vbNullString Then
14140       ExpandCombo .Contact_State  ' ** Module Procedure: modUtilities.
14150     End If
14160   End With

EXITP:
14170   Exit Sub

ERRH:
14180   Select Case ERR.Number
        Case Else
14190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14200   End Select
14210   Resume EXITP

End Sub

Private Sub Contact_State_GotFocus()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_GotFocus"

14310   With Me
14320     .Contact_State.SelLength = 0
14330     .Contact_State.SelStart = 9
14340   End With

EXITP:
14350   Exit Sub

ERRH:
14360   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
14370   Case Else
14380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14390   End Select
14400   Resume EXITP

End Sub

Private Sub Contact_State_AfterUpdate()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_AfterUpdate"

14510   With Me
14520     If IsNull(.Contact_State) = False Then
14530       If Trim(.Contact_State) <> vbNullString Then
14540         .Contact_State.Value = FormatUpperCase(Nz(.Contact_State, vbNullString))
14550         Select Case IsNull(.Parent.Contact_State_Pref)
              Case True
14560           .Parent.Contact_State_Pref = .Contact_State
14570         Case False
14580           If .Parent.Contact_State_Pref <> .Contact_State Then
14590             .Parent.Contact_State_Pref = .Contact_State
14600           End If
14610         End Select
14620         .Contact_Changed = True
14630         cmdSave_Click  ' ** Procedure: Above.
14640       Else
14650         If IsNull(varOldValue) = False Then
14660           .Contact_State = Null
14670           .Contact_Changed = True
14680           cmdSave_Click  ' ** Procedure: Above.
14690         End If
14700       End If
14710     Else
14720       If IsNull(varOldValue) = False Then
14730         .Contact_Changed = True
14740         cmdSave_Click  ' ** Procedure: Above.
14750       End If
14760     End If
14770   End With

EXITP:
14780   Exit Sub

ERRH:
14790   Select Case ERR.Number
        Case Else
14800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14810   End Select
14820   Resume EXITP

End Sub

Private Sub Contact_State_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14920   intShiftDown = (Shift And acShiftMask) > 0
14930   intAltDown = (Shift And acAltMask) > 0
14940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14970       With Me
14980         intRetVal = 0
14990         If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
15000           .Contact_Zip.SetFocus
15010         ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
15020           .Contact_Zip_Format.SetFocus
15030         End If
15040       End With
15050     End Select
15060   End If

        ' ** Shift keys.
15070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15090       With Me
15100         intRetVal = 0
15110         .Contact_City.SetFocus
15120       End With
15130     End Select
15140   End If

EXITP:
15150   KeyCode = intRetVal
15160   Exit Sub

ERRH:
15170   intRetVal = 0
15180   Select Case ERR.Number
        Case Else
15190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15200   End Select
15210   Resume EXITP

End Sub

Private Sub Contact_State_NotInList(NewData As String, Response As Integer)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_NotInList"

15310   With Me
15320     .Contact_State = Null
15330     MsgBox "Please select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
15340     Response = acDataErrContinue
15350   End With

EXITP:
15360   Exit Sub

ERRH:
15370   Select Case ERR.Number
        Case Else
15380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15390   End Select
15400   Resume EXITP

End Sub

Private Sub Contact_State_Exit(Cancel As Integer)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_Exit"

15510   varOldValue = Null

EXITP:
15520   Exit Sub

ERRH:
15530   Select Case ERR.Number
        Case Else
15540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15550   End Select
15560   Resume EXITP

End Sub

Private Sub Contact_City_lbl_cmd_Zip_DblClick(Cancel As Integer)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_lbl_cmd_Zip_DblClick"

15610   SortNow "Contact_Zip_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
15620   Exit Sub

ERRH:
15630   Select Case ERR.Number
        Case Else
15640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15650   End Select
15660   Resume EXITP

End Sub

Private Sub Contact_Zip_Enter()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Enter"

15710   With Me
15720     varOldValue = .Contact_Zip
15730   End With

EXITP:
15740   Exit Sub

ERRH:
15750   Select Case ERR.Number
        Case Else
15760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15770   End Select
15780   Resume EXITP

End Sub

Private Sub Contact_Zip_GotFocus()

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_GotFocus"

15810   With Me
15820     .Contact_Zip.SelLength = 0
15830     .Contact_Zip.SelStart = 19
15840   End With

EXITP:
15850   Exit Sub

ERRH:
15860   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
15870   Case Else
15880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15890   End Select
15900   Resume EXITP

End Sub

Private Sub Contact_Zip_AfterUpdate()
' ** This field has no InputMask.
' ** Canadian Zip codes are entered in Contact_PostalCode.

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_AfterUpdate"

        Dim varZip As Variant
        Dim intPos01 As Integer
        Dim strTmp01 As String

16010   With Me
16020     If IsNull(.Contact_Zip) = False Then
16030       If Trim(.Contact_Zip) <> vbNullString Then
16040         strTmp01 = Trim(.Contact_Zip)
16050         intPos01 = InStr(strTmp01, "-")
16060         If intPos01 > 0 Then
16070           If Len(strTmp01) > intPos01 Then
16080             strTmp01 = Left(strTmp01, (intPos01 - 1)) & Mid(strTmp01, (intPos01 + 1))
16090             .Contact_Zip = strTmp01
16100           Else
16110             strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
16120             .Contact_Zip = strTmp01
16130           End If
16140           If IsNumeric(strTmp01) = False Or (Len(strTmp01) <> 5 And Len(strTmp01) <> 9) Then
16150             strTmp01 = vbNullString
16160             .Contact_Zip = Null
16170           End If
16180         End If
16190         .Contact_Changed = True

16200         cmdSave_Click  ' ** Procedure: Above.

16210         varZip = Null

16220         If strTmp01 <> vbNullString Then
16230           Select Case Len(strTmp01)
                Case 5
16240             varZip = strTmp01
16250           Case 9
16260             varZip = FormatZip9(strTmp01)  ' ** Module Function: modStringFuncs.
16270           End Select
16280           .Contact_Zip_Format = varZip
16290           .Contact_Changed = True
16300           cmdSave_Click  ' ** Procedure: Above.
16310         Else
16320           If IsNull(varOldValue) = False Then
16330             .Contact_Zip = Null
16340             .Contact_Changed = True
16350             cmdSave_Click  ' ** Procedure: Above.
16360           End If
16370           If IsNull(.Contact_Zip_Format) = False Then
16380             .Contact_Zip_Format = Null
16390             .Contact_Changed = True
16400             cmdSave_Click  ' ** Procedure: Above.
16410           End If
16420         End If
16430       Else
16440         If IsNull(varOldValue) = False Then
16450           .Contact_Zip = Null
16460           .Contact_Changed = True
16470           cmdSave_Click  ' ** Procedure: Above.
16480         End If
16490         If IsNull(.Contact_Zip_Format) = False Then
16500           .Contact_Zip_Format = Null
16510           .Contact_Changed = True
16520           cmdSave_Click  ' ** Procedure: Above.
16530         End If
16540       End If
16550     Else
16560       If IsNull(varOldValue) = False Then
16570         .Contact_Changed = True
16580         cmdSave_Click  ' ** Procedure: Above.
16590       End If
16600       If IsNull(.Contact_Zip_Format) = False Then
16610         .Contact_Zip_Format = Null
16620         .Contact_Changed = True
16630         cmdSave_Click  ' ** Procedure: Above.
16640       End If
16650     End If
16660   End With

EXITP:
16670   Exit Sub

ERRH:
16680   Select Case ERR.Number
        Case Else
16690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16700   End Select
16710   Resume EXITP

End Sub

Private Sub Contact_Zip_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16820   intShiftDown = (Shift And acShiftMask) > 0
16830   intAltDown = (Shift And acAltMask) > 0
16840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16870       With Me
16880         intRetVal = 0
16890         If .Contact_Phone1.Visible = True And .Contact_Phone1.Enabled = True Then
16900           .Contact_Phone1.SetFocus
16910         ElseIf .Contact_Phone1_Format.Visible = True And .Contact_Phone1_Format.Enabled = True Then
16920           .Contact_Phone1_Format.SetFocus
16930         End If
16940       End With
16950     End Select
16960   End If

        ' ** Shift keys.
16970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16990       With Me
17000         intRetVal = 0
17010         .Contact_State.SetFocus
17020       End With
17030     End Select
17040   End If

EXITP:
17050   KeyCode = intRetVal
17060   Exit Sub

ERRH:
17070   intRetVal = 0
17080   Select Case ERR.Number
        Case Else
17090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17100   End Select
17110   Resume EXITP

End Sub

Private Sub Contact_Zip_Exit(Cancel As Integer)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Exit"

17210   varOldValue = Null

EXITP:
17220   Exit Sub

ERRH:
17230   Select Case ERR.Number
        Case Else
17240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17250   End Select
17260   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_Enter()

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_Enter"

17310   With Me
17320     varOldValue = .Contact_Zip_Format
17330   End With

EXITP:
17340   Exit Sub

ERRH:
17350   Select Case ERR.Number
        Case Else
17360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17370   End Select
17380   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_GotFocus()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_GotFocus"

17410   With Me
17420     .Contact_Zip_Format.SelLength = 0
17430     .Contact_Zip_Format.SelStart = 19
17440   End With

EXITP:
17450   Exit Sub

ERRH:
17460   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
17470   Case Else
17480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17490   End Select
17500   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_AfterUpdate()
' ** This field has an InputMask: 00000"-9999";;_
' ** Canadian Zip codes are entered in Contact_PostalCode.

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String

17610   With Me
17620     If IsNull(.Contact_Zip_Format) = False Then
17630       If Trim(.Contact_Zip_Format) <> vbNullString Then
17640         If Right(.Contact_Zip_Format.text, 4) <> "____" And Right(.Contact_Zip_Format.text, 4) <> "    " And _
                  (Not IsNumeric(Mid(.Contact_Zip_Format.text, 7, 1)) Or Not IsNumeric(Mid(.Contact_Zip_Format.text, 8, 1)) Or _
                  Not IsNumeric(Mid(.Contact_Zip_Format.text, 9, 1)) Or Not IsNumeric(Mid(.Contact_Zip_Format.text, 10, 1))) Then
17650           MsgBox "Please enter a valid zip code in the form 00000 or 00000-0000.", vbInformation + vbOKOnly, "Invalid Entry"
17660         Else
17670           .Contact_Changed = True

17680           cmdSave_Click  ' ** Procedure: Above.

17690           strTmp01 = Trim(.Contact_Zip_Format)
17700           Select Case Len(strTmp01)
                Case 5
                  ' ** strTmp01 stands.
17710           Case 10
17720             intPos01 = InStr(strTmp01, "-")
17730             strTmp01 = Left(strTmp01, (intPos01 - 1)) & Mid(strTmp01, (intPos01 + 1))
17740           End Select
17750           .Contact_Zip = strTmp01
17760           cmdSave_Click  ' ** Procedure: Above.
17770         End If
17780       Else
17790         If IsNull(varOldValue) = False Then
17800           .Contact_Zip_Format = Null
17810           .Contact_Changed = True
17820           cmdSave_Click  ' ** Procedure: Above.
17830         End If
17840         If IsNull(.Contact_Zip) = False Then
17850           .Contact_Zip = Null
17860           .Contact_Changed = True
17870           cmdSave_Click  ' ** Procedure: Above.
17880         End If
17890       End If
17900     Else
17910       If IsNull(varOldValue) = False Then
17920         .Contact_Changed = True
17930         cmdSave_Click  ' ** Procedure: Above.
17940       End If
17950       If IsNull(.Contact_Zip) = False Then
17960         .Contact_Zip = Null
17970         .Contact_Changed = True
17980         cmdSave_Click  ' ** Procedure: Above.
17990       End If
18000     End If
18010   End With

EXITP:
18020   Exit Sub

ERRH:
18030   Select Case ERR.Number
        Case Else
18040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18050   End Select
18060   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_KeyDown(KeyCode As Integer, Shift As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18120   intShiftDown = (Shift And acShiftMask) > 0
18130   intAltDown = (Shift And acAltMask) > 0
18140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18170       With Me
18180         intRetVal = 0
18190         If .Contact_Phone1.Visible = True And .Contact_Phone1.Enabled = True Then
18200           .Contact_Phone1.SetFocus
18210         ElseIf .Contact_Phone1_Format.Visible = True And .Contact_Phone1_Format.Enabled = True Then
18220           .Contact_Phone1_Format.SetFocus
18230         End If
18240       End With
18250     End Select
18260   End If

        ' ** Shift keys.
18270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18290       With Me
18300         intRetVal = 0
18310         .Contact_State.SetFocus
18320       End With
18330     End Select
18340   End If

EXITP:
18350   KeyCode = intRetVal
18360   Exit Sub

ERRH:
18370   intRetVal = 0
18380   Select Case ERR.Number
        Case Else
18390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18400   End Select
18410   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_Exit(Cancel As Integer)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_Exit"

18510   varOldValue = Null

EXITP:
18520   Exit Sub

ERRH:
18530   Select Case ERR.Number
        Case Else
18540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18550   End Select
18560   Resume EXITP

End Sub

Private Sub Contact_Country_lbl_cmd_Country_DblClick(Cancel As Integer)

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_lbl_cmd_Country_DblClick"

18610   SortNow "Contact_Country_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
18620   Exit Sub

ERRH:
18630   Select Case ERR.Number
        Case Else
18640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18650   End Select
18660   Resume EXITP

End Sub

Private Sub Contact_Country_Enter()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_Enter"

18710   With Me
18720     varOldValue = .Contact_Country
18730   End With

EXITP:
18740   Exit Sub

ERRH:
18750   Select Case ERR.Number
        Case Else
18760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18770   End Select
18780   Resume EXITP

End Sub

Private Sub Contact_Country_GotFocus()

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_GotFocus"

18810   With Me
18820     .Contact_Country.SelLength = 0
18830     .Contact_Country.SelStart = 30
18840   End With

EXITP:
18850   Exit Sub

ERRH:
18860   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
18870   Case Else
18880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18890   End Select
18900   Resume EXITP

End Sub

Private Sub Contact_Country_AfterUpdate()

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_AfterUpdate"

        Dim strTmp01 As String

19010   With Me
19020     If IsNull(.Contact_Country) = False Then
19030       If Trim(.Contact_Country) <> vbNullString Then
19040         strTmp01 = Trim(FixQuotes(.Contact_Country))  ' ** Module Function: modStringFuncs.
19050         If Compare_StringA_StringB(.Contact_Country, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
19060           .Contact_Country = strTmp01
19070           .Contact_Changed = True
19080           cmdSave_Click  ' ** Procedure: Above.
19090         End If
19100       Else
19110         If IsNull(varOldValue) = False Then
19120           .Contact_Country = Null
19130           .Contact_Changed = True
19140           cmdSave_Click  ' ** Procedure: Above.
19150         End If
19160       End If
19170     Else
19180       If IsNull(varOldValue) = False Then
19190         .Contact_Changed = True
19200         cmdSave_Click  ' ** Procedure: Above.
19210       End If
19220     End If
19230   End With

EXITP:
19240   Exit Sub

ERRH:
19250   Select Case ERR.Number
        Case Else
19260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19270   End Select
19280   Resume EXITP

End Sub

Private Sub Contact_Country_KeyDown(KeyCode As Integer, Shift As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19320   intShiftDown = (Shift And acShiftMask) > 0
19330   intAltDown = (Shift And acAltMask) > 0
19340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19370       With Me
19380         intRetVal = 0
19390         .Contact_PostalCode.SetFocus
19400       End With
19410     End Select
19420   End If

        ' ** Shift keys.
19430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19450       With Me
19460         intRetVal = 0
19470         .Contact_City.SetFocus
19480       End With
19490     End Select
19500   End If

EXITP:
19510   KeyCode = intRetVal
19520   Exit Sub

ERRH:
19530   intRetVal = 0
19540   Select Case ERR.Number
        Case Else
19550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19560   End Select
19570   Resume EXITP

End Sub

Private Sub Contact_Country_Exit(Cancel As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_Exit"

19610   varOldValue = Null

EXITP:
19620   Exit Sub

ERRH:
19630   Select Case ERR.Number
        Case Else
19640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19650   End Select
19660   Resume EXITP

End Sub

Private Sub Contact_Country_lbl_cmd_PostalCode_DblClick(Cancel As Integer)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_lbl_cmd_PostalCode_DblClick"

19710   SortNow "Contact_PostalCode_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
19720   Exit Sub

ERRH:
19730   Select Case ERR.Number
        Case Else
19740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19750   End Select
19760   Resume EXITP

End Sub

Private Sub Contact_PostalCode_Enter()

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_Enter"

19810   With Me
19820     varOldValue = .Contact_PostalCode
19830   End With

EXITP:
19840   Exit Sub

ERRH:
19850   Select Case ERR.Number
        Case Else
19860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19870   End Select
19880   Resume EXITP

End Sub

Private Sub Contact_PostalCode_GotFocus()

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_GotFocus"

19910   With Me
19920     .Contact_PostalCode.SelLength = 0
19930     .Contact_PostalCode.SelStart = 10
19940   End With

EXITP:
19950   Exit Sub

ERRH:
19960   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
19970   Case Else
19980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19990   End Select
20000   Resume EXITP

End Sub

Private Sub Contact_PostalCode_AfterUpdate()

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_AfterUpdate"

        Dim strTmp01 As String

20110   With Me
20120     If IsNull(.Contact_PostalCode) = False Then
20130       If Trim(.Contact_PostalCode) <> vbNullString Then
20140         strTmp01 = Trim(FixQuotes(.Contact_PostalCode))  ' ** Module Function: modStringFuncs.
20150         If Compare_StringA_StringB(.Contact_PostalCode, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
20160           .Contact_PostalCode = strTmp01
20170           .Contact_Changed = True
20180           cmdSave_Click  ' ** Procedure: Above.
20190         End If
20200       Else
20210         If IsNull(varOldValue) = False Then
20220           .Contact_PostalCode = Null
20230           .Contact_Changed = True
20240           cmdSave_Click  ' ** Procedure: Above.
20250         End If
20260       End If
20270     Else
20280       If IsNull(varOldValue) = False Then
20290         .Contact_Changed = True
20300         cmdSave_Click  ' ** Procedure: Above.
20310       End If
20320     End If
20330   End With

EXITP:
20340   Exit Sub

ERRH:
20350   Select Case ERR.Number
        Case Else
20360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20370   End Select
20380   Resume EXITP

End Sub

Private Sub Contact_PostalCode_KeyDown(KeyCode As Integer, Shift As Integer)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20420   intShiftDown = (Shift And acShiftMask) > 0
20430   intAltDown = (Shift And acAltMask) > 0
20440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20470       With Me
20480         intRetVal = 0
20490         If .Contact_Phone1.Visible = True And .Contact_Phone1.Enabled = True Then
20500           .Contact_Phone1.SetFocus
20510         ElseIf .Contact_Phone1_Format.Visible = True And .Contact_Phone1_Format.Enabled = True Then
20520           .Contact_Phone1_Format.SetFocus
20530         End If
20540       End With
20550     End Select
20560   End If

        ' ** Shift keys.
20570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20590       With Me
20600         intRetVal = 0
20610         .Contact_Country.SetFocus
20620       End With
20630     End Select
20640   End If

EXITP:
20650   KeyCode = intRetVal
20660   Exit Sub

ERRH:
20670   intRetVal = 0
20680   Select Case ERR.Number
        Case Else
20690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20700   End Select
20710   Resume EXITP

End Sub

Private Sub Contact_PostalCode_Exit(Cancel As Integer)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_Exit"

20810   varOldValue = Null

EXITP:
20820   Exit Sub

ERRH:
20830   Select Case ERR.Number
        Case Else
20840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20850   End Select
20860   Resume EXITP

End Sub

Private Sub Contact_Phone1_lbl_cmd_Phone1_DblClick(Cancel As Integer)

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_lbl_cmd_Phone1_DblClick"

20910   SortNow "Contact_Phone1_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
20920   Exit Sub

ERRH:
20930   Select Case ERR.Number
        Case Else
20940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20950   End Select
20960   Resume EXITP

End Sub

Private Sub Contact_Phone1_Enter()

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Enter"

21010   With Me
21020     varOldValue = .Contact_Phone1
21030   End With

EXITP:
21040   Exit Sub

ERRH:
21050   Select Case ERR.Number
        Case Else
21060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21070   End Select
21080   Resume EXITP

End Sub

Private Sub Contact_Phone1_GotFocus()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_GotFocus"

21110   With Me
21120     .Contact_Phone1.SelLength = 0
21130     .Contact_Phone1.SelStart = 19
21140   End With

EXITP:
21150   Exit Sub

ERRH:
21160   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
21170   Case Else
21180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21190   End Select
21200   Resume EXITP

End Sub

Private Sub Contact_Phone1_AfterUpdate()
' ** This field has no InputMask.

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_AfterUpdate"

        Dim strTmp01 As String

21310   With Me
21320     If IsNull(.Contact_Phone1) = False Then
21330       If Trim(.Contact_Phone1) <> vbNullString Then
21340         strTmp01 = Trim(.Contact_Phone1)
21350         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
21360         If strTmp01 <> .Contact_Phone1 Then
21370           .Contact_Phone1 = strTmp01
21380         End If
21390         .Contact_Changed = True

21400         cmdSave_Click  ' ** Procedure: Above.

21410         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
21420         .Contact_Phone1_Format = strTmp01
21430         cmdSave_Click  ' ** Procedure: Above.
21440       Else
21450         If IsNull(varOldValue) = False Then
21460           .Contact_Phone1 = Null
21470           .Contact_Changed = True
21480           cmdSave_Click  ' ** Procedure: Above.
21490         End If
21500         If IsNull(.Contact_Phone1_Format) = False Then
21510           .Contact_Phone1_Format = Null
21520           .Contact_Changed = True
21530           cmdSave_Click  ' ** Procedure: Above.
21540         End If
21550       End If
21560     Else
21570       If IsNull(varOldValue) = False Then
21580         .Contact_Changed = True
21590         cmdSave_Click  ' ** Procedure: Above.
21600       End If
21610       If IsNull(.Contact_Phone1_Format) = False Then
21620         .Contact_Phone1_Format = Null
21630         .Contact_Changed = True
21640         cmdSave_Click  ' ** Procedure: Above.
21650       End If
21660     End If
21670   End With

EXITP:
21680   Exit Sub

ERRH:
21690   Select Case ERR.Number
        Case Else
21700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21710   End Select
21720   Resume EXITP

End Sub

Private Sub Contact_Phone1_KeyDown(KeyCode As Integer, Shift As Integer)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21820   intShiftDown = (Shift And acShiftMask) > 0
21830   intAltDown = (Shift And acAltMask) > 0
21840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21870       With Me
21880         intRetVal = 0
21890         .Contact_Phone2.SetFocus
21900       End With
21910     End Select
21920   End If

        ' ** Shift keys.
21930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
21950       With Me
21960         intRetVal = 0
21970         If .opgLocContact.Visible = True Then
21980           Select Case .opgLocContact
                Case .opgLocContact_optUSA.OptionValue
21990             If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
22000               .Contact_Zip.SetFocus
22010             ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
22020               .Contact_Zip_Format.SetFocus
22030             End If
22040           Case .opgLocContact_optOther.OptionValue
22050             .Contact_PostalCode.SetFocus
22060           End Select
22070         Else
22080           If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
22090             .Contact_Zip.SetFocus
22100           ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
22110             .Contact_Zip_Format.SetFocus
22120           End If
22130         End If
22140       End With
22150     End Select
22160   End If

EXITP:
22170   KeyCode = intRetVal
22180   Exit Sub

ERRH:
22190   intRetVal = 0
22200   Select Case ERR.Number
        Case Else
22210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22220   End Select
22230   Resume EXITP

End Sub

Private Sub Contact_Phone1_Exit(Cancel As Integer)

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Exit"

22310   varOldValue = Null

EXITP:
22320   Exit Sub

ERRH:
22330   Select Case ERR.Number
        Case Else
22340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22350   End Select
22360   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_Enter()

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_Enter"

22410   With Me
22420     varOldValue = .Contact_Phone1_Format
22430   End With

EXITP:
22440   Exit Sub

ERRH:
22450   Select Case ERR.Number
        Case Else
22460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22470   End Select
22480   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_GotFocus()

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_GotFocus"

22510   With Me
22520     .Contact_Phone1_Format.SelLength = 0
22530     .Contact_Phone1_Format.SelStart = 19
22540   End With

EXITP:
22550   Exit Sub

ERRH:
22560   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
22570   Case Else
22580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22590   End Select
22600   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_AfterUpdate()
' ** This field has an InputMask: !"(000) 000-0000";;_
'FOREIGN PHONE NUMS NEED COUNTRY CODE, ETC.!

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_AfterUpdate"

        Dim strTmp01 As String

22710   With Me
22720     If IsNull(.Contact_Phone1_Format) = False Then
22730       If Trim(.Contact_Phone1_Format) <> vbNullString Then
22740         strTmp01 = Trim(.Contact_Phone1)
22750         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
22760         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
22770         If strTmp01 <> .Contact_Phone1_Format Then
22780           .Contact_Phone1_Format = strTmp01
22790         End If
22800         .Contact_Changed = True

22810         cmdSave_Click  ' ** Procedure: Above.

22820         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
22830         .Contact_Phone1 = strTmp01
22840         cmdSave_Click  ' ** Procedure: Above.
22850       Else
22860         If IsNull(varOldValue) = False Then
22870           .Contact_Phone1_Format = Null
22880           .Contact_Changed = True
22890           cmdSave_Click  ' ** Procedure: Above.
22900         End If
22910         If IsNull(.Contact_Phone1) = False Then
22920           .Contact_Phone1 = Null
22930           .Contact_Changed = True
22940           cmdSave_Click  ' ** Procedure: Above.
22950         End If
22960       End If
22970     Else
22980       If IsNull(varOldValue) = False Then
22990         .Contact_Changed = True
23000         cmdSave_Click  ' ** Procedure: Above.
23010       End If
23020       If IsNull(.Contact_Phone1) = False Then
23030         .Contact_Phone1 = Null
23040         .Contact_Changed = True
23050         cmdSave_Click  ' ** Procedure: Above.
23060       End If
23070     End If
23080   End With

EXITP:
23090   Exit Sub

ERRH:
23100   Select Case ERR.Number
        Case Else
23110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23120   End Select
23130   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_KeyDown(KeyCode As Integer, Shift As Integer)

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23220   intShiftDown = (Shift And acShiftMask) > 0
23230   intAltDown = (Shift And acAltMask) > 0
23240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23270       With Me
23280         intRetVal = 0
23290         .Contact_Phone2_Format.SetFocus
23300       End With
23310     End Select
23320   End If

        ' ** Shift keys.
23330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
23350       With Me
23360         intRetVal = 0
23370         If .opgLocContact.Visible = True Then
23380           Select Case .opgLocContact
                Case .opgLocContact_optUSA.OptionValue
23390             If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
23400               .Contact_Zip.SetFocus
23410             ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
23420               .Contact_Zip_Format.SetFocus
23430             End If
23440           Case .opgLocContact_optOther.OptionValue
23450             .Contact_PostalCode.SetFocus
23460           End Select
23470         Else
23480           If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
23490             .Contact_Zip.SetFocus
23500           ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
23510             .Contact_Zip_Format.SetFocus
23520           End If
23530         End If
23540       End With
23550     End Select
23560   End If

EXITP:
23570   KeyCode = intRetVal
23580   Exit Sub

ERRH:
23590   intRetVal = 0
23600   Select Case ERR.Number
        Case Else
23610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23620   End Select
23630   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_Exit(Cancel As Integer)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_Exit"

23710   varOldValue = Null

EXITP:
23720   Exit Sub

ERRH:
23730   Select Case ERR.Number
        Case Else
23740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23750   End Select
23760   Resume EXITP

End Sub

Private Sub Contact_Phone1_lbl_cmd_Phone2_DblClick(Cancel As Integer)

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_lbl_cmd_Phone2_DblClick"

23810   SortNow "Contact_Phone2_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
23820   Exit Sub

ERRH:
23830   Select Case ERR.Number
        Case Else
23840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23850   End Select
23860   Resume EXITP

End Sub

Private Sub Contact_Phone2_Enter()

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Enter"

23910   With Me
23920     varOldValue = .Contact_Phone2
23930   End With

EXITP:
23940   Exit Sub

ERRH:
23950   Select Case ERR.Number
        Case Else
23960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23970   End Select
23980   Resume EXITP

End Sub

Private Sub Contact_Phone2_GotFocus()

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_GotFocus"

24010   With Me
24020     .Contact_Phone2.SelLength = 0
24030     .Contact_Phone2.SelStart = 19
24040   End With

EXITP:
24050   Exit Sub

ERRH:
24060   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
24070   Case Else
24080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24090   End Select
24100   Resume EXITP

End Sub

Private Sub Contact_Phone2_AfterUpdate()
' ** This field has no InputMask.

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_AfterUpdate"

        Dim strTmp01 As String

24210   With Me
24220     If IsNull(.Contact_Phone2) = False Then
24230       If Trim(.Contact_Phone2) <> vbNullString Then
24240         strTmp01 = Trim(.Contact_Phone2)
24250         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
24260         If strTmp01 <> .Contact_Phone2 Then
24270           .Contact_Phone2 = strTmp01
24280         End If
24290         .Contact_Changed = True

24300         cmdSave_Click  ' ** Procedure: Above.

24310         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
24320         .Contact_Phone2_Format = strTmp01
24330         cmdSave_Click  ' ** Procedure: Above.
24340       Else
24350         If IsNull(varOldValue) = False Then
24360           .Contact_Phone2 = Null
24370           .Contact_Changed = True
24380           cmdSave_Click  ' ** Procedure: Above.
24390         End If
24400         If IsNull(.Contact_Phone2_Format) = False Then
24410           .Contact_Phone2_Format = Null
24420           .Contact_Changed = True
24430           cmdSave_Click  ' ** Procedure: Above.
24440         End If
24450       End If
24460     Else
24470       If IsNull(varOldValue) = False Then
24480         .Contact_Changed = True
24490         cmdSave_Click  ' ** Procedure: Above.
24500       End If
24510       If IsNull(.Contact_Phone2_Format) = False Then
24520         .Contact_Phone2_Format = Null
24530         .Contact_Changed = True
24540         cmdSave_Click  ' ** Procedure: Above.
24550       End If
24560     End If
24570   End With

EXITP:
24580   Exit Sub

ERRH:
24590   Select Case ERR.Number
        Case Else
24600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24610   End Select
24620   Resume EXITP

End Sub

Private Sub Contact_Phone2_KeyDown(KeyCode As Integer, Shift As Integer)

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24720   intShiftDown = (Shift And acShiftMask) > 0
24730   intAltDown = (Shift And acAltMask) > 0
24740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24770       With Me
24780         intRetVal = 0
24790         .Contact_Email.SetFocus
24800       End With
24810     End Select
24820   End If

        ' ** Shift keys.
24830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24850       With Me
24860         intRetVal = 0
24870         .Contact_Phone1.SetFocus
24880       End With
24890     End Select
24900   End If

EXITP:
24910   KeyCode = intRetVal
24920   Exit Sub

ERRH:
24930   intRetVal = 0
24940   Select Case ERR.Number
        Case Else
24950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24960   End Select
24970   Resume EXITP

End Sub

Private Sub Contact_Phone2_Exit(Cancel As Integer)

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Exit"

25010   varOldValue = Null

EXITP:
25020   Exit Sub

ERRH:
25030   Select Case ERR.Number
        Case Else
25040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25050   End Select
25060   Resume EXITP

End Sub

Private Sub Contact_Phone2_Format_Enter()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Format_Enter"

25110   With Me
25120     varOldValue = .Contact_Phone2_Format
25130   End With

EXITP:
25140   Exit Sub

ERRH:
25150   Select Case ERR.Number
        Case Else
25160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25170   End Select
25180   Resume EXITP

End Sub

Private Sub Contact_Phone2_Format_GotFocus()

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Format_GotFocus"

25210   With Me
25220     .Contact_Phone2_Format.SelLength = 0
25230     .Contact_Phone2_Format.SelStart = 19
25240   End With

EXITP:
25250   Exit Sub

ERRH:
25260   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
25270   Case Else
25280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25290   End Select
25300   Resume EXITP

End Sub

Private Sub Contact_Phone2_Format_AfterUpdate()
' ** This field has an InputMask: !"(000) 000-0000";;_
'FOREIGN PHONE NUMS NEED COUNTRY CODE, ETC.!

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Format_AfterUpdate"

        Dim strTmp01 As String

25410   With Me
25420     If IsNull(.Contact_Phone2_Format) = False Then
25430       If Trim(.Contact_Phone2_Format) <> vbNullString Then
25440         strTmp01 = Trim(.Contact_Phone2)
25450         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
25460         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
25470         If strTmp01 <> .Contact_Phone2_Format Then
25480           .Contact_Phone2_Format = strTmp01
25490         End If
25500         .Contact_Changed = True

25510         cmdSave_Click  ' ** Procedure: Above.

25520         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
25530         .Contact_Phone2 = strTmp01
25540         cmdSave_Click  ' ** Procedure: Above.
25550       Else
25560         If IsNull(varOldValue) = False Then
25570           .Contact_Phone2_Format = Null
25580           .Contact_Changed = True
25590           cmdSave_Click  ' ** Procedure: Above.
25600         End If
25610         If IsNull(.Contact_Phone2) = False Then
25620           .Contact_Phone2 = Null
25630           .Contact_Changed = True
25640           cmdSave_Click  ' ** Procedure: Above.
25650         End If
25660       End If
25670     Else
25680       If IsNull(varOldValue) = False Then
25690         .Contact_Changed = True
25700         cmdSave_Click  ' ** Procedure: Above.
25710       End If
25720       If IsNull(.Contact_Phone2) = False Then
25730         .Contact_Phone2 = Null
25740         .Contact_Changed = True
25750         cmdSave_Click  ' ** Procedure: Above.
25760       End If
25770     End If
25780   End With

EXITP:
25790   Exit Sub

ERRH:
25800   Select Case ERR.Number
        Case Else
25810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25820   End Select
25830   Resume EXITP

End Sub

Private Sub Contact_Phone2_Format_KeyDown(KeyCode As Integer, Shift As Integer)

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Format_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25920   intShiftDown = (Shift And acShiftMask) > 0
25930   intAltDown = (Shift And acAltMask) > 0
25940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25970       With Me
25980         intRetVal = 0
25990         .Contact_Email.SetFocus
26000       End With
26010     End Select
26020   End If

        ' ** Shift keys.
26030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
26050       With Me
26060         intRetVal = 0
26070         .Contact_Phone1_Format.SetFocus
26080       End With
26090     End Select
26100   End If

EXITP:
26110   KeyCode = intRetVal
26120   Exit Sub

ERRH:
26130   intRetVal = 0
26140   Select Case ERR.Number
        Case Else
26150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26160   End Select
26170   Resume EXITP

End Sub

Private Sub Contact_Phone2_Format_Exit(Cancel As Integer)

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Format_Exit"

26210   varOldValue = Null

EXITP:
26220   Exit Sub

ERRH:
26230   Select Case ERR.Number
        Case Else
26240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26250   End Select
26260   Resume EXITP

End Sub

Private Sub Contact_Email_lbl_cmd_Email_DblClick(Cancel As Integer)

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_lbl_cmd_Email_DblClick"

26310   SortNow "Contact_Email_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
26320   Exit Sub

ERRH:
26330   Select Case ERR.Number
        Case Else
26340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26350   End Select
26360   Resume EXITP

End Sub

Private Sub Contact_Email_Enter()

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_Enter"

26410   With Me
26420     varOldValue = .Contact_Email
26430   End With

EXITP:
26440   Exit Sub

ERRH:
26450   Select Case ERR.Number
        Case Else
26460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26470   End Select
26480   Resume EXITP

End Sub

Private Sub Contact_Email_GotFocus()

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_GotFocus"

26510   With Me
26520     .Contact_Email.SelLength = 0
26530     .Contact_Email.SelStart = 99
26540   End With

EXITP:
26550   Exit Sub

ERRH:
26560   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
26570   Case Else
26580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26590   End Select
26600   Resume EXITP

End Sub

Private Sub Contact_Email_AfterUpdate()
' ** Make sure it's a valid email address.

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String

26710   With Me
26720     If IsNull(.Contact_Email) = False Then
26730       If Trim(.Contact_Email) <> vbNullString Then
26740         strTmp01 = Trim(.Contact_Email)
26750         intPos01 = InStr(strTmp01, "@")
26760         If intPos01 > 0 Then
26770           intPos01 = InStr(intPos01, strTmp01, ".")
26780           If intPos01 > 0 Then
26790             .Contact_MailTo = "MailTo:" & strTmp01
26800             .Contact_Changed = True
26810             cmdSave_Click  ' ** Procedure: Above.
26820           Else
26830             Beep
26840             MsgBox "Entry must be a valid email address.", vbInformation + vbOKOnly, "Invalid Entry"
26850             .Contact_Email.SetFocus
26860           End If
26870         Else
26880           Beep
26890           MsgBox "Entry must be a valid email address.", vbInformation + vbOKOnly, "Invalid Entry"
26900           .Contact_Email.SetFocus
26910         End If
26920       Else
26930         If IsNull(varOldValue) = False Then
26940           .Contact_Email = Null
26950           .Contact_MailTo = Null
26960           .Contact_Changed = True
26970           cmdSave_Click  ' ** Procedure: Above.
26980         End If
26990       End If
27000     Else
27010       If IsNull(varOldValue) = False Then
27020         .Contact_Changed = True
27030         cmdSave_Click  ' ** Procedure: Above.
27040       End If
27050     End If
27060   End With

EXITP:
27070   Exit Sub

ERRH:
27080   Select Case ERR.Number
        Case Else
27090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27100   End Select
27110   Resume EXITP

End Sub

Private Sub Contact_Email_KeyDown(KeyCode As Integer, Shift As Integer)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27220   intShiftDown = (Shift And acShiftMask) > 0
27230   intAltDown = (Shift And acAltMask) > 0
27240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27270       With Me
27280         intRetVal = 0
27290         If .Contact_Fax.Visible = True And .Contact_Fax.Enabled = True Then
27300           .Contact_Fax.SetFocus
27310         ElseIf .Contact_Fax_Format.Visible = True And .Contact_Fax_Format.Enabled = True Then
27320           .Contact_Fax_Format.SetFocus
27330         End If
27340       End With
27350     End Select
27360   End If

        ' ** Shift keys.
27370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27390       With Me
27400         intRetVal = 0
27410         If .Contact_Phone2.Visible = True And .Contact_Phone2.Enabled = True Then
27420           .Contact_Phone2.SetFocus
27430         ElseIf .Contact_Phone2_Format.Visible = True And .Contact_Phone2_Format.Enabled = True Then
27440           .Contact_Phone2_Format.SetFocus
27450         End If
27460       End With
27470     End Select
27480   End If

        ' ** Ctrl keys.
27490   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
27500     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
27510       With Me
27520         intRetVal = 0
27530         If .Contact_Email_cmd.Visible = True And .Contact_Email_cmd.Enabled = True Then
27540           .Contact_Email_cmd.SetFocus
27550         End If
27560       End With
27570     End Select
27580   End If

EXITP:
27590   KeyCode = intRetVal
27600   Exit Sub

ERRH:
27610   intRetVal = 0
27620   Select Case ERR.Number
        Case Else
27630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27640   End Select
27650   Resume EXITP

End Sub

Private Sub Contact_Email_Exit(Cancel As Integer)

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_Exit"

27710   varOldValue = Null

EXITP:
27720   Exit Sub

ERRH:
27730   Select Case ERR.Number
        Case Else
27740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27750   End Select
27760   Resume EXITP

End Sub

Public Sub Contact_Email_cmd_Click()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_cmd_Click"

        Dim strTmp01 As String

27810   With Me
27820     If IsNull(.Contact_Email) = False Then
27830       If Trim(.Contact_Email) <> vbNullString Then
27840         DoCmd.Hourglass True
27850         DoEvents
27860         strTmp01 = "MailTo:" & .Contact_Email
27870         If IsNull(.Contact_MailTo) = True Then
27880           .Contact_MailTo = strTmp01
27890           .Contact_Changed = True
27900           cmdSave_Click  ' ** Procedure: Above.
27910         Else
27920           If .Contact_MailTo <> strTmp01 Then
27930             .Contact_MailTo = strTmp01
27940             .Contact_Changed = True
27950             cmdSave_Click  ' ** Procedure: Above.
27960           End If
27970         End If
27980         .Contact_Email_cmd.HyperlinkAddress = .Contact_MailTo
27990         .Contact_Email_cmd.Hyperlink.Follow
28000         .Contact_Email_cmd.HyperlinkAddress = vbNullString
28010       Else
28020         Beep
28030       End If
28040     Else
28050       Beep
28060     End If
28070     DoCmd.Hourglass False
28080   End With

EXITP:
28090   Exit Sub

ERRH:
28100   DoCmd.Hourglass False
28110   Select Case ERR.Number
        Case Else
28120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28130   End Select
28140   Resume EXITP

End Sub

Private Sub Contact_Email_cmd_KeyDown(KeyCode As Integer, Shift As Integer)

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_cmd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28220   intShiftDown = (Shift And acShiftMask) > 0
28230   intAltDown = (Shift And acAltMask) > 0
28240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28270       With Me
28280         intRetVal = 0
28290         If .Contact_Fax.Visible = True And .Contact_Fax.Enabled = True Then
28300           .Contact_Fax.SetFocus
28310         ElseIf .Contact_Fax_Format.Visible = True And .Contact_Fax_Format.Enabled = True Then
28320           .Contact_Fax_Format.SetFocus
28330         End If
28340       End With
28350     End Select
28360   End If

        ' ** Shift keys.
28370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28390       With Me
28400         intRetVal = 0
28410         .Contact_Email.SetFocus
28420       End With
28430     End Select
28440   End If

EXITP:
28450   KeyCode = intRetVal
28460   Exit Sub

ERRH:
28470   intRetVal = 0
28480   Select Case ERR.Number
        Case Else
28490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28500   End Select
28510   Resume EXITP

End Sub

Private Sub Contact_Email_lbl_cmd_Fax_DblClick(Cancel As Integer)

28600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_lbl_cmd_Fax_DblClick"

28610   SortNow "Contact_Fax_lbl_DblClick"  ' ** Procedure: Below.

EXITP:
28620   Exit Sub

ERRH:
28630   Select Case ERR.Number
        Case Else
28640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28650   End Select
28660   Resume EXITP

End Sub

Private Sub Contact_Fax_Enter()

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Enter"

28710   With Me
28720     varOldValue = .Contact_Fax
28730   End With

EXITP:
28740   Exit Sub

ERRH:
28750   Select Case ERR.Number
        Case Else
28760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28770   End Select
28780   Resume EXITP

End Sub

Private Sub Contact_Fax_GotFocus()

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_GotFocus"

28810   With Me
28820     .Contact_Fax.SelLength = 0
28830     .Contact_Fax.SelStart = 19
28840   End With

EXITP:
28850   Exit Sub

ERRH:
28860   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
28870   Case Else
28880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28890   End Select
28900   Resume EXITP

End Sub

Private Sub Contact_Fax_AfterUpdate()
' ** This field has no InputMask.

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_AfterUpdate"

        Dim strTmp01 As String

29010   With Me
29020     If IsNull(.Contact_Fax) = False Then
29030       If Trim(.Contact_Fax) <> vbNullString Then
29040         strTmp01 = Trim(.Contact_Fax)
29050         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
29060         If strTmp01 <> .Contact_Fax Then
29070           .Contact_Fax = strTmp01
29080         End If
29090         .Contact_Changed = True

29100         cmdSave_Click  ' ** Procedure: Above.

29110         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
29120         .Contact_Fax_Format = strTmp01
29130         cmdSave_Click  ' ** Procedure: Above.
29140       Else
29150         If IsNull(varOldValue) = False Then
29160           .Contact_Fax = Null
29170           .Contact_Changed = True
29180           cmdSave_Click  ' ** Procedure: Above.
29190         End If
29200         If IsNull(.Contact_Fax_Format) = False Then
29210           .Contact_Fax_Format = Null
29220           .Contact_Changed = True
29230           cmdSave_Click  ' ** Procedure: Above.
29240         End If
29250       End If
29260     Else
29270       If IsNull(varOldValue) = False Then
29280         .Contact_Changed = True
29290         cmdSave_Click  ' ** Procedure: Above.
29300       End If
29310       If IsNull(.Contact_Fax_Format) = False Then
29320         .Contact_Fax_Format = Null
29330         .Contact_Changed = True
29340         cmdSave_Click  ' ** Procedure: Above.
29350       End If
29360     End If
29370   End With

EXITP:
29380   Exit Sub

ERRH:
29390   Select Case ERR.Number
        Case Else
29400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29410   End Select
29420   Resume EXITP

End Sub

Private Sub Contact_Fax_KeyDown(KeyCode As Integer, Shift As Integer)

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29520   intShiftDown = (Shift And acShiftMask) > 0
29530   intAltDown = (Shift And acAltMask) > 0
29540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29570       With Me
29580         intRetVal = 0
29590         lngRecsCur = RecCnt  ' ** Function: Below.
29600         If .CurrentRecord < lngRecsCur Then
29610           .MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
29620           If .Contact_Name.Visible = True And .Contact_Name.Enabled = True Then
29630             .Contact_Name.SetFocus
29640           ElseIf .Contact_Name_LastFirst.Visible = True And .Contact_Name_LastFirst.Enabled = True Then
29650             .Contact_Name_LastFirst.SetFocus
29660           End If
29670         Else
29680           DoCmd.SelectObject acForm, .Parent.Name, False
29690           .Parent.opgView.SetFocus
29700         End If
29710       End With
29720     End Select
29730   End If

        ' ** Shift keys.
29740   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29750     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29760       With Me
29770         intRetVal = 0
29780         .Contact_Email.SetFocus
29790       End With
29800     End Select
29810   End If

        ' ** Ctrl-Shift keys.
29820   If intCtrlDown And (Not intAltDown) And intShiftDown Then
29830     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29840       With Me
29850         intRetVal = 0
29860         If .Contact_Email_cmd.Visible = True And .Contact_Email_cmd.Enabled = True Then
29870           .Contact_Email_cmd.SetFocus
29880         End If
29890       End With
29900     End Select
29910   End If

EXITP:
29920   KeyCode = intRetVal
29930   Exit Sub

ERRH:
29940   intRetVal = 0
29950   Select Case ERR.Number
        Case Else
29960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29970   End Select
29980   Resume EXITP

End Sub

Private Sub Contact_Fax_Exit(Cancel As Integer)

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Exit"

30010   varOldValue = Null

EXITP:
30020   Exit Sub

ERRH:
30030   Select Case ERR.Number
        Case Else
30040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30050   End Select
30060   Resume EXITP

End Sub

Private Sub Contact_Fax_Format_Enter()

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Format_Enter"

30110   With Me
30120     varOldValue = .Contact_Fax_Format
30130   End With

EXITP:
30140   Exit Sub

ERRH:
30150   Select Case ERR.Number
        Case Else
30160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30170   End Select
30180   Resume EXITP

End Sub

Private Sub Contact_Fax_Format_GotFocus()

30200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Format_GotFocus"

30210   With Me
30220     .Contact_Fax_Format.SelLength = 0
30230     .Contact_Fax_Format.SelStart = 19
30240   End With

EXITP:
30250   Exit Sub

ERRH:
30260   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
30270   Case Else
30280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30290   End Select
30300   Resume EXITP

End Sub

Private Sub Contact_Fax_Format_AfterUpdate()
' ** This field has an InputMask: !"(000) 000-0000";;_
'FOREIGN PHONE NUMS NEED COUNTRY CODE, ETC.!

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Format_AfterUpdate"

        Dim strTmp01 As String

30410   With Me
30420     If IsNull(.Contact_Fax_Format) = False Then
30430       If Trim(.Contact_Fax_Format) <> vbNullString Then
30440         strTmp01 = Trim(.Contact_Fax)
30450         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
30460         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
30470         If strTmp01 <> .Contact_Fax_Format Then
30480           .Contact_Fax_Format = strTmp01
30490         End If
30500         .Contact_Changed = True

30510         cmdSave_Click  ' ** Procedure: Above.

30520         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
30530         .Contact_Fax = strTmp01
30540         cmdSave_Click  ' ** Procedure: Above.
30550       Else
30560         If IsNull(varOldValue) = False Then
30570           .Contact_Fax_Format = Null
30580           .Contact_Changed = True
30590           cmdSave_Click  ' ** Procedure: Above.
30600         End If
30610         If IsNull(.Contact_Fax) = False Then
30620           .Contact_Fax = Null
30630           .Contact_Changed = True
30640           cmdSave_Click  ' ** Procedure: Above.
30650         End If
30660       End If
30670     Else
30680       If IsNull(varOldValue) = False Then
30690         .Contact_Changed = True
30700         cmdSave_Click  ' ** Procedure: Above.
30710       End If
30720       If IsNull(.Contact_Fax) = False Then
30730         .Contact_Fax = Null
30740         .Contact_Changed = True
30750         cmdSave_Click  ' ** Procedure: Above.
30760       End If
30770     End If
30780   End With

EXITP:
30790   Exit Sub

ERRH:
30800   Select Case ERR.Number
        Case Else
30810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30820   End Select
30830   Resume EXITP

End Sub

Private Sub Contact_Fax_Format_KeyDown(KeyCode As Integer, Shift As Integer)

30900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Format_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30920   intShiftDown = (Shift And acShiftMask) > 0
30930   intAltDown = (Shift And acAltMask) > 0
30940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30970       With Me
30980         intRetVal = 0
30990         lngRecsCur = RecCnt  ' ** Function: Below.
31000         If .CurrentRecord < lngRecsCur Then
31010           .MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
31020           If .Contact_Name.Visible = True And .Contact_Name.Enabled = True Then
31030             .Contact_Name.SetFocus
31040           ElseIf .Contact_Name_LastFirst.Visible = True And .Contact_Name_LastFirst.Enabled = True Then
31050             .Contact_Name_LastFirst.SetFocus
31060           End If
31070         Else
31080           DoCmd.SelectObject acForm, .Parent.Name, False
31090           .Parent.opgView.SetFocus
31100         End If
31110       End With
31120     End Select
31130   End If

        ' ** Shift keys.
31140   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31150     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31160       With Me
31170         intRetVal = 0
31180         .Contact_Email.SetFocus
31190       End With
31200     End Select
31210   End If

EXITP:
31220   KeyCode = intRetVal
31230   Exit Sub

ERRH:
31240   intRetVal = 0
31250   Select Case ERR.Number
        Case Else
31260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31270   End Select
31280   Resume EXITP

End Sub

Private Sub Contact_Fax_Format_Exit(Cancel As Integer)

31300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Format_Exit"

31310   varOldValue = Null

EXITP:
31320   Exit Sub

ERRH:
31330   Select Case ERR.Number
        Case Else
31340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
31350   End Select
31360   Resume EXITP

End Sub

Public Sub EnableCountry(blnEnable As Boolean)

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "EnableCountry"

        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim strTmp01 As String

31410   With Me
31420     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
31430       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
31440     End If
31450     .FocusHolder2.SetFocus
31460     blnSortHere = False: blnResort = False
31470     strTmp01 = vbNullString
31480     Select Case blnEnable
          Case True
            ' ** Original configuration, with Country.
31490       If .Contact_Country.Visible = False Then
31500         .Width = lngForm_Width
31510         If .Sort_line.Left = .Contact_Email_lbl_cmd_Email.Left Then
31520           blnSortHere = True
31530           strTmp01 = "Contact_Email"
31540         ElseIf .Sort_line.Left = .Contact_Email_lbl_cmd_Fax.Left Then
31550           blnSortHere = True
31560           strTmp01 = "Contact_Fax"
31570         End If
31580         .Contact_Email.Left = lngEmail_Left
31590         .Contact_Fax.Left = .Contact_Email.Left
31600         .Contact_Fax_Format.Left = .Contact_Email.Left
31610         .Contact_Email_lbl.Left = .Contact_Email.Left
31620         .Contact_Email_lbl_dim_hi.Left = (.Contact_Email_lbl.Left + lngTpp)
31630         .Contact_Email_lbl_line.Left = .Contact_Email_lbl.Left
31640         .Contact_Email_lbl_line_dim_hi.Left = (.Contact_Email_lbl_line.Left + lngTpp)
31650         .Contact_Email_lbl_cmd_Email.Left = .Contact_Email.Left
31660         .Contact_Email_lbl_cmd_Fax.Left = ((.Contact_Email.Left + .Contact_Email.Width) - .Contact_Email_lbl_cmd_Fax.Width)
31670         .Contact_Email_cmd.Left = lngEmailCmd_Left
31680         If blnSortHere = True Then
31690           Select Case strTmp01
                Case "Contact_Email"
31700             .Sort_line.Left = .Contact_Email_lbl_cmd_Email.Left
31710             .Sort_lbl.Left = ((.Contact_Email_lbl_cmd_Email.Left + .Contact_Email_lbl_cmd_Email.Width) - lngSortLbl_Width)
31720           Case "Contact_Fax"
31730             .Sort_line.Left = .Contact_Email_lbl_cmd_Fax.Left
31740             .Sort_lbl.Left = ((.Contact_Email_lbl_cmd_Fax.Left + .Contact_Email_lbl_cmd_Fax.Width) - lngSortLbl_Width)
31750           End Select
31760         End If
31770         blnSortHere = False: strTmp01 = vbNullString
31780         If .Sort_line.Left = .Contact_Phone1_lbl_cmd_Phone1.Left Then
31790           blnSortHere = True
31800           strTmp01 = "Contact_Phone1"
31810         ElseIf .Sort_line.Left = .Contact_Phone1_lbl_cmd_Phone2.Left Then
31820           blnSortHere = True
31830           strTmp01 = "Contact_Phone2"
31840         End If
31850         .Contact_Phone1.Left = lngPhone_Left
31860         .Contact_Phone1_Format.Left = .Contact_Phone1.Left
31870         .Contact_Phone2.Left = .Contact_Phone1.Left
31880         .Contact_Phone2_Format.Left = .Contact_Phone1.Left
31890         .Contact_Phone1_lbl.Left = (.Contact_Phone1.Left - (2& * lngTpp))
31900         .Contact_Phone1_lbl_dim_hi.Left = (.Contact_Phone1_lbl.Left + lngTpp)
31910         .Contact_Phone1_lbl_line.Left = .Contact_Phone1_lbl.Left
31920         .Contact_Phone1_lbl_line_dim_hi.Left = (.Contact_Phone1_lbl_line.Left + lngTpp)
31930         .Contact_Phone1_lbl_cmd_Phone1.Left = .Contact_Phone1.Left
31940         .Contact_Phone1_lbl_cmd_Phone2.Left = ((.Contact_Phone1.Left + .Contact_Phone1.Width) - .Contact_Phone1_lbl_cmd_Phone2.Width)
31950         If blnSortHere = True Then
31960           Select Case strTmp01
                Case "Contact_Phone1"
31970             .Sort_line.Left = .Contact_Phone1_lbl_cmd_Phone1.Left
31980             .Sort_lbl.Left = (((.Contact_Phone1_lbl_cmd_Phone1.Left + .Contact_Phone1_lbl_cmd_Phone1.Width) - _
                    lngSortLbl_Width) - (1& * lngTpp))
31990           Case "Contact_Phone2"
32000             .Sort_line.Left = .Contact_Phone1_lbl_cmd_Phone2.Left
32010             .Sort_lbl.Left = (((.Contact_Phone1_lbl_cmd_Phone2.Left + .Contact_Phone1_lbl_cmd_Phone2.Width) - _
                    lngSortLbl_Width) + (7& * lngTpp))
32020           End Select
32030         End If
32040         blnSortHere = False: strTmp01 = vbNullString
32050         .Contact_Country_lbl.Visible = True
32060         .Contact_Country_lbl_dim_hi.Left = (.Contact_Country_lbl.Left + lngTpp)
32070         .Contact_Country_lbl_line.Visible = True
32080         .Contact_Country_lbl_line_dim_hi.Left = (.Contact_Country_lbl_line.Left + lngTpp)
32090         .Contact_Country_lbl_cmd_Country.Visible = True
32100         .Contact_Country_lbl_cmd_PostalCode.Visible = True
32110         .Contact_Country.Visible = True
32120         .Contact_Country_box.Visible = True
32130         .Contact_PostalCode.Visible = True
32140         .Contact_PostalCode_box.Visible = True
32150         If .Sort_line.Left = .Contact_City_lbl_cmd_City.Left Then
32160           blnSortHere = True
32170           strTmp01 = "Contact_City"
32180         ElseIf .Sort_line.Left = .Contact_City_lbl_cmd_State.Left Then
32190           blnSortHere = True
32200           strTmp01 = "Contact_State"
32210         ElseIf .Sort_line.Left = .Contact_City_lbl_cmd_Zip.Left Then
32220           blnSortHere = True
32230           strTmp01 = "Contact_Zip"
32240         End If
32250         .Contact_City.Left = lngCity_Left
32260         .Contact_City_lbl.Left = .Contact_City.Left
32270         .Contact_City_lbl_dim_hi.Left = (.Contact_City_lbl.Left + lngTpp)
32280         .Contact_City_lbl_line.Left = .Contact_City_lbl.Left
32290         .Contact_City_lbl_line_dim_hi.Left = (.Contact_City_lbl_line.Left + lngTpp)
32300         .Contact_City_lbl_cmd_City.Left = .Contact_City.Left
32310         .Contact_City_lbl_cmd_State.Left = (.Contact_City_lbl_cmd_City.Left + .Contact_City_lbl_cmd_City.Width)
32320         .Contact_City_lbl_cmd_Zip.Left = ((.Contact_City.Left + .Contact_City.Width) - .Contact_City_lbl_cmd_Zip.Width)
32330         .Contact_State.Left = .Contact_City.Left
32340         .Contact_State_box.Left = (.Contact_State.Left - lngTpp)
32350         .Contact_Zip.Left = ((.Contact_City.Left + .Contact_City.Width) - .Contact_Zip.Width)
32360         .Contact_Zip_Format.Left = .Contact_Zip.Left
32370         .Contact_Zip_box.Left = (.Contact_Zip.Left - lngTpp)
32380         If blnSortHere = True Then
32390           Select Case strTmp01
                Case "Contact_City"
32400             .Sort_line.Left = .Contact_City_lbl_cmd_City.Left
32410             .Sort_lbl.Left = ((.Contact_City_lbl_cmd_City.Left + .Contact_City_lbl_cmd_City.Width) - lngSortLbl_Width)
32420           Case "Contact_State"
32430             .Sort_line.Left = .Contact_City_lbl_cmd_State.Left
32440             .Sort_lbl.Left = ((.Contact_City_lbl_cmd_State.Left + .Contact_City_lbl_cmd_State.Width) - lngSortLbl_Width)
32450           Case "Contact_Zip"
32460             .Sort_line.Left = .Contact_City_lbl_cmd_Zip.Left
32470             .Sort_lbl.Left = (((.Contact_City_lbl_cmd_Zip.Left + .Contact_City_lbl_cmd_Zip.Width) - lngSortLbl_Width) + (2& * lngTpp))
32480           End Select
32490         End If
32500         blnSortHere = False: strTmp01 = vbNullString
32510         .opgLocContact_lbl.Visible = True
32520         .opgLocContact_lbl_dim_hi.Left = (.opgLocContact_lbl.Left + lngTpp)
32530         .opgLocContact_lbl2.Visible = True
32540         .opgLocContact_lbl2_dim_hi.Left = (.opgLocContact_lbl2.Left + lngTpp)
32550         .opgLocContact_lbl_line.Visible = True
32560         .opgLocContact_lbl_line_dim_hi.Left = (.opgLocContact_lbl_line.Left + lngTpp)
32570         .opgLocContact.Visible = True
32580         .opgLocContact_box.Visible = True
32590         .opgLocContact_optUSA_lbl_on.Visible = True
32600         .opgLocContact_optUSA_lbl_off.Visible = True
32610         .opgLocContact_optOther_lbl_on.Visible = True
32620         .opgLocContact_optOther_lbl_off.Visible = True
32630       End If
32640     Case False
32650       If .Contact_Country.Visible = True Then
32660         If .Sort_line.Left = .opgLocContact_lbl_line.Left Then blnSortHere = True
32670         .opgLocContact_lbl.Visible = False
32680         .opgLocContact_lbl_dim_hi.Visible = False
32690         .opgLocContact_lbl2.Visible = False
32700         .opgLocContact_lbl2_dim_hi.Visible = False
32710         .opgLocContact_lbl_line.Visible = False
32720         .opgLocContact_lbl_line_dim_hi.Visible = False
32730         .opgLocContact.Visible = False
32740         .opgLocContact_box.Visible = False
32750         .opgLocContact_optUSA_lbl_on.Visible = False
32760         .opgLocContact_optUSA_lbl_off.Visible = False
32770         .opgLocContact_optOther_lbl_on.Visible = False
32780         .opgLocContact_optOther_lbl_off.Visible = False
32790         If blnSortHere = True Then
32800           .Sort_line.Visible = False
32810           .Sort_lbl.Visible = False
32820           blnResort = True
32830         End If
32840         blnSortHere = False: strTmp01 = vbNullString
32850         If .Sort_line.Left = .Contact_Country_lbl_cmd_Country.Left Then
32860           blnSortHere = True
32870           strTmp01 = "Contact_Country"
32880         ElseIf .Sort_line.Left = .Contact_Country_lbl_cmd_PostalCode.Left Then
32890           blnSortHere = True
32900           strTmp01 = "Contact_PostalCode"
32910         End If
32920         .Contact_Country_lbl.Visible = False
32930         .Contact_Country_lbl_dim_hi.Visible = False
32940         .Contact_Country_lbl_line.Visible = False
32950         .Contact_Country_lbl_line_dim_hi.Visible = False
32960         .Contact_Country_lbl_cmd_Country.Visible = False
32970         .Contact_Country_lbl_cmd_PostalCode.Visible = False
32980         .Contact_Country.Visible = False
32990         .Contact_Country_box.Visible = False
33000         .Contact_PostalCode.Visible = False
33010         .Contact_PostalCode_box.Visible = False
33020         If blnSortHere = True Then  ' ** Either one, doesn't matter.
33030           .Sort_line.Visible = False
33040           .Sort_lbl.Visible = False
33050           blnResort = True
33060         End If
33070         blnSortHere = False: strTmp01 = vbNullString
33080         If .Sort_line.Left = .Contact_City_lbl_cmd_City.Left Then
33090           blnSortHere = True
33100           strTmp01 = "Contact_City"
33110         ElseIf .Sort_line.Left = .Contact_City_lbl_cmd_State.Left Then
33120           blnSortHere = True
33130           strTmp01 = "Contact_State"
33140         ElseIf .Sort_line.Left = .Contact_City_lbl_cmd_Zip.Left Then
33150           blnSortHere = True
33160           strTmp01 = "Contact_Zip"
33170         End If
33180         .Contact_City.Left = ((lngCity_Left - lngLocContact_Width) - lngTpp)  'CHECK THIS!
33190         .Contact_City_lbl.Left = .Contact_City.Left
33200         .Contact_City_lbl_dim_hi.Left = (.Contact_City_lbl.Left + lngTpp)
33210         .Contact_City_lbl_line.Left = .Contact_City_lbl.Left
33220         .Contact_City_lbl_line_dim_hi.Left = (.Contact_City_lbl_line.Left + lngTpp)
33230         .Contact_City_lbl_cmd_City.Left = .Contact_City.Left
33240         .Contact_City_lbl_cmd_State.Left = (.Contact_City_lbl_cmd_City.Left + .Contact_City_lbl_cmd_City.Width)
33250         .Contact_City_lbl_cmd_Zip.Left = ((.Contact_City.Left + .Contact_City.Width) - .Contact_City_lbl_cmd_Zip.Width)
33260         .Contact_State.Left = .Contact_City.Left
33270         .Contact_State_box.Left = (.Contact_State.Left - lngTpp)
33280         .Contact_Zip.Left = ((.Contact_City.Left + .Contact_City.Width) - .Contact_Zip.Width)
33290         .Contact_Zip_Format.Left = .Contact_Zip.Left
33300         .Contact_Zip_box.Left = (.Contact_Zip.Left - lngTpp)
33310         If blnSortHere = True Then
33320           Select Case strTmp01
                Case "Contact_City"
33330             .Sort_line.Left = .Contact_City_lbl_cmd_City.Left
33340             .Sort_lbl.Left = ((.Contact_City_lbl_cmd_City.Left + .Contact_City_lbl_cmd_City.Width) - lngSortLbl_Width)
33350           Case "Contact_State"
33360             .Sort_line.Left = .Contact_City_lbl_cmd_State.Left
33370             .Sort_lbl.Left = ((.Contact_City_lbl_cmd_State.Left + .Contact_City_lbl_cmd_State.Width) - lngSortLbl_Width)
33380           Case "Contact_Zip"
33390             .Sort_line.Left = .Contact_City_lbl_cmd_Zip.Left
33400             .Sort_lbl.Left = (((.Contact_City_lbl_cmd_Zip.Left + .Contact_City_lbl_cmd_Zip.Width) - lngSortLbl_Width) + (2& * lngTpp))
33410           End Select
33420         End If
33430         blnSortHere = False: strTmp01 = vbNullString
33440         If .Sort_line.Left = .Contact_Phone1_lbl_cmd_Phone1.Left Then
33450           blnSortHere = True
33460           strTmp01 = "Contact_Phone1"
33470         ElseIf .Sort_line.Left = .Contact_Phone1_lbl_cmd_Phone2.Left Then
33480           blnSortHere = True
33490           strTmp01 = "Contact_Phone2"
33500         End If
33510         .Contact_Phone1.Left = ((lngPhone_Left - (lngLocContact_Width + lngCountry_Width)) - lngTpp)  'CHECK THIS!
33520         .Contact_Phone1_Format.Left = .Contact_Phone1.Left
33530         .Contact_Phone2.Left = .Contact_Phone1.Left
33540         .Contact_Phone2_Format.Left = .Contact_Phone1.Left
33550         .Contact_Phone1_lbl.Left = (.Contact_Phone1.Left - (2& * lngTpp))
33560         .Contact_Phone1_lbl_dim_hi.Left = (.Contact_Phone1_lbl.Left + lngTpp)
33570         .Contact_Phone1_lbl_line.Left = .Contact_Phone1.Left
33580         .Contact_Phone1_lbl_line_dim_hi.Left = (.Contact_Phone1_lbl_line.Left + lngTpp)
33590         .Contact_Phone1_lbl_cmd_Phone1.Left = .Contact_Phone1.Left
33600         .Contact_Phone1_lbl_cmd_Phone2.Left = ((.Contact_Phone1.Left + .Contact_Phone1.Width) - .Contact_Phone1_lbl_cmd_Phone2.Width)
33610         If blnSortHere = True Then
33620           Select Case strTmp01
                Case "Contact_Phone1"
33630             .Sort_line.Left = .Contact_Phone1_lbl_cmd_Phone1.Left
33640             .Sort_lbl.Left = (((.Contact_Phone1_lbl_cmd_Phone1.Left + .Contact_Phone1_lbl_cmd_Phone1.Width) - _
                    lngSortLbl_Width) - (1& * lngTpp))
33650           Case "Contact_Phone2"
33660             .Sort_line.Left = .Contact_Phone1_lbl_cmd_Phone2.Left
33670             .Sort_lbl.Left = (((.Contact_Phone1_lbl_cmd_Phone2.Left + .Contact_Phone1_lbl_cmd_Phone2.Width) - _
                    lngSortLbl_Width) + (7& * lngTpp))
33680           End Select
33690         End If
33700         blnSortHere = False: strTmp01 = vbNullString
33710         If .Sort_line.Left = .Contact_Email_lbl_cmd_Email.Left Then
33720           blnSortHere = True
33730           strTmp01 = "Contact_Email"
33740         ElseIf .Sort_line.Left = .Contact_Email_lbl_cmd_Fax.Left Then
33750           blnSortHere = True
33760           strTmp01 = "Contact_Fax"
33770         End If
33780         .Contact_Email.Left = ((lngEmail_Left - (lngLocContact_Width + lngCountry_Width)) - lngTpp)  'CHECK THIS!
33790         .Contact_Fax.Left = .Contact_Email.Left
33800         .Contact_Fax_Format.Left = .Contact_Email.Left
33810         .Contact_Email_lbl.Left = .Contact_Email.Left
33820         .Contact_Email_lbl_dim_hi.Left = (.Contact_Email_lbl.Left + lngTpp)
33830         .Contact_Email_lbl_line.Left = .Contact_Email_lbl.Left
33840         .Contact_Email_lbl_line_dim_hi.Left = (.Contact_Email_lbl_line.Left + lngTpp)
33850         .Contact_Email_lbl_cmd_Email.Left = .Contact_Email.Left
33860         .Contact_Email_lbl_cmd_Fax.Left = ((.Contact_Email.Left + .Contact_Email.Width) - .Contact_Email_lbl_cmd_Fax.Width)
33870         .Contact_Email_cmd.Left = (.Contact_Email.Left + (lngEmailCmd_Left - lngEmail_Left))
33880         If blnSortHere = True Then
33890           Select Case strTmp01
                Case "Contact_Email"
33900             .Sort_line.Left = .Contact_Email_lbl_cmd_Email.Left
33910             .Sort_lbl.Left = ((.Contact_Email_lbl_cmd_Email.Left + .Contact_Email_lbl_cmd_Email.Width) - lngSortLbl_Width)
33920           Case "Contact_Fax"
33930             .Sort_line.Left = .Contact_Email_lbl_cmd_Fax.Left
33940             .Sort_lbl.Left = ((.Contact_Email_lbl_cmd_Fax.Left + .Contact_Email_lbl_cmd_Fax.Width) - lngSortLbl_Width)
33950           End Select
33960         End If
33970         blnSortHere = False: strTmp01 = vbNullString
33980         .Width = (lngForm_Width - (lngLocContact_Width + lngCountry_Width))
33990       End If
34000     End Select
34010     If blnResort = True Then
34020       SortNow "Form_Load"  ' ** Procedure: Below.
34030     End If
34040   End With

EXITP:
34050   Exit Sub

ERRH:
34060   Select Case ERR.Number
        Case Else
34070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34080   End Select
34090   Resume EXITP

End Sub

Public Sub ResetChanged()

34100 On Error GoTo ERRH

        Const THIS_PROC As String = "ResetChanged"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long

34110   cmdSave_Click  ' ** Procedure: Above.
34120   DoEvents

34130   Set dbs = CurrentDb
34140   With dbs
34150     Set rst = .OpenRecordset("tblAccount_Contact_Staging", dbOpenDynaset, dbConsistent)
34160     With rst
34170       .MoveLast
34180       lngRecs = .RecordCount
34190       .MoveFirst
34200       For lngX = 1& To lngRecs
34210         If ![Contact_Changed] = True Then
34220           .Edit
34230           ![Contact_Changed] = False
34240           ![Contact_DateModified] = Now()
34250           .Update
34260         End If
34270         If lngX < lngRecs Then .MoveNext
34280       Next
34290       .Close
34300     End With
34310     Set rst = Nothing
34320   End With  ' ** dbs.
34330   Set dbs = Nothing

34340   With Me
34350     .Refresh
34360     .Requery
34370   End With

EXITP:
34380   Set rst = Nothing
34390   Set dbs = Nothing
34400   Exit Sub

ERRH:
34410   Select Case ERR.Number
        Case Else
34420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34430   End Select
34440   Resume EXITP

End Sub

Public Function DelRec_AC(lngContactID As Long) As Boolean

34500 On Error GoTo ERRH

        Const THIS_PROC As String = "DelRec_AC"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnRetVal As Boolean

34510   blnRetVal = True

34520   With Me
34530     Set dbs = CurrentDb
34540     With dbs
            ' ** Delete tblAccount_Contact_Staging, by specified [conid].
34550       Set qdf = .QueryDefs("qryAccount_Contacts_19")
34560       With qdf.Parameters
34570         ![conid] = lngContactID
34580       End With
34590       qdf.Execute
34600       .Close
34610     End With
34620     .Requery
34630   End With

EXITP:
34640   Set qdf = Nothing
34650   Set dbs = Nothing
34660   DelRec_AC = blnRetVal
34670   Exit Function

ERRH:
34680   blnRetVal = False
34690   Select Case ERR.Number
        Case Else
34700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34710   End Select
34720   Resume EXITP

End Function

Public Function RecCnt() As Long

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

34810   lngRetVal = 0&

34820   With Me
34830     Set rst = .RecordsetClone
34840     With rst
34850       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
34860       Else
34870         .MoveLast
34880         lngRetVal = .RecordCount
34890       End If
34900       .Close
34910     End With
34920   End With

EXITP:
34930   Set rst = Nothing
34940   RecCnt = lngRetVal
34950   Exit Function

ERRH:
34960   lngRetVal = 0&
34970   Select Case ERR.Number
        Case Else
34980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34990   End Select
35000   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

35110   Select Case intWhere
        Case acCmdRecordsGoToFirst
35120     DoCmd.RunCommand acCmdRecordsGoToFirst
35130   Case acCmdRecordsGoToLast
35140     DoCmd.RunCommand acCmdRecordsGoToLast
35150   Case acCmdRecordsGoToPrevious
35160     DoCmd.RunCommand acCmdRecordsGoToPrevious
35170   Case acCmdRecordsGoToNext
35180     DoCmd.RunCommand acCmdRecordsGoToNext
35190   Case acCmdRecordsGoToNew
35200     DoCmd.RunCommand acCmdRecordsGoToNew
35210   Case Else
35220     If IsMissing(varID) = False Then
35230       With Me
35240         Set rst = .RecordsetClone
35250         With rst
35260           .FindFirst "[Contact_ID] = " & CStr(varID)
35270           If .NoMatch = False Then
35280             Me.Bookmark = .Bookmark
35290           End If
35300           .Close
35310         End With
35320       End With
35330     End If
35340   End Select

EXITP:
35350   Set rst = Nothing
35360   Exit Sub

ERRH:
35370   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
35380   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
35390   Case Else
35400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35410   End Select
35420   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

35500 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim ctl As Access.Control
        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Static lngTmp03 As Long

        Const strStdAsc As String = ", [alphasort], [Contact_Number]"

35510   With Me
35520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
35530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
35540     End If
35550     lngRecsCur = RecCnt  ' ** Function: Above.
35560     If lngRecsCur > 0& Then
35570       .Controls(strSortLbl).Visible = False
35580       .Controls(strSortLine).Visible = False
35590       .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push of the right side of the form.
35600       .Sort_lbl_opg1.Visible = False
35610       .Sort_lbl_opg2.Visible = False
35620       If strProc = "Form_Load" Then
35630         strSortNow = strSortOrig
35640         lngSortLbl_Left = ((.Controls("accountno1_lbl_cmd_accountno").Left + .Controls("accountno1_lbl_cmd_accountno").Width) - lngSortLbl_Width)
35650         lngSortLbl_Top = .Controls(strSortLine).Top
35660         lngSortLine_Left = .Controls("accountno1_lbl_cmd_accountno").Left
35670         lngSortLine_Width = (.Controls("accountno1_lbl_cmd_accountno").Width + lngTpp)
35680         .Controls(strSortLbl).Left = lngSortLbl_Left
35690         .Controls(strSortLbl).Top = lngSortLbl_Top
35700         .Controls(strSortLine).Left = lngSortLine_Left
35710         .Controls(strSortLine).Width = lngSortLine_Width
35720         .Controls(strSortLbl).Caption = strArwUp
35730         .Controls(strSortLbl).ForeColor = CLR_DKBLU
35740       Else
35750         strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
35760         Select Case strCalled
              Case "accountno"
35770           lngSortLbl_Left = ((.Controls("accountno1_lbl_cmd_accountno").Left + .Controls("accountno1_lbl_cmd_accountno").Width) - _
                  lngSortLbl_Width)
35780           lngSortLine_Left = .Controls("accountno1_lbl_cmd_accountno").Left
35790           lngSortLine_Width = (.Controls("accountno1_lbl_cmd_accountno").Width + lngTpp)
35800         Case "shortname"
35810           lngSortLbl_Left = (((.Controls("accountno1_lbl_cmd_shortname").Left + .Controls("accountno1_lbl_cmd_shortname").Width) - _
                  lngSortLbl_Width) + (6& * lngTpp))
35820           lngSortLine_Left = .Controls("accountno1_lbl_cmd_shortname").Left
35830           lngSortLine_Width = (.Controls("accountno1_lbl_cmd_shortname").Width + lngTpp)
35840         Case "Contact_Address1"
35850           lngSortLbl_Left = ((.Controls("Contact_Address1_lbl_cmd_Address1").Left + .Controls("Contact_Address1_lbl_cmd_Address1").Width) - _
                  lngSortLbl_Width)
35860           lngSortLine_Left = .Controls("Contact_Address1_lbl_cmd_Address1").Left
35870           lngSortLine_Width = (.Controls("Contact_Address1_lbl_cmd_Address1").Width + lngTpp)
35880         Case "Contact_Address2"
35890           lngSortLbl_Left = ((.Controls("Contact_Address1_lbl_cmd_Address2").Left + .Controls("Contact_Address1_lbl_cmd_Address2").Width) - _
                  lngSortLbl_Width)
35900           lngSortLine_Left = .Controls("Contact_Address1_lbl_cmd_Address2").Left
35910           lngSortLine_Width = (.Controls("Contact_Address1_lbl_cmd_Address2").Width + lngTpp)
35920         Case "Contact_City"
35930           lngSortLbl_Left = ((.Controls("Contact_City_lbl_cmd_City").Left + .Controls("Contact_City_lbl_cmd_City").Width) - _
                  lngSortLbl_Width)
35940           lngSortLine_Left = .Controls("Contact_City_lbl_cmd_City").Left
35950           lngSortLine_Width = (.Controls("Contact_City_lbl_cmd_City").Width + lngTpp)
35960         Case "Contact_State"
35970           lngSortLbl_Left = ((.Controls("Contact_City_lbl_cmd_State").Left + .Controls("Contact_City_lbl_cmd_State").Width) - _
                  lngSortLbl_Width)
35980           lngSortLine_Left = .Controls("Contact_City_lbl_cmd_State").Left
35990           lngSortLine_Width = (.Controls("Contact_City_lbl_cmd_State").Width + lngTpp)
36000         Case "Contact_Zip"
36010           lngSortLbl_Left = (((.Controls("Contact_City_lbl_cmd_Zip").Left + .Controls("Contact_City_lbl_cmd_Zip").Width) - _
                  lngSortLbl_Width) + (2& * lngTpp))
36020           lngSortLine_Left = .Controls("Contact_City_lbl_cmd_Zip").Left
36030           lngSortLine_Width = (.Controls("Contact_City_lbl_cmd_Zip").Width + lngTpp)
36040         Case "Contact_Country"
36050           lngSortLbl_Left = ((.Controls("Contact_Country_lbl_cmd_Country").Left + .Controls("Contact_Country_lbl_cmd_Country").Width) - _
                  lngSortLbl_Width)
36060           lngSortLine_Left = .Controls("Contact_Country_lbl_cmd_Country").Left
36070           lngSortLine_Width = (.Controls("Contact_Country_lbl_cmd_Country").Width + lngTpp)
36080         Case "Contact_PostalCode"
36090           lngSortLbl_Left = (((.Controls("Contact_Country_lbl_cmd_PostalCode").Left + .Controls("Contact_Country_lbl_cmd_PostalCode").Width) - _
                  lngSortLbl_Width) + (6& * lngTpp))
36100           lngSortLine_Left = .Controls("Contact_Country_lbl_cmd_PostalCode").Left
36110           lngSortLine_Width = (.Controls("Contact_Country_lbl_cmd_PostalCode").Width + lngTpp)
36120         Case "Contact_Phone1"
36130           lngSortLbl_Left = (((.Controls("Contact_Phone1_lbl_cmd_Phone1").Left + .Controls("Contact_Phone1_lbl_cmd_Phone1").Width) - _
                  lngSortLbl_Width) - (1& * lngTpp))
36140           lngSortLine_Left = .Controls("Contact_Phone1_lbl_cmd_Phone1").Left
36150           lngSortLine_Width = (.Controls("Contact_Phone1_lbl_cmd_Phone1").Width + lngTpp)
36160         Case "Contact_Phone2"
36170           lngSortLbl_Left = (((.Controls("Contact_Phone1_lbl_cmd_Phone2").Left + .Controls("Contact_Phone1_lbl_cmd_Phone2").Width) - _
                  lngSortLbl_Width) + (7& * lngTpp))
36180           lngSortLine_Left = .Controls("Contact_Phone1_lbl_cmd_Phone2").Left
36190           lngSortLine_Width = (.Controls("Contact_Phone1_lbl_cmd_Phone2").Width + lngTpp)
36200         Case "Contact_Email"
36210           lngSortLbl_Left = ((.Controls("Contact_Email_lbl_cmd_Email").Left + .Controls("Contact_Email_lbl_cmd_Email").Width) - _
                  lngSortLbl_Width)
36220           lngSortLine_Left = .Controls("Contact_Email_lbl_cmd_Email").Left
36230           lngSortLine_Width = (.Controls("Contact_Email_lbl_cmd_Email").Width + lngTpp)
36240         Case "Contact_Fax"
36250           lngSortLbl_Left = ((.Controls("Contact_Email_lbl_cmd_Fax").Left + .Controls("Contact_Email_lbl_cmd_Fax").Width) - _
                  lngSortLbl_Width)
36260           lngSortLine_Left = .Controls("Contact_Email_lbl_cmd_Fax").Left
36270           lngSortLine_Width = (.Controls("Contact_Email_lbl_cmd_Fax").Width + lngTpp)
36280         Case Else  ' ** Contact_Number, Contact_Name, opgLocContact.
36290           lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
36300           lngSortLine_Left = .Controls(strCalled & "_lbl").Left
36310           lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
36320         End Select
36330         lngSortLbl_Top = .Controls(strSortLine).Top
36340         .Controls(strSortLbl).Left = lngSortLbl_Left
36350         .Controls(strSortLbl).Top = lngSortLbl_Top
36360         .Controls(strSortLine).Left = lngSortLine_Left
36370         .Controls(strSortLine).Width = lngSortLine_Width
36380         If strCalled = "accountno" Or strCalled = "shortname" Then
36390           Select Case strCalled
                Case "accountno"
36400             strSortAsc = strSortOrig
36410           Case "shortname"
36420             strSortAsc = "[shortname], [Contact_Number]"
36430           End Select
36440           If strSortNow = strSortAsc Then
36450             Select Case strCalled
                  Case "accountno"
36460               strSortNow = "[alphasort] DESC, [Contact_Number] DESC"
36470             Case "shortname"
36480               strSortNow = "[shortname] DESC, [Contact_Number] DESC"
36490             End Select
36500             .Controls(strSortLbl).Caption = strArwDn
36510             .Controls(strSortLbl).ForeColor = CLR_DKRED
36520           Else
36530             strSortNow = strSortAsc
36540             .Controls(strSortLbl).Caption = strArwUp
36550             .Controls(strSortLbl).ForeColor = CLR_DKBLU
36560           End If
36570         Else
36580           .Controls(strSortLbl).Caption = strArwUp
36590           .Controls(strSortLbl).ForeColor = CLR_DKBLU
36600           Select Case strCalled
                Case "Contact_Number"
36610             strSortAsc = "[Contact_Number], [alphasort], [shortname]"
36620             lngSortLbl_Left = (lngSortLbl_Left + (4& * lngTpp))
36630             .Controls(strSortLbl).Left = lngSortLbl_Left
36640           Case "Contact_Name"
36650             Select Case .Controls(strCalled).Visible
                  Case True
36660               strSortAsc = "[Contact_Name]" & strStdAsc
36670             Case False
36680               strSortAsc = "[Contact_Name_LastFirst]" & strStdAsc
36690             End Select
36700           Case "Contact_Address1"
36710             strSortAsc = "[Contact_Address1]" & strStdAsc
36720           Case "Contact_Address2"
36730             strSortAsc = "[Contact_Address2]" & strStdAsc
36740           Case "Contact_City"
36750             strSortAsc = "[Contact_City]" & strStdAsc
36760           Case "opgLocContact"
36770             strSortAsc = "[opgLocContact]" & strStdAsc
36780             lngSortLbl_Top = (lngSortLbl_Top - (2& * lngTpp))
36790             lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
36800             .Controls(strSortLbl).Top = lngSortLbl_Top
36810             .Controls(strSortLbl).Left = lngSortLbl_Left
36820           Case "Contact_State"
36830             strSortAsc = "[Contact_State]" & strStdAsc
36840           Case "Contact_Zip"
36850             Select Case .Controls(strCalled).Visible
                  Case True
36860               strSortAsc = "[Contact_Zip]" & strStdAsc
36870             Case False
36880               strSortAsc = "[Contact_Zip_Format]" & strStdAsc
36890             End Select
36900           Case "Contact_Country"
36910             strSortAsc = "[Contact_Country]" & strStdAsc
36920           Case "Contact_PostalCode"
36930             strSortAsc = "[Contact_PostalCode]" & strStdAsc
36940           Case "Contact_Phone1"
36950             Select Case .Controls(strCalled).Visible
                  Case True
36960               strSortAsc = "[Contact_Phone1]" & strStdAsc
36970             Case False
36980               strSortAsc = "[Contact_Phone1_Format]" & strStdAsc
36990             End Select
37000           Case "Contact_Phone1"
37010             Select Case .Controls(strCalled).Visible
                  Case True
37020               strSortAsc = "[Contact_Phone1]" & strStdAsc
37030             Case False
37040               strSortAsc = "[Contact_Phone1_Format]" & strStdAsc
37050             End Select
37060           Case "Contact_Fax"
37070             Select Case .Controls(strCalled).Visible
                  Case True
37080               strSortAsc = "[Contact_Fax]" & strStdAsc
37090             Case False
37100               strSortAsc = "[Contact_Fax_Format]" & strStdAsc
37110             End Select
37120           Case "Contact_Email"
37130             strSortAsc = "[Contact_Email]" & strStdAsc
37140           End Select
37150           If strSortNow = strSortAsc Then
37160             intCnt = (CharCnt(strSortAsc, ",") + 1)  ' ** Module Function: modStringFuncs.
37170             Select Case intCnt
                  Case 1
                    ' ** [accounttype_description] DESC.
37180               strSortNow = strSortNow & " DESC"
37190             Case 2
                    ' ** [accounttypegroup_name] DESC, [accounttype] DESC.
37200               intPos01 = InStr(strSortAsc, ",")
37210               strSortNow = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
37220             Case 3
37230               intPos01 = InStr(strSortAsc, ",")
37240               strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
37250               strTmp02 = Mid(strSortAsc, intPos01)
37260               intPos01 = InStr(2, strSortAsc, ",")
37270               strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC" & Mid(strTmp02, intPos01) & " DESC"
37280               strSortNow = strTmp01
37290             End Select
37300             .Controls(strSortLbl).Caption = strArwDn
37310             .Controls(strSortLbl).ForeColor = CLR_DKRED
37320           Else
37330             strSortNow = strSortAsc
37340           End If
37350         End If
37360       End If
37370       .Controls(strSortLbl).Visible = True
37380       If strCalled = "opgLocContact" Then
37390         .Sort_lbl_opg1.Visible = True
37400         .Sort_lbl_opg2.Visible = True
37410       Else
37420         .Controls(strSortLine).Visible = True
37430       End If
37440       .OrderBy = strSortNow
37450       .OrderByOn = True
37460     Else
37470       If lngTmp03 = 0& Then
37480         .Sort_lbl.Left = (.Sort_lbl.Left + (3& * lngTpp))
37490         lngTmp03 = lngTmp03 + 1&
37500       End If
37510     End If  ' ** lngRecsCur.
37520   End With

EXITP:
37530   Set ctl = Nothing
37540   Exit Sub

ERRH:
37550   Select Case ERR.Number
        Case Else
37560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37570   End Select
37580   Resume EXITP

End Sub

Public Function SortNow_Get() As String

37600 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

37610   If strSortNow = vbNullString Then
37620     strRetVal = strSortOrig
37630   Else
37640     strRetVal = strSortNow
37650   End If

EXITP:
37660   SortNow_Get = strRetVal
37670   Exit Function

ERRH:
37680   strRetVal = RET_ERR
37690   Select Case ERR.Number
        Case Else
37700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
37710   End Select
37720   Resume EXITP

End Function


