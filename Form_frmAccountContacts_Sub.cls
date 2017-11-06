VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountContacts_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountContacts_Sub"

'VGC 07/24/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd on frmAccountContacts}
' **   Enable Country:   B {chkEnableCountry_Wide on frmAccountContacts}
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
' **   Send Email:       E {Contact_Email_cmd on frmAccountContacts_Sub2}
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

Private lngCountry_Left As Long, lngPostalCode_Left As Long, lngForm_Width As Long
Private lngOpgLoc_Left As Long, lngOpgLoc_Diff As Long, lngState_Left As Long, lngZip_Left As Long
Private lngPhone1_Left As Long, lngPhone2_Left As Long, lngFax_Left As Long, lngEmail_Left As Long
Private lngShortName_Diff As Long, lngCountry_Diff As Long, lngShortName_Width As Long, lngCtlSep As Long

Private strSortNow As String, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private varOldValue As Variant
Private lngRecsCur As Long, lngTpp As Long, blnAddComma As Boolean
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

200         lngSortLbl_Width = .Sort_lbl.Width
210         blnAddComma = False

220         .form_width_line.Left = 0&  ' ** Just get it out of the way.
230         lngForm_Width = .Width
240         lngOpgLoc_Left = .opgLocContact.Left
250         lngState_Left = .Contact_State.Left
260         lngZip_Left = .Contact_Zip.Left
270         lngCountry_Left = .Contact_Country.Left
280         lngPostalCode_Left = .Contact_PostalCode.Left
290         lngPhone1_Left = .Contact_Phone1.Left
300         lngPhone2_Left = .Contact_Phone2.Left
310         lngFax_Left = .Contact_Fax.Left
320         lngEmail_Left = .Contact_Email.Left
330         lngShortName_Diff = (.shortname.Width - .accountno.Width)
340         lngShortName_Width = .shortname.Width
350         lngCtlSep = (4& * lngTpp)
360         lngCountry_Diff = (lngPhone1_Left - lngCountry_Left)
370         lngOpgLoc_Diff = (lngState_Left - lngOpgLoc_Left)

380       Else
390         Cancel = -1
400       End If
410     End With

420     If Cancel = -1 Then
430       Beep
440       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
450     End If

EXITP:
460     Exit Sub

ERRH:
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me
620       .accountno.Enabled = False
630       .accountno.Locked = True
640       .accountno.ForeColor = CLR_DISABLED_FG
650       .accountno.BackColor = CLR_DISABLED_BG
660     End With

EXITP:
670     Exit Sub

ERRH:
680     Select Case ERR.Number
        Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

810     With Me
820       .Username = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
830       .Contact_DateModified = Now()
840     End With

EXITP:
850     Exit Sub

ERRH:
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim frmSub2 As Access.Form
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Enable Country:   B {chkEnableCountry_Wide on frmAccountContacts}
        ' **   Exit:             X {cmdClose on frmAccountContacts}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Zip Unfrmat:      D {opgFormatZip_optUnformatted on frmAccountContacts}
        ' **   Send Email:       E {Contact_Email_cmd on frmAccountContacts_Sub2}
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
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyUp
970         intRetVal = 0
980         MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
990       Case vbKeyDown
1000        intRetVal = 0
1010        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1020      Case vbKeyEscape
1030        With Me
1040          intRetVal = 0
1050          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountContacts.
1060        End With
1070      End Select
1080    End If

        ' ** Alt keys.
1090    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1100      Select Case intRetVal
          Case vbKeyB
1110        With Me
1120          intRetVal = 0
1130          Select Case .Parent.opgView
              Case .Parent.opgView_optWide.OptionValue
1140            If .Parent.chkEnableCountry_Wide.Enabled = True Then
1150              DoCmd.SelectObject acForm, .Parent.Name, False
1160              .Parent.chkEnableCountry_Wide.SetFocus
1170            Else
1180              Beep
1190            End If
1200          Case .Parent.opgView_optCompact.OptionValue
1210            If .Parent.chkEnableCountry_Compact.Enabled = True Then
1220              DoCmd.SelectObject acForm, .Parent.Name, False
1230              .Parent.chkEnableCountry_Compact.SetFocus
1240            Else
1250              Beep
1260            End If
1270          End Select
1280        End With
1290      Case vbKeyX
1300        With Me
1310          intRetVal = 0
1320          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountContacts.
1330        End With
1340      End Select
1350    End If

        ' ** Ctrl keys.
1360    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1370      Select Case intRetVal
          Case vbKeyD
            ' **   Zip Unfrmat:      D {opgFormatZip_optUnformatted}
1380        With Me
1390          intRetVal = 0
1400          DoCmd.SelectObject acForm, .Parent.Name, False
1410          .Parent.opgFormatZip = .Parent.opgFormatZip_optUnformatted.OptionValue
1420          .Parent.opgFormatZip.SetFocus
1430          .Parent.opgFormatZip_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
1440        End With
1450      Case vbKeyE
1460        With Me
1470          intRetVal = 0
1480          Select Case .Parent.opgView
              Case .Parent.opgView_optWide.OptionValue
1490            Beep
1500          Case .Parent.opgView_optCompact.OptionValue  ' ** I know, they shouldn't event be able to get here!
1510            Set frmSub2 = .Parent.frmAccountContacts_Sub2.Form
1520            If frmSub2.Contact_Email_cmd.Visible = True And frmSub2.Contact_Email_cmd.Enabled = True Then
1530              DoCmd.SelectObject acForm, .Parent.Name, False
1540              .Parent.frmAccountContacts_Sub2.SetFocus
1550              frmSub2.Contact_Email_cmd.SetFocus
1560              frmSub2.Contact_Email_cmd_Click  ' ** Form Procedure: frmAccountContacts_Sub2.
1570            End If
1580          End Select
1590        End With
1600      Case vbKeyF
            ' **   Zip Frmat:        F {opgFormatZip_optFormatted}
1610        With Me
1620          intRetVal = 0
1630          DoCmd.SelectObject acForm, .Parent.Name, False
1640          .Parent.opgFormatZip = .Parent.opgFormatZip_optFormatted.OptionValue
1650          .Parent.opgFormatZip.SetFocus
1660          .Parent.opgFormatZip_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
1670        End With
1680      Case vbKeyG
1690        With Me
1700          intRetVal = 0
1710          If .Parent.chkPageOf.Enabled = True Then
1720            DoCmd.SelectObject acForm, .Parent.Name, False
1730            .Parent.chkPageOf.SetFocus
1740            .Parent.chkPageOf = (Not .Parent.chkPageOf)
1750            .Parent.chkPageOf_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
1760          End If
1770        End With
1780      Case vbKeyI
            ' **   Name Last/First:  I {opgFormatName_optLastFirst}
1790        With Me
1800          intRetVal = 0
1810          DoCmd.SelectObject acForm, .Parent.Name, False
1820          .Parent.opgFormatName = .Parent.opgFormatName_optLastFirst.OptionValue
1830          .Parent.opgFormatName.SetFocus
1840          .Parent.opgFormatName_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
1850        End With
1860      Case vbKeyM
            ' **   Phone Unfrmat:    M {opgFormatPhone_optUnformatted}
1870        With Me
1880          intRetVal = 0
1890          DoCmd.SelectObject acForm, .Parent.Name, False
1900          .Parent.opgFormatPhone = .Parent.opgFormatPhone_optUnformatted.OptionValue
1910          .Parent.opgFormatPhone.SetFocus
1920          .Parent.opgFormatPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
1930        End With
1940      Case vbKeyO
1950        With Me
1960          intRetVal = 0
1970          Select Case .Parent.opgView
              Case .Parent.opgView_optWide.OptionValue
1980            If .opgLocContact.Visible = True And .opgLocContact.Enabled = True Then
1990              .opgLocContact.SetFocus
2000              .opgLocContact = .opgLocContact_optOther.OptionValue
2010              opgLocContact_AfterUpdate  ' ** Procedure: Below.
2020            Else
2030              Beep
2040            End If
2050          Case .Parent.opgView_optCompact.OptionValue  ' ** I know, they shouldn't event be able to get here!
2060            Beep
2070          End Select
2080        End With
2090      Case vbKeyR
            ' **   Phone Frmat:      R {opgFormatPhone_optFormatted}
2100        With Me
2110          intRetVal = 0
2120          DoCmd.SelectObject acForm, .Parent.Name, False
2130          .Parent.opgFormatPhone = .Parent.opgFormatPhone_optFormatted.OptionValue
2140          .Parent.opgFormatPhone.SetFocus
2150          .Parent.opgFormatPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2160        End With
2170      Case vbKeyS
2180        intRetVal = 0
2190        cmdSave_Click  ' ** Procedure: Below.
2200      Case vbKeyT
            ' **   Name As Written:  T {opgFormatName_optAsWritten}
2210        With Me
2220          intRetVal = 0
2230          DoCmd.SelectObject acForm, .Parent.Name, False
2240          .Parent.opgFormatName = .Parent.opgFormatName_optAsWritten.OptionValue
2250          .Parent.opgFormatName.SetFocus
2260          .Parent.opgFormatName_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2270        End With
2280      Case vbKeyU
2290        With Me
2300          intRetVal = 0
2310          Select Case .Parent.opgView
              Case .Parent.opgView_optWide.OptionValue
2320            If .opgLocContact.Visible = True And .opgLocContact.Enabled = True Then
2330              .opgLocContact.SetFocus
2340              .opgLocContact = .opgLocContact_optUSA.OptionValue
2350              opgLocContact_AfterUpdate  ' ** Procedure: Below.
2360            Else
2370              Beep
2380            End If
2390          Case .Parent.opgView_optCompact.OptionValue  ' ** I know, they shouldn't even be able to get here!
2400            Beep
2410          End Select
2420        End With
2430      Case vbKeyTab, vbKeyReturn
2440        With Me
2450          intRetVal = 0
2460          DoCmd.SelectObject acForm, .Parent.Name, False
2470          .Parent.opgView.SetFocus
2480        End With
2490      Case vbKeyPageUp
2500        intRetVal = 0
2510        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2520      Case vbKeyPageDown
2530        intRetVal = 0
2540        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2550      Case vbKeyHome
2560        With Me
2570          intRetVal = 0
2580          If .Contact_Name.Visible = True And .Contact_Name.Enabled = True Then
2590            .Contact_Name.SetFocus
2600          Else
2610            .Contact_Name_LastFirst.SetFocus
2620          End If
2630        End With
2640      Case vbKeyEnd
2650        With Me
2660          intRetVal = 0
2670          .Contact_Email.SetFocus
2680        End With
2690      End Select
2700    End If

        ' ** Ctrl-Alt keys.
2710    If intCtrlDown And intAltDown And (Not intShiftDown) Then
2720      Select Case intRetVal
          Case vbKeyA
            ' **   CSZ Separate:     A {opgPrintCSZ_optSeparate, opgPrintCSZCP_optSeparate}
2730        With Me
2740          intRetVal = 0
2750          DoCmd.SelectObject acForm, .Parent.Name, False
2760          If .Parent.opgPrintCSZ.Visible = True And .Parent.opgPrintCSZ.Enabled = True Then
2770            .Parent.opgPrintCSZ = .Parent.opgPrintCSZ_optSeparate.OptionValue
2780            .Parent.opgPrintCSZ.SetFocus
2790            .Parent.opgPrintCSZ_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2800          ElseIf .Parent.opgPrintCSZCP.Visible = True And .Parent.opgPrintCSZCP.Enabled = True Then
2810            .Parent.opgPrintCSZCP = .Parent.opgPrintCSZCP_optSeparate.OptionValue
2820            .Parent.opgPrintCSZCP.SetFocus
2830            .Parent.opgPrintCSZCP_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2840          End If
2850        End With
2860      Case vbKeyB
            ' **   CSZ Combined:     B {opgPrintCSZ_optCombined, opgPrintCSZCP_optCombined}
2870        With Me
2880          intRetVal = 0
2890          DoCmd.SelectObject acForm, .Parent.Name, False
2900          If .Parent.opgPrintCSZ.Visible = True And .Parent.opgPrintCSZ.Enabled = True Then
2910            .Parent.opgPrintCSZ = .Parent.opgPrintCSZ_optCombined.OptionValue
2920            .Parent.opgPrintCSZ.SetFocus
2930            .Parent.opgPrintCSZ_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2940          ElseIf .Parent.opgPrintCSZCP.Visible = True And .Parent.opgPrintCSZCP.Enabled = True Then
2950            .Parent.opgPrintCSZCP = .Parent.opgPrintCSZCP_optCombined.OptionValue
2960            .Parent.opgPrintCSZCP.SetFocus
2970            .Parent.opgPrintCSZCP_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
2980          End If
2990        End With
3000      Case vbKeyC
            ' **   Addr Combined:    C {opgPrintAddress_optCombined}
3010        With Me
3020          intRetVal = 0
3030          DoCmd.SelectObject acForm, .Parent.Name, False
3040          .Parent.opgPrintAddress = .Parent.opgPrintAddress_optCombined.OptionValue
3050          .Parent.opgPrintAddress.SetFocus
3060          .Parent.opgPrintAddress_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3070        End With
3080      Case vbKeyN
            ' **   Addr None:        N {opgPrintAddress_optNone}
3090        With Me
3100          intRetVal = 0
3110          DoCmd.SelectObject acForm, .Parent.Name, False
3120          .Parent.opgPrintAddress = .Parent.opgPrintAddress_optNone.OptionValue
3130          .Parent.opgPrintAddress.SetFocus
3140          .Parent.opgPrintAddress_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3150        End With
3160      Case vbKeyO
            ' **   CSZ None:         O {opgPrintCSZ_optNone, opgPrintCSZCP_optNone}
3170        With Me
3180          intRetVal = 0
3190          DoCmd.SelectObject acForm, .Parent.Name, False
3200          If .Parent.opgPrintCSZ.Visible = True And .Parent.opgPrintCSZ.Enabled = True Then
3210            .Parent.opgPrintCSZ = .Parent.opgPrintCSZ_optNone.OptionValue
3220            .Parent.opgPrintCSZ.SetFocus
3230            .Parent.opgPrintCSZ_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3240          ElseIf .Parent.opgPrintCSZCP.Visible = True And .Parent.opgPrintCSZCP.Enabled = True Then
3250            .Parent.opgPrintCSZCP = .Parent.opgPrintCSZCP_optNone.OptionValue
3260            .Parent.opgPrintCSZCP.SetFocus
3270            .Parent.opgPrintCSZCP_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3280          End If
3290        End With
3300      Case vbKeyS
            ' **   Addr Separate:    S {opgPrintAddress_optSeparate}
3310        With Me
3320          intRetVal = 0
3330          DoCmd.SelectObject acForm, .Parent.Name, False
3340          .Parent.opgPrintAddress = .Parent.opgPrintAddress_optSeparate.OptionValue
3350          .Parent.opgPrintAddress.SetFocus
3360          .Parent.opgPrintAddress_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3370        End With
3380      End Select
3390    End If

        ' ** Ctrl-Shift keys.
3400    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3410      Select Case intRetVal
          Case vbKeyC
            ' **   Phone Combined:   C {opgPrintPhone_optCombined}
3420        With Me
3430          intRetVal = 0
3440          DoCmd.SelectObject acForm, .Parent.Name, False
3450          .Parent.opgPrintPhone = .Parent.opgPrintPhone_optCombined.OptionValue
3460          .Parent.opgPrintPhone.SetFocus
3470          .Parent.opgPrintPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3480        End With
3490      Case vbKeyE
            ' **   Email None:       E {opgPrintEmail_optNone}
3500        With Me
3510          intRetVal = 0
3520          DoCmd.SelectObject acForm, .Parent.Name, False
3530          .Parent.opgPrintEmail = .Parent.opgPrintEmail_optNone.OptionValue
3540          .Parent.opgPrintEmail.SetFocus
3550          .Parent.opgPrintEmail_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3560        End With
3570      Case vbKeyF
3580        With Me
3590          intRetVal = 0
3600          DoCmd.SelectObject acForm, .Parent.Name, False
3610          .Parent.FocusHolder.SetFocus
3620        End With
3630      Case vbKeyN
            ' **   Phone None:       N {opgPrintPhone_optNone}
3640        With Me
3650          intRetVal = 0
3660          DoCmd.SelectObject acForm, .Parent.Name, False
3670          .Parent.opgPrintPhone = .Parent.opgPrintPhone_optNone.OptionValue
3680          .Parent.opgPrintPhone.SetFocus
3690          .Parent.opgPrintPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3700        End With
3710      Case vbKeyO
            ' **   Fax None:         O {opgPrintFax_optNone}
3720        With Me
3730          intRetVal = 0
3740          DoCmd.SelectObject acForm, .Parent.Name, False
3750          .Parent.opgPrintFax = .Parent.opgPrintFax_optNone.OptionValue
3760          .Parent.opgPrintFax.SetFocus
3770          .Parent.opgPrintFax_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3780        End With
3790      Case vbKeyP
            ' **   Fax Print:        P {opgPrintFax_optPrint}
3800        With Me
3810          intRetVal = 0
3820          DoCmd.SelectObject acForm, .Parent.Name, False
3830          .Parent.opgPrintFax = .Parent.opgPrintFax_optPrint.OptionValue
3840          .Parent.opgPrintFax.SetFocus
3850          .Parent.opgPrintFax_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3860        End With
3870      Case vbKeyS
            ' **   Phone Separate:   S {opgPrintPhone_optSeparate}
3880        With Me
3890          intRetVal = 0
3900          DoCmd.SelectObject acForm, .Parent.Name, False
3910          .Parent.opgPrintPhone = .Parent.opgPrintPhone_optSeparate.OptionValue
3920          .Parent.opgPrintPhone.SetFocus
3930          .Parent.opgPrintPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
3940        End With
3950      Case vbKeyT
            ' **   Email Print:      T {opgPrintEmail_optPrint}
3960        With Me
3970          intRetVal = 0
3980          DoCmd.SelectObject acForm, .Parent.Name, False
3990          .Parent.opgPrintEmail = .Parent.opgPrintEmail_optPrint.OptionValue
4000          .Parent.opgPrintEmail.SetFocus
4010          .Parent.opgPrintEmail_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
4020        End With
4030      Case vbKeyY
            ' **   Phone 1 Only:     Y {opgPrintPhone_opt1Only}
4040        With Me
4050          intRetVal = 0
4060          DoCmd.SelectObject acForm, .Parent.Name, False
4070          .Parent.opgPrintPhone = .Parent.opgPrintPhone_opt1Only.OptionValue
4080          .Parent.opgPrintPhone.SetFocus
4090          .Parent.opgPrintPhone_AfterUpdate  ' ** Form Procedure: frmAccountContacts.
4100        End With
4110      Case vbKeyTab, vbKeyReturn
4120        With Me
4130          intRetVal = 0
4140          DoCmd.SelectObject acForm, .Parent.Name, False
4150          .Parent.cmdClose.SetFocus
4160        End With
4170      End Select
4180    End If

EXITP:
4190    Set frmSub2 = Nothing
4200    KeyCode = intRetVal
4210    Exit Sub

ERRH:
4220    intRetVal = 0
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Public Sub cmdSave_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

4310    DoCmd.RunCommand acCmdSaveRecord
4320    DoEvents

EXITP:
4330    Exit Sub

ERRH:
4340    Select Case ERR.Number
        Case Else
4350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4360    End Select
4370    Resume EXITP

End Sub

Private Sub accountno_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "accountno_lbl_DblClick"

4410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub shortname_lbl_DblClick(Cancel As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_lbl_DblClick"

4510    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub Contact_Name_lbl_DblClick(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_lbl_DblClick"

4610    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub Contact_Name_Enter()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_Enter"

4710    With Me
4720      varOldValue = Null
4730  On Error Resume Next
4740      varOldValue = .Contact_Name
4750  On Error GoTo ERRH
4760    End With

EXITP:
4770    Exit Sub

ERRH:
4780    Select Case ERR.Number
        Case 2427  ' ** You entered an expression that has no value.
          ' ** No records.
4790    Case Else
4800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4810    End Select
4820    Resume EXITP

End Sub

Private Sub Contact_Name_GotFocus()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_GotFocus"

4910    With Me
4920      .Contact_Name.SelLength = 0
4930      .Contact_Name.SelStart = 99
4940    End With

EXITP:
4950    Exit Sub

ERRH:
4960    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
4970    Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Sub

Private Sub Contact_Name_AfterUpdate()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_AfterUpdate"

        Dim varLast As Variant, varFirst As Variant, varSfx As Variant, varLastFirst As Variant
        Dim strTmp01 As String

5110    With Me
5120      If IsNull(.Contact_Name) = False Then
5130        If Trim(.Contact_Name) <> vbNullString Then
5140          strTmp01 = Trim(FixQuotes(.Contact_Name))  ' ** Module Function: modStringFuncs.
5150          If Compare_StringA_StringB(.Contact_Name, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
5160            .Contact_Name = strTmp01
5170          End If
5180          .Contact_Changed = True

5190          cmdSave_Click  ' ** Procedure: Above.

5200          varLast = Null: varFirst = Null: varSfx = Null: varLastFirst = Null

              ' ** These should be the same expressions as in the queries.
              'varLast = IIf(InStr(strTmp01, "Regional Center") > 0 Or InStr(strTmp01, "(") > 0 Or _
              '  InStr(strTmp01, "Cnty") > 0, strTmp01, IIf(InStr(GetLastWord(strTmp01, True), "~") > 0, _
              '  Left(GetLastWord(strTmp01, True), (InStr(GetLastWord(strTmp01, True), "~") - 1)), _
              '  GetLastWord(strTmp01, True)))   ' ** Module Function: modStringFuncs.
5210          If InStr(strTmp01, "Regional Center") > 0 Or InStr(strTmp01, "(") > 0 Or InStr(strTmp01, "Cnty") > 0 Then
5220            varLast = strTmp01
5230          Else
5240            If InStr(GetLastWord(strTmp01, True), "~") > 0 Then  ' ** Module Function: modStringFuncs.
5250              varLast = Left(GetLastWord(strTmp01, True), (InStr(GetLastWord(strTmp01, True), "~") - 1))  ' ** Module Function: modStringFuncs.
5260            Else
5270              varLast = GetLastWord(strTmp01, True)  ' ** Module Function: modStringFuncs.
5280            End If
5290          End If

5300          varSfx = IIf(InStr(strTmp01, "Regional Center") > 0 Or InStr(strTmp01, "(") > 0 Or _
                InStr(strTmp01, "Cnty") > 0, Null, IIf(InStr(GetLastWord(strTmp01, True), "~") > 0, _
                Mid(GetLastWord(strTmp01, True), (InStr(GetLastWord(strTmp01, True), "~") + 1)), Null))  ' ** Module Function: modStringFuncs.
5310          varFirst = IIf(IsNull(strTmp01) = True, Null, IIf(varLast = strTmp01, Null, _
                Trim(Left(strTmp01, (InStr(strTmp01, varLast) - 1))) & IIf(IsNull(varSfx) = True, "", _
                IIf(Left(varSfx, 1) = ",", varSfx, " " & varSfx))))
5320          varLastFirst = IIf(IsNull(strTmp01) = True, Null, varLast & IIf(IsNull(varFirst) = True, "", ", " & varFirst))

5330          Select Case IsNull(varFirst)
              Case True
5340            If IsNull(.Contact_Name_First) = False Then
5350              .Contact_Name_First = Null
5360            End If
5370          Case False
5380            If Trim(varFirst) <> vbNullString Then
5390              .Contact_Name_First = varFirst
5400            Else
5410              .Contact_Name_First = Null
5420            End If
5430          End Select
5440          Select Case IsNull(varLast)
              Case True
5450            If IsNull(.Contact_Name_Last) = False Then
5460              .Contact_Name_Last = Null
5470            End If
5480          Case False
5490            If Trim(varLast) <> vbNullString Then
5500              .Contact_Name_Last = varLast
5510            Else
5520              .Contact_Name_Last = Null
5530            End If
5540          End Select
5550          Select Case IsNull(varLastFirst)
              Case True
5560            If IsNull(.Contact_Name_LastFirst) = False Then
5570              .Contact_Name_LastFirst = Null
5580            End If
5590          Case False
5600            If Trim(varLastFirst) <> vbNullString Then
5610              .Contact_Name_LastFirst = varLastFirst
5620            Else
5630              .Contact_Name_LastFirst = Null
5640            End If
5650          End Select

5660          cmdSave_Click  ' ** Procedure: Above.

5670        Else
5680          If IsNull(varOldValue) = False Then
5690            .Contact_Name = Null
5700            .Contact_Name_First = Null
5710            .Contact_Name_Last = Null
5720            .Contact_Name_LastFirst = Null
5730            .Contact_Changed = True
5740            cmdSave_Click  ' ** Procedure: Above.
5750          End If
5760        End If  ' ** vbNullString.
5770      Else
5780        If IsNull(varOldValue) = False Then
5790          .Contact_Name_First = Null
5800          .Contact_Name_Last = Null
5810          .Contact_Name_LastFirst = Null
5820          .Contact_Changed = True
5830          cmdSave_Click  ' ** Procedure: Above.
5840        End If
5850      End If  ' ** IsNull.
5860    End With

EXITP:
5870    Exit Sub

ERRH:
5880    Select Case ERR.Number
        Case Else
5890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5900    End Select
5910    Resume EXITP

End Sub

Private Sub Contact_Name_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_KeyDown"

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
6090          .Contact_Address1.SetFocus
6100        End With
6110      End Select
6120    End If

        ' ** Shift keys.
6130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6150        With Me
6160          intRetVal = 0
6170          If .CurrentRecord > 1 Then
6180            MoveRec acCmdRecordsGoToPrevious
6190            .Contact_Email.SetFocus
6200          Else
6210            DoCmd.SelectObject acForm, .Parent.Name, False
6220            .Parent.cmdClose.SetFocus
6230          End If
6240        End With
6250      End Select
6260    End If

EXITP:
6270    KeyCode = intRetVal
6280    Exit Sub

ERRH:
6290    intRetVal = 0
6300    Select Case ERR.Number
        Case Else
6310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6320    End Select
6330    Resume EXITP

End Sub

Private Sub Contact_Name_Exit(Cancel As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_Exit"

6410    varOldValue = Null

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub Contact_Name_LastFirst_Enter()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Name_LastFirst_Enter"

6510    With Me
6520      varOldValue = Null
6530  On Error Resume Next
6540      varOldValue = .Contact_Name_LastFirst
6550  On Error GoTo ERRH
6560    End With

EXITP:
6570    Exit Sub

ERRH:
6580    Select Case ERR.Number
        Case Else
6590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6600    End Select
6610    Resume EXITP

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

Private Sub Contact_Address1_lbl_DblClick(Cancel As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_lbl_DblClick"

8910    SortNow THIS_PROC  ' ** Procedure: Below.

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

Private Sub Contact_Address1_Exit(Cancel As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address1_Exit"

9610    varOldValue = Null

EXITP:
9620    Exit Sub

ERRH:
9630    Select Case ERR.Number
        Case Else
9640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9650    End Select
9660    Resume EXITP

End Sub

Private Sub Contact_Address2_lbl_DblClick(Cancel As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_lbl_DblClick"

9710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
9720    Exit Sub

ERRH:
9730    Select Case ERR.Number
        Case Else
9740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9750    End Select
9760    Resume EXITP

End Sub

Private Sub Contact_Address2_Enter()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_Enter"

9810    With Me
9820      varOldValue = .Contact_Address2
9830    End With

EXITP:
9840    Exit Sub

ERRH:
9850    Select Case ERR.Number
        Case Else
9860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9870    End Select
9880    Resume EXITP

End Sub

Private Sub Contact_Address2_GotFocus()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_GotFocus"

9910    With Me
9920      .Contact_Address2.SelLength = 0
9930      .Contact_Address2.SelStart = 99
9940    End With

EXITP:
9950    Exit Sub

ERRH:
9960    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
9970    Case Else
9980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9990    End Select
10000   Resume EXITP

End Sub

Private Sub Contact_Address2_AfterUpdate()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_AfterUpdate"

        Dim strTmp01 As String

10110   With Me
10120     If IsNull(.Contact_Address2) = False Then
10130       If Trim(.Contact_Address2) <> vbNullString Then
10140         strTmp01 = Trim(FixQuotes(.Contact_Address2))  ' ** Module Function: modStringFuncs.
10150         If Compare_StringA_StringB(.Contact_Address2, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
10160           .Contact_Address2 = strTmp01
10170           .Contact_Changed = True
10180           cmdSave_Click  ' ** Procedure: Above.
10190         End If
10200       Else
10210         If IsNull(varOldValue) = False Then
10220           .Contact_Address2 = Null
10230           .Contact_Changed = True
10240           cmdSave_Click  ' ** Procedure: Above.
10250         End If
10260       End If
10270     Else
10280       If IsNull(varOldValue) = False Then
10290         .Contact_Changed = True
10300         cmdSave_Click  ' ** Procedure: Above.
10310       End If
10320     End If
10330   End With

EXITP:
10340   Exit Sub

ERRH:
10350   Select Case ERR.Number
        Case Else
10360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10370   End Select
10380   Resume EXITP

End Sub

Private Sub Contact_Address2_Exit(Cancel As Integer)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Address2_Exit"

10410   varOldValue = Null

EXITP:
10420   Exit Sub

ERRH:
10430   Select Case ERR.Number
        Case Else
10440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10450   End Select
10460   Resume EXITP

End Sub

Private Sub Contact_City_lbl_DblClick(Cancel As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_lbl_DblClick"

10510   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
10520   Exit Sub

ERRH:
10530   Select Case ERR.Number
        Case Else
10540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10550   End Select
10560   Resume EXITP

End Sub

Private Sub Contact_City_Enter()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_Enter"

10610   With Me
10620     varOldValue = .Contact_City
10630   End With

EXITP:
10640   Exit Sub

ERRH:
10650   Select Case ERR.Number
        Case Else
10660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10670   End Select
10680   Resume EXITP

End Sub

Private Sub Contact_City_GotFocus()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_GotFocus"

10710   With Me
10720     .Contact_City.SelLength = 0
10730     .Contact_City.SelStart = 99
10740   End With

EXITP:
10750   Exit Sub

ERRH:
10760   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
10770   Case Else
10780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10790   End Select
10800   Resume EXITP

End Sub

Private Sub Contact_City_AfterUpdate()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_AfterUpdate"

        Dim strTmp01 As String

10910   With Me
10920     If IsNull(.Contact_City) = False Then
10930       If Trim(.Contact_City) <> vbNullString Then
10940         strTmp01 = Trim(FixQuotes(.Contact_City))  ' ** Module Function: modStringFuncs.
10950         If Compare_StringA_StringB(.Contact_City, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
10960           .Contact_City = strTmp01
10970           .Contact_Changed = True
10980           cmdSave_Click  ' ** Procedure: Above.
10990         End If
11000       Else
11010         If IsNull(varOldValue) = False Then
11020           .Contact_City = Null
11030           .Contact_Changed = True
11040           cmdSave_Click  ' ** Procedure: Above.
11050         End If
11060       End If
11070     Else
11080       If IsNull(varOldValue) = False Then
11090         .Contact_Changed = True
11100         cmdSave_Click  ' ** Procedure: Above.
11110       End If
11120     End If
11130   End With

EXITP:
11140   Exit Sub

ERRH:
11150   Select Case ERR.Number
        Case Else
11160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11170   End Select
11180   Resume EXITP

End Sub

Private Sub Contact_City_KeyDown(KeyCode As Integer, Shift As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11220   intShiftDown = (Shift And acShiftMask) > 0
11230   intAltDown = (Shift And acAltMask) > 0
11240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11270       With Me
11280         intRetVal = 0
11290         Select Case .Parent.chkEnableCountry_Wide
              Case True
11300           Select Case IsNull(.Contact_Country)
                Case True
11310             .Contact_State.SetFocus
11320           Case False
11330             .Contact_Country.SetFocus
11340           End Select
11350         Case False
11360           .Contact_State.SetFocus
11370         End Select
11380       End With
11390     End Select
11400   End If

        ' ** Shift keys.
11410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11420     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11430       With Me
11440         intRetVal = 0
11450         .Contact_Address2.SetFocus
11460       End With
11470     End Select
11480   End If

EXITP:
11490   KeyCode = intRetVal
11500   Exit Sub

ERRH:
11510   intRetVal = 0
11520   Select Case ERR.Number
        Case Else
11530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11540   End Select
11550   Resume EXITP

End Sub

Private Sub Contact_City_Exit(Cancel As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_City_Exit"

11610   varOldValue = Null

EXITP:
11620   Exit Sub

ERRH:
11630   Select Case ERR.Number
        Case Else
11640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11650   End Select
11660   Resume EXITP

End Sub

Private Sub opgLocContact_lbl_DblClick(Cancel As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_lbl_DblClick"

11710   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
11720   Exit Sub

ERRH:
11730   Select Case ERR.Number
        Case Else
11740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11750   End Select
11760   Resume EXITP

End Sub

Private Sub opgLocContact_Enter()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_Enter"

11810   With Me
11820     varOldValue = .opgLocContact
11830   End With

EXITP:
11840   Exit Sub

ERRH:
11850   Select Case ERR.Number
        Case Else
11860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11870   End Select
11880   Resume EXITP

End Sub

Public Sub opgLocContact_AfterUpdate()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_AfterUpdate"

11910   With Me
11920     If .opgLocContact <> varOldValue Then
11930       Select Case .opgLocContact
            Case .opgLocContact_optUSA.OptionValue
11940         .opgLocContact_optUSA_lbl.FontBold = True
11950         .opgLocContact_optOther_lbl.FontBold = False
11960         If IsNull(.Contact_Country) = False Then
11970           .Contact_Country = Null
11980         End If
11990         If IsNull(.Contact_PostalCode) = False Then
12000           .Contact_PostalCode = Null
12010         End If
12020         varOldValue = .opgLocContact
12030         .Contact_Country.Enabled = False
12040         .Contact_Country.Locked = True
12050         .Contact_PostalCode.Enabled = False
12060         .Contact_PostalCode.Locked = True
12070         .Contact_State.Enabled = True
12080         .Contact_State.Locked = False
12090         .Contact_Zip.Enabled = True
12100         .Contact_Zip.Locked = False
12110         .Contact_Zip_Format.Enabled = True
12120         .Contact_Zip_Format.Locked = False
12130       Case .opgLocContact_optOther.OptionValue
12140         .opgLocContact_optUSA_lbl.FontBold = False
12150         .opgLocContact_optOther_lbl.FontBold = True
12160         If IsNull(.Contact_State) = False Then
12170           .Contact_State = Null
12180         End If
12190         If IsNull(.Contact_Zip) = False Then
12200           .Contact_Zip = Null
12210         End If
12220         If IsNull(.Contact_Zip_Format) = False Then
12230           .Contact_Zip_Format = Null
12240         End If
12250         varOldValue = .opgLocContact
12260         .Contact_Country.Enabled = True
12270         .Contact_Country.Locked = False
12280         .Contact_PostalCode.Enabled = True
12290         .Contact_PostalCode.Locked = False
12300         .Contact_State.Enabled = False
12310         .Contact_State.Locked = True
12320         .Contact_Zip.Enabled = False
12330         .Contact_Zip.Locked = True
12340         .Contact_Zip_Format.Enabled = False
12350         .Contact_Zip_Format.Locked = True
12360       End Select
12370       .Contact_Changed = True
12380       cmdSave_Click  ' ** Procedure: Above.
12390     End If
12400   End With

EXITP:
12410   Exit Sub

ERRH:
12420   Select Case ERR.Number
        Case Else
12430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12440   End Select
12450   Resume EXITP

End Sub

Private Sub opgLocContact_optUSA_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_optUSA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12520   intShiftDown = (Shift And acShiftMask) > 0
12530   intAltDown = (Shift And acAltMask) > 0
12540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12570       With Me
12580         intRetVal = 0
12590         Select Case .opgLocContact
              Case .opgLocContact_optUSA.OptionValue
12600           .Contact_State.SetFocus
12610         Case .opgLocContact_optOther.OptionValue
12620           .Contact_Country.SetFocus
12630         End Select
12640       End With
12650     End Select
12660   End If

        ' ** Shift keys.
12670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12690       With Me
12700         intRetVal = 0
12710         .Contact_City.SetFocus
12720       End With
12730     End Select
12740   End If

EXITP:
12750   KeyCode = intRetVal
12760   Exit Sub

ERRH:
12770   intRetVal = 0
12780   Select Case ERR.Number
        Case Else
12790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12800   End Select
12810   Resume EXITP

End Sub

Private Sub opgLocContact_optOther_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_optOther_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12920   intShiftDown = (Shift And acShiftMask) > 0
12930   intAltDown = (Shift And acAltMask) > 0
12940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12970       With Me
12980         intRetVal = 0
12990         Select Case .opgLocContact
              Case .opgLocContact_optUSA.OptionValue
13000           .Contact_State.SetFocus
13010         Case .opgLocContact_optOther.OptionValue
13020           .Contact_Country.SetFocus
13030         End Select
13040       End With
13050     End Select
13060   End If

        ' ** Shift keys.
13070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13090       With Me
13100         intRetVal = 0
13110         .Contact_City.SetFocus
13120       End With
13130     End Select
13140   End If

EXITP:
13150   KeyCode = intRetVal
13160   Exit Sub

ERRH:
13170   intRetVal = 0
13180   Select Case ERR.Number
        Case Else
13190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13200   End Select
13210   Resume EXITP

End Sub

Private Sub opgLocContact_Exit(Cancel As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocContact_Exit"

13310   varOldValue = Null

EXITP:
13320   Exit Sub

ERRH:
13330   Select Case ERR.Number
        Case Else
13340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13350   End Select
13360   Resume EXITP

End Sub

Private Sub Contact_State_lbl_DblClick(Cancel As Integer)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_lbl_DblClick"

13410   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
13420   Exit Sub

ERRH:
13430   Select Case ERR.Number
        Case Else
13440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13450   End Select
13460   Resume EXITP

End Sub

Private Sub Contact_State_Enter()

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_Enter"

13510   With Me
13520     varOldValue = .Contact_State
13530     If .Contact_State.text = vbNullString Then
13540       ExpandCombo .Contact_State  ' ** Module Procedure: modUtilities.
13550     End If
13560   End With

EXITP:
13570   Exit Sub

ERRH:
13580   Select Case ERR.Number
        Case Else
13590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13600   End Select
13610   Resume EXITP

End Sub

Private Sub Contact_State_GotFocus()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_GotFocus"

13710   With Me
13720     .Contact_State.SelLength = 0
13730     .Contact_State.SelStart = 9
13740   End With

EXITP:
13750   Exit Sub

ERRH:
13760   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
13770   Case Else
13780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13790   End Select
13800   Resume EXITP

End Sub

Private Sub Contact_State_AfterUpdate()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_AfterUpdate"

13910   With Me
13920     If IsNull(.Contact_State) = False Then
13930       If Trim(.Contact_State) <> vbNullString Then
13940         .Contact_State.Value = FormatUpperCase(Nz(.Contact_State, vbNullString))
13950         Select Case IsNull(.Parent.Contact_State_Pref)
              Case True
13960           .Parent.Contact_State_Pref = .Contact_State
13970         Case False
13980           If .Parent.Contact_State_Pref <> .Contact_State Then
13990             .Parent.Contact_State_Pref = .Contact_State
14000           End If
14010         End Select
14020         .Contact_Changed = True
14030         cmdSave_Click  ' ** Procedure: Above.
14040       Else
14050         If IsNull(varOldValue) = False Then
14060           .Contact_State = Null
14070           .Contact_Changed = True
14080           cmdSave_Click  ' ** Procedure: Above.
14090         End If
14100       End If
14110     Else
14120       If IsNull(varOldValue) = False Then
14130         .Contact_Changed = True
14140         cmdSave_Click  ' ** Procedure: Above.
14150       End If
14160     End If
14170   End With

EXITP:
14180   Exit Sub

ERRH:
14190   Select Case ERR.Number
        Case Else
14200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14210   End Select
14220   Resume EXITP

End Sub

Private Sub Contact_State_NotInList(NewData As String, Response As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_NotInList"

14310   With Me
14320     .Contact_State = Null
14330     MsgBox "Please select a valid state code.", vbInformation + vbOKOnly, "Invalid Entry"
14340     Response = acDataErrContinue
14350   End With

EXITP:
14360   Exit Sub

ERRH:
14370   Select Case ERR.Number
        Case Else
14380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14390   End Select
14400   Resume EXITP

End Sub

Private Sub Contact_State_Exit(Cancel As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_State_Exit"

14510   varOldValue = Null

EXITP:
14520   Exit Sub

ERRH:
14530   Select Case ERR.Number
        Case Else
14540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14550   End Select
14560   Resume EXITP

End Sub

Private Sub Contact_Zip_lbl_DblClick(Cancel As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_lbl_DblClick"

14610   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
14620   Exit Sub

ERRH:
14630   Select Case ERR.Number
        Case Else
14640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14650   End Select
14660   Resume EXITP

End Sub

Private Sub Contact_Zip_Enter()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Enter"

14710   With Me
14720     varOldValue = .Contact_Zip
14730   End With

EXITP:
14740   Exit Sub

ERRH:
14750   Select Case ERR.Number
        Case Else
14760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14770   End Select
14780   Resume EXITP

End Sub

Private Sub Contact_Zip_GotFocus()

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_GotFocus"

14810   With Me
14820     .Contact_Zip.SelLength = 0
14830     .Contact_Zip.SelStart = 19
14840   End With

EXITP:
14850   Exit Sub

ERRH:
14860   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
14870   Case Else
14880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14890   End Select
14900   Resume EXITP

End Sub

Private Sub Contact_Zip_AfterUpdate()
' ** This field has no InputMask.
' ** Canadian Zip codes are entered in Contact_PostalCode.

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_AfterUpdate"

        Dim varZip As Variant
        Dim intPos01 As Integer
        Dim strTmp01 As String

15010   With Me
15020     If IsNull(.Contact_Zip) = False Then
15030       If Trim(.Contact_Zip) <> vbNullString Then
15040         strTmp01 = Trim(.Contact_Zip)
15050         intPos01 = InStr(strTmp01, "-")
15060         If intPos01 > 0 Then
15070           If Len(strTmp01) > intPos01 Then
15080             strTmp01 = Left(strTmp01, (intPos01 - 1)) & Mid(strTmp01, (intPos01 + 1))
15090             .Contact_Zip = strTmp01
15100           Else
15110             strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
15120             .Contact_Zip = strTmp01
15130           End If
15140           If IsNumeric(strTmp01) = False Or (Len(strTmp01) <> 5 And Len(strTmp01) <> 9) Then
15150             strTmp01 = vbNullString
15160             .Contact_Zip = Null
15170           End If
15180         End If
15190         .Contact_Changed = True

15200         cmdSave_Click  ' ** Procedure: Above.

15210         varZip = Null

15220         If strTmp01 <> vbNullString Then
15230           Select Case Len(strTmp01)
                Case 5
15240             varZip = strTmp01
15250           Case 9
15260             varZip = FormatZip9(strTmp01)  ' ** Module Function: modStringFuncs.
15270           End Select
15280           .Contact_Zip_Format = varZip
15290           .Contact_Changed = True
15300           cmdSave_Click  ' ** Procedure: Above.
15310         Else
15320           If IsNull(varOldValue) = False Then
15330             .Contact_Zip = Null
15340             .Contact_Changed = True
15350             cmdSave_Click  ' ** Procedure: Above.
15360           End If
15370           If IsNull(.Contact_Zip_Format) = False Then
15380             .Contact_Zip_Format = Null
15390             .Contact_Changed = True
15400             cmdSave_Click  ' ** Procedure: Above.
15410           End If
15420         End If
15430       Else
15440         If IsNull(varOldValue) = False Then
15450           .Contact_Zip = Null
15460           .Contact_Changed = True
15470           cmdSave_Click  ' ** Procedure: Above.
15480         End If
15490         If IsNull(.Contact_Zip_Format) = False Then
15500           .Contact_Zip_Format = Null
15510           .Contact_Changed = True
15520           cmdSave_Click  ' ** Procedure: Above.
15530         End If
15540       End If
15550     Else
15560       If IsNull(varOldValue) = False Then
15570         .Contact_Changed = True
15580         cmdSave_Click  ' ** Procedure: Above.
15590       End If
15600       If IsNull(.Contact_Zip_Format) = False Then
15610         .Contact_Zip_Format = Null
15620         .Contact_Changed = True
15630         cmdSave_Click  ' ** Procedure: Above.
15640       End If
15650     End If
15660   End With

EXITP:
15670   Exit Sub

ERRH:
15680   Select Case ERR.Number
        Case Else
15690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15700   End Select
15710   Resume EXITP

End Sub

Private Sub Contact_Zip_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15820   intShiftDown = (Shift And acShiftMask) > 0
15830   intAltDown = (Shift And acAltMask) > 0
15840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15870       With Me
15880         intRetVal = 0
15890         Select Case .Parent.chkEnableCountry_Wide
              Case True
15900           If IsNull(.Contact_State) = True And IsNull(.Contact_Zip) = True And IsNull(.Contact_Zip_Format) = True Then
15910             .Contact_Country.SetFocus
15920           Else
15930             If .Contact_Phone1.Visible = True And .Contact_Phone1.Enabled = True Then
15940               .Contact_Phone1.SetFocus
15950             ElseIf .Contact_Phone1_Format.Visible = True And .Contact_Phone1_Format.Enabled = True Then
15960               .Contact_Phone1_Format.SetFocus
15970             End If
15980           End If
15990         Case False
16000           If .Contact_Phone1.Visible = True And .Contact_Phone1.Enabled = True Then
16010             .Contact_Phone1.SetFocus
16020           ElseIf .Contact_Phone1_Format.Visible = True And .Contact_Phone1_Format.Enabled = True Then
16030             .Contact_Phone1_Format.SetFocus
16040           End If
16050         End Select
16060       End With
16070     End Select
16080   End If

        ' ** Shift keys.
16090   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16100     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16110       With Me
16120         intRetVal = 0
16130         .Contact_State.SetFocus
16140       End With
16150     End Select
16160   End If

EXITP:
16170   KeyCode = intRetVal
16180   Exit Sub

ERRH:
16190   intRetVal = 0
16200   Select Case ERR.Number
        Case Else
16210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16220   End Select
16230   Resume EXITP

End Sub

Private Sub Contact_Zip_Exit(Cancel As Integer)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Exit"

16310   varOldValue = Null

EXITP:
16320   Exit Sub

ERRH:
16330   Select Case ERR.Number
        Case Else
16340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16350   End Select
16360   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_Enter()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_Enter"

16410   With Me
16420     varOldValue = .Contact_Zip_Format
16430   End With

EXITP:
16440   Exit Sub

ERRH:
16450   Select Case ERR.Number
        Case Else
16460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16470   End Select
16480   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_GotFocus()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_GotFocus"

16510   With Me
16520     .Contact_Zip_Format.SelLength = 0
16530     .Contact_Zip_Format.SelStart = 19
16540   End With

EXITP:
16550   Exit Sub

ERRH:
16560   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
16570   Case Else
16580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16590   End Select
16600   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_AfterUpdate()
' ** This field has an InputMask: 00000"-9999";;_
' ** Canadian Zip codes are entered in Contact_PostalCode.

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String

16710   With Me
16720     If IsNull(.Contact_Zip_Format) = False Then
16730       If Trim(.Contact_Zip_Format) <> vbNullString Then
16740         If Right(.Contact_Zip_Format.text, 4) <> "____" And Right(.Contact_Zip_Format.text, 4) <> "    " And _
                  (Not IsNumeric(Mid(.Contact_Zip_Format.text, 7, 1)) Or Not IsNumeric(Mid(.Contact_Zip_Format.text, 8, 1)) Or _
                  Not IsNumeric(Mid(.Contact_Zip_Format.text, 9, 1)) Or Not IsNumeric(Mid(.Contact_Zip_Format.text, 10, 1))) Then
16750           MsgBox "Please enter a valid zip code in the form 00000 or 00000-0000.", vbInformation + vbOKOnly, "Invalid Entry"
16760         Else
16770           .Contact_Changed = True

16780           cmdSave_Click  ' ** Procedure: Above.

16790           strTmp01 = Trim(.Contact_Zip_Format)
16800           Select Case Len(strTmp01)
                Case 5
                  ' ** strTmp01 stands.
16810           Case 10
16820             intPos01 = InStr(strTmp01, "-")
16830             strTmp01 = Left(strTmp01, (intPos01 - 1)) & Mid(strTmp01, (intPos01 + 1))
16840           End Select
16850           .Contact_Zip = strTmp01
16860           cmdSave_Click  ' ** Procedure: Above.
16870         End If
16880       Else
16890         If IsNull(varOldValue) = False Then
16900           .Contact_Zip_Format = Null
16910           .Contact_Changed = True
16920           cmdSave_Click  ' ** Procedure: Above.
16930         End If
16940         If IsNull(.Contact_Zip) = False Then
16950           .Contact_Zip = Null
16960           .Contact_Changed = True
16970           cmdSave_Click  ' ** Procedure: Above.
16980         End If
16990       End If
17000     Else
17010       If IsNull(varOldValue) = False Then
17020         .Contact_Changed = True
17030         cmdSave_Click  ' ** Procedure: Above.
17040       End If
17050       If IsNull(.Contact_Zip) = False Then
17060         .Contact_Zip = Null
17070         .Contact_Changed = True
17080         cmdSave_Click  ' ** Procedure: Above.
17090       End If
17100     End If
17110   End With

EXITP:
17120   Exit Sub

ERRH:
17130   Select Case ERR.Number
        Case Else
17140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17150   End Select
17160   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_KeyDown(KeyCode As Integer, Shift As Integer)

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17220   intShiftDown = (Shift And acShiftMask) > 0
17230   intAltDown = (Shift And acAltMask) > 0
17240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17270       With Me
17280         intRetVal = 0
17290         Select Case .Parent.chkEnableCountry_Wide
              Case True
17300           If IsNull(.Contact_State) = True And IsNull(.Contact_Zip) = True And IsNull(.Contact_Zip_Format) = True Then
17310             .Contact_Country.SetFocus
17320           Else
17330             If .Contact_Phone1.Visible = True And .Contact_Phone1.Enabled = True Then
17340               .Contact_Phone1.SetFocus
17350             ElseIf .Contact_Phone1_Format.Visible = True And .Contact_Phone1_Format.Enabled = True Then
17360               .Contact_Phone1_Format.SetFocus
17370             End If
17380           End If
17390         Case False
17400           If .Contact_Phone1.Visible = True And .Contact_Phone1.Enabled = True Then
17410             .Contact_Phone1.SetFocus
17420           ElseIf .Contact_Phone1_Format.Visible = True And .Contact_Phone1_Format.Enabled = True Then
17430             .Contact_Phone1_Format.SetFocus
17440           End If
17450         End Select
17460       End With
17470     End Select
17480   End If

        ' ** Shift keys.
17490   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17500     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17510       With Me
17520         intRetVal = 0
17530         .Contact_State.SetFocus
17540       End With
17550     End Select
17560   End If

EXITP:
17570   KeyCode = intRetVal
17580   Exit Sub

ERRH:
17590   intRetVal = 0
17600   Select Case ERR.Number
        Case Else
17610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17620   End Select
17630   Resume EXITP

End Sub

Private Sub Contact_Zip_Format_Exit(Cancel As Integer)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Zip_Format_Exit"

17710   varOldValue = Null

EXITP:
17720   Exit Sub

ERRH:
17730   Select Case ERR.Number
        Case Else
17740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17750   End Select
17760   Resume EXITP

End Sub

Private Sub Contact_Country_lbl_DblClick(Cancel As Integer)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_lbl_DblClick"

17810   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
17820   Exit Sub

ERRH:
17830   Select Case ERR.Number
        Case Else
17840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17850   End Select
17860   Resume EXITP

End Sub

Private Sub Contact_Country_Enter()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_Enter"

17910   With Me
17920     varOldValue = .Contact_Country
17930   End With

EXITP:
17940   Exit Sub

ERRH:
17950   Select Case ERR.Number
        Case Else
17960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17970   End Select
17980   Resume EXITP

End Sub

Private Sub Contact_Country_GotFocus()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_GotFocus"

18010   With Me
18020     .Contact_Country.SelLength = 0
18030     .Contact_Country.SelStart = 30
18040   End With

EXITP:
18050   Exit Sub

ERRH:
18060   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
18070   Case Else
18080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18090   End Select
18100   Resume EXITP

End Sub

Private Sub Contact_Country_AfterUpdate()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_AfterUpdate"

        Dim strTmp01 As String

18210   With Me
18220     If IsNull(.Contact_Country) = False Then
18230       If Trim(.Contact_Country) <> vbNullString Then
18240         strTmp01 = Trim(FixQuotes(.Contact_Country))  ' ** Module Function: modStringFuncs.
18250         If Compare_StringA_StringB(.Contact_Country, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
18260           .Contact_Country = strTmp01
18270           .Contact_Changed = True
18280           cmdSave_Click  ' ** Procedure: Above.
18290         End If
18300       Else
18310         If IsNull(varOldValue) = False Then
18320           .Contact_Country = Null
18330           .Contact_Changed = True
18340           cmdSave_Click  ' ** Procedure: Above.
18350         End If
18360       End If
18370     Else
18380       If IsNull(varOldValue) = False Then
18390         .Contact_Changed = True
18400         cmdSave_Click  ' ** Procedure: Above.
18410       End If
18420     End If
18430   End With

EXITP:
18440   Exit Sub

ERRH:
18450   Select Case ERR.Number
        Case Else
18460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18470   End Select
18480   Resume EXITP

End Sub

Private Sub Contact_Country_KeyDown(KeyCode As Integer, Shift As Integer)

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18520   intShiftDown = (Shift And acShiftMask) > 0
18530   intAltDown = (Shift And acAltMask) > 0
18540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18570       With Me
18580         intRetVal = 0
18590         .Contact_PostalCode.SetFocus
18600       End With
18610     End Select
18620   End If

        ' ** Shift keys.
18630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18650       With Me
18660         intRetVal = 0
18670         If IsNull(.Contact_Country) = True And IsNull(.Contact_PostalCode) = True Then
18680           If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
18690             .Contact_Zip.SetFocus
18700           ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
18710             .Contact_Zip_Format.SetFocus
18720           End If
18730         Else
18740           .Contact_City.SetFocus
18750         End If
18760       End With
18770     End Select
18780   End If

EXITP:
18790   KeyCode = intRetVal
18800   Exit Sub

ERRH:
18810   intRetVal = 0
18820   Select Case ERR.Number
        Case Else
18830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18840   End Select
18850   Resume EXITP

End Sub

Private Sub Contact_Country_Exit(Cancel As Integer)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Country_Exit"

18910   varOldValue = Null

EXITP:
18920   Exit Sub

ERRH:
18930   Select Case ERR.Number
        Case Else
18940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18950   End Select
18960   Resume EXITP

End Sub

Private Sub Contact_PostalCode_lbl_DblClick(Cancel As Integer)

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_lbl_DblClick"

19010   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19020   Exit Sub

ERRH:
19030   Select Case ERR.Number
        Case Else
19040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19050   End Select
19060   Resume EXITP

End Sub

Private Sub Contact_PostalCode_Enter()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_Enter"

19110   With Me
19120     varOldValue = .Contact_PostalCode
19130   End With

EXITP:
19140   Exit Sub

ERRH:
19150   Select Case ERR.Number
        Case Else
19160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19170   End Select
19180   Resume EXITP

End Sub

Private Sub Contact_PostalCode_GotFocus()

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_GotFocus"

19210   With Me
19220     .Contact_PostalCode.SelLength = 0
19230     .Contact_PostalCode.SelStart = 10
19240   End With

EXITP:
19250   Exit Sub

ERRH:
19260   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
19270   Case Else
19280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19290   End Select
19300   Resume EXITP

End Sub

Private Sub Contact_PostalCode_AfterUpdate()

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_AfterUpdate"

        Dim strTmp01 As String

19410   With Me
19420     If IsNull(.Contact_PostalCode) = False Then
19430       If Trim(.Contact_PostalCode) <> vbNullString Then
19440         strTmp01 = Trim(FixQuotes(.Contact_PostalCode))  ' ** Module Function: modStringFuncs.
19450         If Compare_StringA_StringB(.Contact_PostalCode, "=", strTmp01) = False Then  ' ** Module Function: modStringFuncs.
19460           .Contact_PostalCode = strTmp01
19470           .Contact_Changed = True
19480           cmdSave_Click  ' ** Procedure: Above.
19490         End If
19500       Else
19510         If IsNull(varOldValue) = False Then
19520           .Contact_PostalCode = Null
19530           .Contact_Changed = True
19540           cmdSave_Click  ' ** Procedure: Above.
19550         End If
19560       End If
19570     Else
19580       If IsNull(varOldValue) = False Then
19590         .Contact_Changed = True
19600         cmdSave_Click  ' ** Procedure: Above.
19610       End If
19620     End If
19630   End With

EXITP:
19640   Exit Sub

ERRH:
19650   Select Case ERR.Number
        Case Else
19660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19670   End Select
19680   Resume EXITP

End Sub

Private Sub Contact_PostalCode_Exit(Cancel As Integer)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_PostalCode_Exit"

19710   varOldValue = Null

EXITP:
19720   Exit Sub

ERRH:
19730   Select Case ERR.Number
        Case Else
19740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19750   End Select
19760   Resume EXITP

End Sub

Private Sub Contact_Phone1_lbl_DblClick(Cancel As Integer)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_lbl_DblClick"

19810   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
19820   Exit Sub

ERRH:
19830   Select Case ERR.Number
        Case Else
19840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19850   End Select
19860   Resume EXITP

End Sub

Private Sub Contact_Phone1_Enter()

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Enter"

19910   With Me
19920     varOldValue = .Contact_Phone1
19930   End With

EXITP:
19940   Exit Sub

ERRH:
19950   Select Case ERR.Number
        Case Else
19960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19970   End Select
19980   Resume EXITP

End Sub

Private Sub Contact_Phone1_GotFocus()

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_GotFocus"

20010   With Me
20020     .Contact_Phone1.SelLength = 0
20030     .Contact_Phone1.SelStart = 19
20040   End With

EXITP:
20050   Exit Sub

ERRH:
20060   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
20070   Case Else
20080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20090   End Select
20100   Resume EXITP

End Sub

Private Sub Contact_Phone1_AfterUpdate()
' ** This field has no InputMask.

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_AfterUpdate"

        Dim strTmp01 As String

20210   With Me
20220     If IsNull(.Contact_Phone1) = False Then
20230       If Trim(.Contact_Phone1) <> vbNullString Then
20240         strTmp01 = Trim(.Contact_Phone1)
20250         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
20260         If strTmp01 <> .Contact_Phone1 Then
20270           .Contact_Phone1 = strTmp01
20280         End If
20290         .Contact_Changed = True

20300         cmdSave_Click  ' ** Procedure: Above.

20310         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
20320         .Contact_Phone1_Format = strTmp01
20330         cmdSave_Click  ' ** Procedure: Above.
20340       Else
20350         If IsNull(varOldValue) = False Then
20360           .Contact_Phone1 = Null
20370           .Contact_Changed = True
20380           cmdSave_Click  ' ** Procedure: Above.
20390         End If
20400         If IsNull(.Contact_Phone1_Format) = False Then
20410           .Contact_Phone1_Format = Null
20420           .Contact_Changed = True
20430           cmdSave_Click  ' ** Procedure: Above.
20440         End If
20450       End If
20460     Else
20470       If IsNull(varOldValue) = False Then
20480         .Contact_Changed = True
20490         cmdSave_Click  ' ** Procedure: Above.
20500       End If
20510       If IsNull(.Contact_Phone1_Format) = False Then
20520         .Contact_Phone1_Format = Null
20530         .Contact_Changed = True
20540         cmdSave_Click  ' ** Procedure: Above.
20550       End If
20560     End If
20570   End With

EXITP:
20580   Exit Sub

ERRH:
20590   Select Case ERR.Number
        Case Else
20600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20610   End Select
20620   Resume EXITP

End Sub

Private Sub Contact_Phone1_KeyDown(KeyCode As Integer, Shift As Integer)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_KeyDown"

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
20790         If .Contact_Phone2.Visible = True And .Contact_Phone2.Enabled = True Then
20800           .Contact_Phone2.SetFocus
20810         ElseIf .Contact_Phone2_Format.Visible = True And .Contact_Phone2_Format.Enabled = True Then
20820           .Contact_Phone2_Format.SetFocus
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
20910         Select Case .Parent.chkEnableCountry_Wide
              Case True
20920           If IsNull(.Contact_State) = True And IsNull(.Contact_Zip) = True And IsNull(.Contact_Zip_Format) = True Then
20930             .Contact_PostalCode.SetFocus
20940           Else
20950             If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
20960               .Contact_Zip.SetFocus
20970             ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
20980               .Contact_Zip_Format.SetFocus
20990             End If
21000           End If
21010         Case False
21020           If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
21030             .Contact_Zip.SetFocus
21040           ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
21050             .Contact_Zip_Format.SetFocus
21060           End If
21070         End Select
21080       End With
21090     End Select
21100   End If

EXITP:
21110   KeyCode = intRetVal
21120   Exit Sub

ERRH:
21130   intRetVal = 0
21140   Select Case ERR.Number
        Case Else
21150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21160   End Select
21170   Resume EXITP

End Sub

Private Sub Contact_Phone1_Exit(Cancel As Integer)

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Exit"

21210   varOldValue = Null

EXITP:
21220   Exit Sub

ERRH:
21230   Select Case ERR.Number
        Case Else
21240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21250   End Select
21260   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_Enter()

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_Enter"

21310   With Me
21320     varOldValue = .Contact_Phone1_Format
21330   End With

EXITP:
21340   Exit Sub

ERRH:
21350   Select Case ERR.Number
        Case Else
21360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21370   End Select
21380   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_GotFocus()

21400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_GotFocus"

21410   With Me
21420     .Contact_Phone1_Format.SelLength = 0
21430     .Contact_Phone1_Format.SelStart = 19
21440   End With

EXITP:
21450   Exit Sub

ERRH:
21460   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
21470   Case Else
21480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21490   End Select
21500   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_AfterUpdate()
' ** This field has an InputMask: !"(000) 000-0000";;_
'FOREIGN PHONE NUMS NEED COUNTRY CODE, ETC.!

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_AfterUpdate"

        Dim strTmp01 As String

21610   With Me
21620     If IsNull(.Contact_Phone1_Format) = False Then
21630       If Trim(.Contact_Phone1_Format) <> vbNullString Then
21640         strTmp01 = Trim(.Contact_Phone1)
21650         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
21660         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
21670         If strTmp01 <> .Contact_Phone1_Format Then
21680           .Contact_Phone1_Format = strTmp01
21690         End If
21700         .Contact_Changed = True

21710         cmdSave_Click  ' ** Procedure: Above.

21720         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
21730         .Contact_Phone1 = strTmp01
21740         cmdSave_Click  ' ** Procedure: Above.
21750       Else
21760         If IsNull(varOldValue) = False Then
21770           .Contact_Phone1_Format = Null
21780           .Contact_Changed = True
21790           cmdSave_Click  ' ** Procedure: Above.
21800         End If
21810         If IsNull(.Contact_Phone1) = False Then
21820           .Contact_Phone1 = Null
21830           .Contact_Changed = True
21840           cmdSave_Click  ' ** Procedure: Above.
21850         End If
21860       End If
21870     Else
21880       If IsNull(varOldValue) = False Then
21890         .Contact_Changed = True
21900         cmdSave_Click  ' ** Procedure: Above.
21910       End If
21920       If IsNull(.Contact_Phone1) = False Then
21930         .Contact_Phone1 = Null
21940         .Contact_Changed = True
21950         cmdSave_Click  ' ** Procedure: Above.
21960       End If
21970     End If
21980   End With

EXITP:
21990   Exit Sub

ERRH:
22000   Select Case ERR.Number
        Case Else
22010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22020   End Select
22030   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_KeyDown(KeyCode As Integer, Shift As Integer)

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_KeyDown"

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
22190         If .Contact_Phone2.Visible = True And .Contact_Phone2.Enabled = True Then
22200           .Contact_Phone2.SetFocus
22210         ElseIf .Contact_Phone2_Format.Visible = True And .Contact_Phone2_Format.Enabled = True Then
22220           .Contact_Phone2_Format.SetFocus
22230         End If
22240       End With
22250     End Select
22260   End If

        ' ** Shift keys.
22270   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22280     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22290       With Me
22300         intRetVal = 0
22310         Select Case .Parent.chkEnableCountry_Wide
              Case True
22320           If IsNull(.Contact_State) = True And IsNull(.Contact_Zip) = True And IsNull(.Contact_Zip_Format) = True Then
22330             .Contact_PostalCode.SetFocus
22340           Else
22350             If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
22360               .Contact_Zip.SetFocus
22370             ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
22380               .Contact_Zip_Format.SetFocus
22390             End If
22400           End If
22410         Case False
22420           If .Contact_Zip.Visible = True And .Contact_Zip.Enabled = True Then
22430             .Contact_Zip.SetFocus
22440           ElseIf .Contact_Zip_Format.Visible = True And .Contact_Zip_Format.Enabled = True Then
22450             .Contact_Zip_Format.SetFocus
22460           End If
22470         End Select
22480       End With
22490     End Select
22500   End If

EXITP:
22510   KeyCode = intRetVal
22520   Exit Sub

ERRH:
22530   intRetVal = 0
22540   Select Case ERR.Number
        Case Else
22550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22560   End Select
22570   Resume EXITP

End Sub

Private Sub Contact_Phone1_Format_Exit(Cancel As Integer)

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone1_Format_Exit"

22610   varOldValue = Null

EXITP:
22620   Exit Sub

ERRH:
22630   Select Case ERR.Number
        Case Else
22640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22650   End Select
22660   Resume EXITP

End Sub

Private Sub Contact_Phone2_lbl_DblClick(Cancel As Integer)

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_lbl_DblClick"

22710   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
22720   Exit Sub

ERRH:
22730   Select Case ERR.Number
        Case Else
22740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22750   End Select
22760   Resume EXITP

End Sub

Private Sub Contact_Phone2_Enter()

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Enter"

22810   With Me
22820     varOldValue = .Contact_Phone2
22830   End With

EXITP:
22840   Exit Sub

ERRH:
22850   Select Case ERR.Number
        Case Else
22860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22870   End Select
22880   Resume EXITP

End Sub

Private Sub Contact_Phone2_GotFocus()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_GotFocus"

22910   With Me
22920     .Contact_Phone2.SelLength = 0
22930     .Contact_Phone2.SelStart = 19
22940   End With

EXITP:
22950   Exit Sub

ERRH:
22960   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
22970   Case Else
22980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22990   End Select
23000   Resume EXITP

End Sub

Private Sub Contact_Phone2_AfterUpdate()
' ** This field has no InputMask.

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_AfterUpdate"

        Dim strTmp01 As String

23110   With Me
23120     If IsNull(.Contact_Phone2) = False Then
23130       If Trim(.Contact_Phone2) <> vbNullString Then
23140         strTmp01 = Trim(.Contact_Phone2)
23150         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
23160         If strTmp01 <> .Contact_Phone2 Then
23170           .Contact_Phone2 = strTmp01
23180         End If
23190         .Contact_Changed = True

23200         cmdSave_Click  ' ** Procedure: Above.

23210         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
23220         .Contact_Phone2_Format = strTmp01
23230         cmdSave_Click  ' ** Procedure: Above.
23240       Else
23250         If IsNull(varOldValue) = False Then
23260           .Contact_Phone2 = Null
23270           .Contact_Changed = True
23280           cmdSave_Click  ' ** Procedure: Above.
23290         End If
23300         If IsNull(.Contact_Phone2_Format) = False Then
23310           .Contact_Phone2_Format = Null
23320           .Contact_Changed = True
23330           cmdSave_Click  ' ** Procedure: Above.
23340         End If
23350       End If
23360     Else
23370       If IsNull(varOldValue) = False Then
23380         .Contact_Changed = True
23390         cmdSave_Click  ' ** Procedure: Above.
23400       End If
23410       If IsNull(.Contact_Phone2_Format) = False Then
23420         .Contact_Phone2_Format = Null
23430         .Contact_Changed = True
23440         cmdSave_Click  ' ** Procedure: Above.
23450       End If
23460     End If
23470   End With

EXITP:
23480   Exit Sub

ERRH:
23490   Select Case ERR.Number
        Case Else
23500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23510   End Select
23520   Resume EXITP

End Sub

Private Sub Contact_Phone2_Exit(Cancel As Integer)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Exit"

23610   varOldValue = Null

EXITP:
23620   Exit Sub

ERRH:
23630   Select Case ERR.Number
        Case Else
23640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23650   End Select
23660   Resume EXITP

End Sub

Private Sub Contact_Phone2_Format_Enter()

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Format_Enter"

23710   With Me
23720     varOldValue = .Contact_Phone2_Format
23730   End With

EXITP:
23740   Exit Sub

ERRH:
23750   Select Case ERR.Number
        Case Else
23760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23770   End Select
23780   Resume EXITP

End Sub

Private Sub Contact_Phone2_Format_GotFocus()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Format_GotFocus"

23810   With Me
23820     .Contact_Phone2_Format.SelLength = 0
23830     .Contact_Phone2_Format.SelStart = 19
23840   End With

EXITP:
23850   Exit Sub

ERRH:
23860   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
23870   Case Else
23880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23890   End Select
23900   Resume EXITP

End Sub

Private Sub Contact_Phone2_Format_AfterUpdate()
' ** This field has an InputMask: !"(000) 000-0000";;_
'FOREIGN PHONE NUMS NEED COUNTRY CODE, ETC.!

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Format_AfterUpdate"

        Dim strTmp01 As String

24010   With Me
24020     If IsNull(.Contact_Phone2_Format) = False Then
24030       If Trim(.Contact_Phone2_Format) <> vbNullString Then
24040         strTmp01 = Trim(.Contact_Phone2)
24050         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
24060         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
24070         If strTmp01 <> .Contact_Phone2_Format Then
24080           .Contact_Phone2_Format = strTmp01
24090         End If
24100         .Contact_Changed = True

24110         cmdSave_Click  ' ** Procedure: Above.

24120         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
24130         .Contact_Phone2 = strTmp01
24140         cmdSave_Click  ' ** Procedure: Above.
24150       Else
24160         If IsNull(varOldValue) = False Then
24170           .Contact_Phone2_Format = Null
24180           .Contact_Changed = True
24190           cmdSave_Click  ' ** Procedure: Above.
24200         End If
24210         If IsNull(.Contact_Phone2) = False Then
24220           .Contact_Phone2 = Null
24230           .Contact_Changed = True
24240           cmdSave_Click  ' ** Procedure: Above.
24250         End If
24260       End If
24270     Else
24280       If IsNull(varOldValue) = False Then
24290         .Contact_Changed = True
24300         cmdSave_Click  ' ** Procedure: Above.
24310       End If
24320       If IsNull(.Contact_Phone2) = False Then
24330         .Contact_Phone2 = Null
24340         .Contact_Changed = True
24350         cmdSave_Click  ' ** Procedure: Above.
24360       End If
24370     End If
24380   End With

EXITP:
24390   Exit Sub

ERRH:
24400   Select Case ERR.Number
        Case Else
24410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24420   End Select
24430   Resume EXITP

End Sub

Private Sub Contact_Phone2_Format_Exit(Cancel As Integer)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Phone2_Format_Exit"

24510   varOldValue = Null

EXITP:
24520   Exit Sub

ERRH:
24530   Select Case ERR.Number
        Case Else
24540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24550   End Select
24560   Resume EXITP

End Sub

Private Sub Contact_Fax_lbl_DblClick(Cancel As Integer)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_lbl_DblClick"

24610   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
24620   Exit Sub

ERRH:
24630   Select Case ERR.Number
        Case Else
24640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24650   End Select
24660   Resume EXITP

End Sub

Private Sub Contact_Fax_Enter()

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Enter"

24710   With Me
24720     varOldValue = .Contact_Fax
24730   End With

EXITP:
24740   Exit Sub

ERRH:
24750   Select Case ERR.Number
        Case Else
24760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24770   End Select
24780   Resume EXITP

End Sub

Private Sub Contact_Fax_GotFocus()

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_GotFocus"

24810   With Me
24820     .Contact_Fax.SelLength = 0
24830     .Contact_Fax.SelStart = 19
24840   End With

EXITP:
24850   Exit Sub

ERRH:
24860   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
24870   Case Else
24880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24890   End Select
24900   Resume EXITP

End Sub

Private Sub Contact_Fax_AfterUpdate()
' ** This field has no InputMask.

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_AfterUpdate"

        Dim strTmp01 As String

25010   With Me
25020     If IsNull(.Contact_Fax) = False Then
25030       If Trim(.Contact_Fax) <> vbNullString Then
25040         strTmp01 = Trim(.Contact_Fax)
25050         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
25060         If strTmp01 <> .Contact_Fax Then
25070           .Contact_Fax = strTmp01
25080         End If
25090         .Contact_Changed = True

25100         cmdSave_Click  ' ** Procedure: Above.

25110         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
25120         .Contact_Fax_Format = strTmp01
25130         cmdSave_Click  ' ** Procedure: Above.
25140       Else
25150         If IsNull(varOldValue) = False Then
25160           .Contact_Fax = Null
25170           .Contact_Changed = True
25180           cmdSave_Click  ' ** Procedure: Above.
25190         End If
25200         If IsNull(.Contact_Fax_Format) = False Then
25210           .Contact_Fax_Format = Null
25220           .Contact_Changed = True
25230           cmdSave_Click  ' ** Procedure: Above.
25240         End If
25250       End If
25260     Else
25270       If IsNull(varOldValue) = False Then
25280         .Contact_Changed = True
25290         cmdSave_Click  ' ** Procedure: Above.
25300       End If
25310       If IsNull(.Contact_Fax_Format) = False Then
25320         .Contact_Fax_Format = Null
25330         .Contact_Changed = True
25340         cmdSave_Click  ' ** Procedure: Above.
25350       End If
25360     End If
25370   End With

EXITP:
25380   Exit Sub

ERRH:
25390   Select Case ERR.Number
        Case Else
25400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25410   End Select
25420   Resume EXITP

End Sub

Private Sub Contact_Fax_Exit(Cancel As Integer)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Exit"

25510   varOldValue = Null

EXITP:
25520   Exit Sub

ERRH:
25530   Select Case ERR.Number
        Case Else
25540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25550   End Select
25560   Resume EXITP

End Sub

Private Sub Contact_Fax_Format_Enter()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Format_Enter"

25610   With Me
25620     varOldValue = .Contact_Fax_Format
25630   End With

EXITP:
25640   Exit Sub

ERRH:
25650   Select Case ERR.Number
        Case Else
25660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25670   End Select
25680   Resume EXITP

End Sub

Private Sub Contact_Fax_Format_GotFocus()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Format_GotFocus"

25710   With Me
25720     .Contact_Fax_Format.SelLength = 0
25730     .Contact_Fax_Format.SelStart = 19
25740   End With

EXITP:
25750   Exit Sub

ERRH:
25760   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
25770   Case Else
25780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25790   End Select
25800   Resume EXITP

End Sub

Private Sub Contact_Fax_Format_AfterUpdate()
' ** This field has an InputMask: !"(000) 000-0000";;_
'FOREIGN PHONE NUMS NEED COUNTRY CODE, ETC.!

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Format_AfterUpdate"

        Dim strTmp01 As String

25910   With Me
25920     If IsNull(.Contact_Fax_Format) = False Then
25930       If Trim(.Contact_Fax_Format) <> vbNullString Then
25940         strTmp01 = Trim(.Contact_Fax)
25950         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
25960         strTmp01 = FormatPhoneNum(strTmp01, "Dash")
25970         If strTmp01 <> .Contact_Fax_Format Then
25980           .Contact_Fax_Format = strTmp01
25990         End If
26000         .Contact_Changed = True

26010         cmdSave_Click  ' ** Procedure: Above.

26020         strTmp01 = FormatPhoneNum(strTmp01, "Rem")
26030         .Contact_Fax = strTmp01
26040         cmdSave_Click  ' ** Procedure: Above.
26050       Else
26060         If IsNull(varOldValue) = False Then
26070           .Contact_Fax_Format = Null
26080           .Contact_Changed = True
26090           cmdSave_Click  ' ** Procedure: Above.
26100         End If
26110         If IsNull(.Contact_Fax) = False Then
26120           .Contact_Fax = Null
26130           .Contact_Changed = True
26140           cmdSave_Click  ' ** Procedure: Above.
26150         End If
26160       End If
26170     Else
26180       If IsNull(varOldValue) = False Then
26190         .Contact_Changed = True
26200         cmdSave_Click  ' ** Procedure: Above.
26210       End If
26220       If IsNull(.Contact_Fax) = False Then
26230         .Contact_Fax = Null
26240         .Contact_Changed = True
26250         cmdSave_Click  ' ** Procedure: Above.
26260       End If
26270     End If
26280   End With

EXITP:
26290   Exit Sub

ERRH:
26300   Select Case ERR.Number
        Case Else
26310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26320   End Select
26330   Resume EXITP

End Sub

Private Sub Contact_Fax_Format_Exit(Cancel As Integer)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Fax_Format_Exit"

26410   varOldValue = Null

EXITP:
26420   Exit Sub

ERRH:
26430   Select Case ERR.Number
        Case Else
26440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26450   End Select
26460   Resume EXITP

End Sub

Private Sub Contact_Email_lbl_DblClick(Cancel As Integer)

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_lbl_DblClick"

26510   SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
26520   Exit Sub

ERRH:
26530   Select Case ERR.Number
        Case Else
26540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26550   End Select
26560   Resume EXITP

End Sub

Private Sub Contact_Email_cmd_Click()

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_cmd_Click"

        Dim strTmp01 As String

26610   With Me
26620     If IsNull(.Contact_Email) = False Then
26630       If Trim(.Contact_Email) <> vbNullString Then
26640         DoCmd.Hourglass True
26650         DoEvents
26660         strTmp01 = "MailTo:" & .Contact_Email
26670         If IsNull(.Contact_MailTo) = True Then
26680           .Contact_MailTo = strTmp01
26690           .Contact_Changed = True
26700           cmdSave_Click  ' ** Procedure: Above.
26710         Else
26720           If .Contact_MailTo <> strTmp01 Then
26730             .Contact_MailTo = strTmp01
26740             .Contact_Changed = True
26750             cmdSave_Click  ' ** Procedure: Above.
26760           End If
26770         End If
26780         .Contact_Email_cmd.HyperlinkAddress = .Contact_MailTo
26790         .Contact_Email_cmd.Hyperlink.Follow
26800         .Contact_Email_cmd.HyperlinkAddress = vbNullString
26810       Else
26820         Beep
26830       End If
26840     Else
26850       Beep
26860     End If
26870     DoCmd.Hourglass False
26880   End With

EXITP:
26890   Exit Sub

ERRH:
26900   Select Case ERR.Number
        Case Else
26910     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26920   End Select
26930   Resume EXITP

End Sub

Private Sub Contact_Email_Enter()

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_Enter"

27010   With Me
27020     varOldValue = .Contact_Email
27030   End With

EXITP:
27040   Exit Sub

ERRH:
27050   Select Case ERR.Number
        Case Else
27060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27070   End Select
27080   Resume EXITP

End Sub

Private Sub Contact_Email_GotFocus()

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_GotFocus"

27110   With Me
27120     .Contact_Email.SelLength = 0
27130     .Contact_Email.SelStart = 99
27140   End With

EXITP:
27150   Exit Sub

ERRH:
27160   Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No records.
27170   Case Else
27180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27190   End Select
27200   Resume EXITP

End Sub

Private Sub Contact_Email_BeforeUpdate(Cancel As Integer)

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_BeforeUpdate"

        Dim intPos01 As Integer

27310   With Me
27320     If IsNull(.Contact_Email) = False Then
27330       If Trim(.Contact_Email) <> vbNullString Then
27340         intPos01 = InStr(.Contact_Email, "@")
27350         If intPos01 = 0 Then
27360           Beep
27370           MsgBox "Please enter a valid email address.", vbInformation + vbOKOnly, "Invalid Entry"
27380           Cancel = -1
27390         End If
27400       End If
27410     End If
27420   End With

EXITP:
27430   Exit Sub

ERRH:
27440   Select Case ERR.Number
        Case Else
27450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27460   End Select
27470   Resume EXITP

End Sub

Private Sub Contact_Email_DblClick(Cancel As Integer)

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_DblClick"

27510   With Me
27520     .Contact_Email_cmd.Enabled = True
27530     Contact_Email_cmd_Click  ' ** Procedure: Above.
27540     .Contact_Email_cmd.Enabled = False
          'I CAN'T SEEM TO MAKE IT NOT HIGHLIGHT!
          'Contact_Email_GotFocus  ' ** Procedure: Above.
27550   End With

EXITP:
27560   Exit Sub

ERRH:
27570   Select Case ERR.Number
        Case Else
27580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27590   End Select
27600   Resume EXITP

End Sub

Private Sub Contact_Email_AfterUpdate()
' ** Make sure it's a valid email address.

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_AfterUpdate"

        Dim intPos01 As Integer
        Dim strTmp01 As String

27710   With Me
27720     If IsNull(.Contact_Email) = False Then
27730       If Trim(.Contact_Email) <> vbNullString Then
27740         strTmp01 = Trim(.Contact_Email)
27750         intPos01 = InStr(strTmp01, "@")
27760         If intPos01 > 0 Then
27770           intPos01 = InStr(intPos01, strTmp01, ".")
27780           If intPos01 > 0 Then
27790             .Contact_MailTo = "MailTo:" & strTmp01
27800             .Contact_Changed = True
27810             cmdSave_Click  ' ** Procedure: Above.
27820           Else
27830             Beep
27840             MsgBox "Entry must be a valid email address.", vbInformation + vbOKOnly, "Invalid Entry"
27850             .Contact_Email.SetFocus
27860           End If
27870         Else
27880           Beep
27890           MsgBox "Entry must be a valid email address.", vbInformation + vbOKOnly, "Invalid Entry"
27900           .Contact_Email.SetFocus
27910         End If
27920       Else
27930         If IsNull(varOldValue) = False Then
27940           .Contact_Email = Null
27950           .Contact_MailTo = Null
27960           .Contact_Changed = True
27970           cmdSave_Click  ' ** Procedure: Above.
27980         End If
27990       End If
28000     Else
28010       If IsNull(varOldValue) = False Then
28020         .Contact_Changed = True
28030         cmdSave_Click  ' ** Procedure: Above.
28040       End If
28050     End If
28060   End With

EXITP:
28070   Exit Sub

ERRH:
28080   Select Case ERR.Number
        Case Else
28090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28100   End Select
28110   Resume EXITP

End Sub

Private Sub Contact_Email_KeyDown(KeyCode As Integer, Shift As Integer)

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_KeyDown"

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
28290         lngRecsCur = RecCnt  ' ** Function: Below.
28300         If .CurrentRecord < lngRecsCur Then
28310           MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
28320           If .Contact_Name.Visible = True And .Contact_Name.Enabled = True Then
28330             .Contact_Name.SetFocus
28340           Else
28350             .Contact_Name_LastFirst.SetFocus
28360           End If
28370         Else
28380           DoCmd.SelectObject acForm, .Parent.Name, False
28390           .Parent.opgView.SetFocus
28400         End If
28410       End With
28420     End Select
28430   End If

        ' ** Shift keys.
28440   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28450     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
28460       With Me
28470         intRetVal = 0
28480         If .Contact_Fax.Visible = True And .Contact_Fax.Enabled = True Then
28490           .Contact_Fax.SetFocus
28500         Else
28510           .Contact_Fax_Format.SetFocus
28520         End If
28530       End With
28540     End Select
28550   End If

EXITP:
28560   KeyCode = intRetVal
28570   Exit Sub

ERRH:
28580   intRetVal = 0
28590   Select Case ERR.Number
        Case Else
28600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28610   End Select
28620   Resume EXITP

End Sub

Private Sub Contact_Email_Exit(Cancel As Integer)

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "Contact_Email_Exit"

28710   varOldValue = Null

EXITP:
28720   Exit Sub

ERRH:
28730   Select Case ERR.Number
        Case Else
28740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28750   End Select
28760   Resume EXITP

End Sub

Public Sub EnableCountry(blnEnable As Boolean)

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "EnableCountry"

        Dim blnAcctNo As Boolean, blnShortName As Boolean
        Dim lngNewForm_Width As Long
        Dim blnSortHere As Boolean, blnResort As Boolean
        Dim lngTmp01 As Long, lngTmp02 As Long, intTmp03 As Integer
        Dim intX As Integer

28810   With Me

28820     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
28830       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
28840     End If
28850     lngRecsCur = RecCnt  ' ** Function: Below.

28860     blnAcctNo = .accountno.Visible
28870     blnShortName = .shortname.Visible
28880     blnSortHere = False: blnResort = False

28890     Select Case blnEnable
          Case True

            ' ** lngForm_Width is accountno only, plus country.
28900       If blnAcctNo = True And blnShortName = True Then
              ' ** Original configuration, plus ShortName.
28910         lngNewForm_Width = (lngForm_Width + lngShortName_Width + lngCtlSep)
28920         lngTmp01 = (lngShortName_Width + lngCtlSep)
28930       Else
28940         If blnAcctNo = True Then
                ' ** Original configuration, no adjustment.
28950           lngNewForm_Width = lngForm_Width
28960           lngTmp01 = 0&
28970         ElseIf blnShortName = True Then
                ' ** No accountno, just the difference in widths.
28980           lngNewForm_Width = (lngForm_Width + lngShortName_Diff)
28990           lngTmp01 = lngShortName_Diff
29000         End If
29010       End If
29020       .Width = lngNewForm_Width

29030       If .Sort_line.Left = .Contact_Email_lbl.Left Then blnSortHere = True
29040       .Contact_Email.Left = (lngEmail_Left + lngTmp01)
29050       .Contact_Email_cmd.Left = (.Contact_Email.Left - (3& * lngTpp))
29060       .Contact_Email_lbl.Left = .Contact_Email.Left
29070       .Contact_Email_lbl2.Left = ((.Contact_Email_lbl.Left + .Contact_Email_lbl.Width) - .Contact_Email_lbl2.Width)
29080       .Contact_Email_lbl_dim_hi.Left = (.Contact_Email_lbl.Left + lngTpp)
29090       .Contact_Email_lbl_line.Left = .Contact_Email_lbl.Left
29100       .Contact_Email_lbl_line_dim_hi.Left = (.Contact_Email_lbl_line.Left + lngTpp)
29110       If blnSortHere = True Then
29120         .Sort_lbl.Left = ((.Contact_Email_lbl.Left + .Contact_Email_lbl.Width) - lngSortLbl_Width)
29130         .Sort_line.Left = .Contact_Email_lbl.Left
29140       End If
29150       blnSortHere = False
29160       If .Sort_line.Left = .Contact_Fax_lbl.Left Then blnSortHere = True
29170       .Contact_Fax_Format.Left = (lngFax_Left + lngTmp01)
29180       .Contact_Fax.Left = (lngFax_Left + lngTmp01)
29190       .Contact_Fax_lbl.Left = .Contact_Fax.Left
29200       .Contact_Fax_lbl_dim_hi.Left = (.Contact_Fax_lbl.Left + lngTpp)
29210       .Contact_Fax_lbl_line.Left = .Contact_Fax_lbl.Left
29220       .Contact_Fax_lbl_line_dim_hi.Left = (.Contact_Fax_lbl_line.Left + lngTpp)
29230       If blnSortHere = True Then
29240         .Sort_lbl.Left = ((.Contact_Fax_lbl.Left + .Contact_Fax_lbl.Width) - lngSortLbl_Width)
29250         .Sort_line.Left = .Contact_Fax_lbl.Left
29260       End If
29270       blnSortHere = False
29280       If .Sort_line.Left = .Contact_Phone2_lbl.Left Then blnSortHere = True
29290       .Contact_Phone2_Format.Left = (lngPhone2_Left + lngTmp01)
29300       .Contact_Phone2.Left = (lngPhone2_Left + lngTmp01)
29310       .Contact_Phone2_lbl.Left = .Contact_Phone2.Left
29320       .Contact_Phone2_lbl_dim_hi.Left = (.Contact_Phone2_lbl.Left + lngTpp)
29330       .Contact_Phone2_lbl_line.Left = .Contact_Phone2_lbl.Left
29340       .Contact_Phone2_lbl_line_dim_hi.Left = (.Contact_Phone2_lbl_line.Left + lngTpp)
29350       If blnSortHere = True Then
29360         .Sort_lbl.Left = ((.Contact_Phone2_lbl.Left + .Contact_Phone2_lbl.Width) - lngSortLbl_Width)
29370         .Sort_line.Left = .Contact_Phone2_lbl.Left
29380       End If
29390       blnSortHere = False
29400       If .Sort_line.Left = .Contact_Phone1_lbl.Left Then blnSortHere = True
29410       .Contact_Phone1_Format.Left = (lngPhone1_Left + lngTmp01)
29420       .Contact_Phone1.Left = (lngPhone1_Left + lngTmp01)
29430       .Contact_Phone1_lbl.Left = .Contact_Phone1.Left
29440       .Contact_Phone1_lbl_dim_hi.Left = (.Contact_Phone1_lbl.Left + lngTpp)
29450       .Contact_Phone1_lbl_line.Left = .Contact_Phone1_lbl.Left
29460       .Contact_Phone1_lbl_line_dim_hi.Left = (.Contact_Phone1_lbl_line.Left + lngTpp)
29470       If blnSortHere = True Then
29480         .Sort_lbl.Left = ((.Contact_Phone1_lbl.Left + .Contact_Phone1_lbl.Width) - lngSortLbl_Width)
29490         .Sort_line.Left = .Contact_Phone1_lbl.Left
29500       End If
29510       blnSortHere = False

29520       lngTmp01 = ((.Contact_City.Left + .Contact_City.Width) + lngCtlSep) + lngTpp
29530       If .opgLocContact.Left <> lngTmp01 Then
29540         If lngTmp01 > .opgLocContact.Left Then
29550           lngTmp02 = (lngTmp01 - .opgLocContact.Left)
29560           intTmp03 = (lngTmp02 / lngTpp)
29570           For intX = 1 To intTmp03
29580             .opgLocContact.Left = (.opgLocContact.Left + lngTpp)
29590             .opgLocContact_optOther_lbl.Left = (.opgLocContact_optOther_lbl.Left + lngTpp)
29600             .opgLocContact_optOther.Left = (.opgLocContact_optOther.Left + lngTpp)
29610             .opgLocContact_optUSA_lbl.Left = (.opgLocContact_optUSA_lbl.Left + lngTpp)
29620             .opgLocContact_optUSA.Left = (.opgLocContact_optUSA.Left + lngTpp)
29630           Next  ' ** intX.
29640         ElseIf lngTmp01 < .opgLocContact.Left Then
29650           lngTmp02 = (.opgLocContact.Left - lngTmp01)
29660           intTmp03 = (lngTmp02 / lngTpp)
29670           For intX = 1 To intTmp03
29680             .opgLocContact.Left = (.opgLocContact.Left - lngTpp)
29690             .opgLocContact_optOther_lbl.Left = (.opgLocContact_optOther_lbl.Left - lngTpp)
29700             .opgLocContact_optOther.Left = (.opgLocContact_optOther.Left - lngTpp)
29710             .opgLocContact_optUSA_lbl.Left = (.opgLocContact_optUSA_lbl.Left - lngTpp)
29720             .opgLocContact_optUSA.Left = (.opgLocContact_optUSA.Left - lngTpp)
29730           Next  ' ** intX.
29740         End If
29750         .opgLocContact_box.Left = .opgLocContact.Left
29760         .opgLocContact_optOther_box.Left = (.opgLocContact_optOther.Left - (5& * lngTpp))
29770         .opgLocContact_optOther_lbl_on.Left = .opgLocContact_optOther_lbl.Left
29780         .opgLocContact_optOther_lbl_off.Left = .opgLocContact_optOther_lbl.Left
29790         .opgLocContact_optUSA_box.Left = (.opgLocContact_optUSA.Left - (5& * lngTpp))
29800         .opgLocContact_optUSA_lbl_on.Left = .opgLocContact_optUSA_lbl.Left
29810         .opgLocContact_optUSA_lbl_off.Left = .opgLocContact_optUSA_lbl.Left
29820       Else
              ' ** I'm not sure we need to do anything if it gets here.
29830       End If
29840       .opgLocContact_lbl.Left = (.opgLocContact.Left - lngTpp)
29850       .opgLocContact_lbl_dim_hi.Left = (.opgLocContact_lbl.Left + lngTpp)
29860       .opgLocContact_lbl_line.Left = .opgLocContact_lbl.Left
29870       .opgLocContact_lbl_line_dim_hi.Left = (.opgLocContact_lbl_line.Left + lngTpp)
29880       .opgLocContact.Visible = True
29890       .opgLocContact_box.Visible = True
29900       .opgLocContact_optUSA_box.Visible = True
29910       .opgLocContact_optUSA_lbl_on.Visible = True
29920       .opgLocContact_optUSA_lbl_off.Visible = True
29930       .opgLocContact_optOther_box.Visible = True
29940       .opgLocContact_optOther_lbl_on.Visible = True
29950       .opgLocContact_optOther_lbl_off.Visible = True
29960       .opgLocContact_lbl.Visible = True
29970       .opgLocContact_lbl_line.Visible = True

29980       If .Sort_line.Left = .Contact_State_lbl.Left Then blnSortHere = True
29990       lngTmp01 = ((.opgLocContact.Left + .opgLocContact.Width) + lngCtlSep)
30000       .Contact_State.Left = lngTmp01
30010       .Contact_State_box.Left = (.Contact_State.Left - lngTpp)
30020       .Contact_State_lbl.Left = .Contact_State.Left
30030       .Contact_State_lbl_dim_hi.Left = (.Contact_State_lbl.Left + lngTpp)
30040       .Contact_State_lbl_line.Left = .Contact_State_lbl.Left
30050       .Contact_State_lbl_line_dim_hi.Left = (.Contact_State_lbl_line.Left + lngTpp)
30060       If blnSortHere = True Then
30070         .Sort_line.Left = .Contact_State_lbl.Left
30080         .Sort_lbl.Left = ((.Contact_State_lbl.Left + .Contact_State_lbl.Width) - lngSortLbl_Width)
30090       End If
30100       blnSortHere = False
30110       If .Sort_line.Left = .Contact_Zip_lbl.Left Then blnSortHere = True
30120       lngTmp01 = ((.Contact_State.Left + .Contact_State.Width) + lngCtlSep)
30130       .Contact_Zip.Left = lngTmp01
30140       .Contact_Zip_Format.Left = .Contact_Zip.Left
30150       .Contact_Zip_box.Left = (.Contact_Zip.Left - lngTpp)
30160       .Contact_Zip_lbl.Left = .Contact_Zip.Left
30170       .Contact_Zip_lbl_dim_hi.Left = (.Contact_Zip_lbl.Left + lngTpp)
30180       .Contact_Zip_lbl_line.Left = .Contact_Zip_lbl.Left
30190       .Contact_Zip_lbl_line_dim_hi.Left = (.Contact_Zip_lbl_line.Left + lngTpp)
30200       If blnSortHere = True Then
30210         .Sort_line.Left = .Contact_Zip_lbl.Left
30220         .Sort_lbl.Left = ((.Contact_Zip_lbl.Left + .Contact_Zip_lbl.Width) - lngSortLbl_Width)
30230       End If
30240       blnSortHere = False
            ' ** Make sure these are positioned correctly.
30250       lngTmp01 = ((.Contact_Zip.Left + .Contact_Zip.Width) + lngCtlSep)
30260       .Contact_Country.Left = lngTmp01
30270       .Contact_Country_box.Left = (.Contact_Country.Left - lngTpp)
30280       .Contact_Country_lbl.Left = .Contact_Country.Left
30290       .Contact_Country_lbl_line.Left = .Contact_Country_lbl.Left
30300       .Contact_Country_lbl_dim_hi.Left = (.Contact_Country_lbl.Left + lngTpp)
30310       .Contact_Country_lbl_line_dim_hi.Left = (.Contact_Country_lbl_line.Left + lngTpp)
30320       .Contact_Country.Visible = True
30330       .Contact_Country_lbl.Visible = True
30340       .Contact_Country_lbl_line.Visible = True
30350       lngTmp01 = ((.Contact_Country.Left + .Contact_Country.Width) + lngCtlSep)
30360       .Contact_PostalCode.Left = lngTmp01
30370       .Contact_PostalCode_box.Left = (.Contact_PostalCode.Left - lngTpp)
30380       .Contact_PostalCode_lbl.Left = .Contact_PostalCode.Left
30390       .Contact_PostalCode_lbl_line.Left = .Contact_PostalCode_lbl.Left
30400       .Contact_PostalCode_lbl_dim_hi.Left = (.Contact_PostalCode_lbl.Left + lngTpp)
30410       .Contact_PostalCode_lbl_line_dim_hi.Left = (.Contact_PostalCode_lbl_line.Left + lngTpp)
30420       .Contact_PostalCode.Visible = True
30430       .Contact_PostalCode_lbl.Visible = True
30440       .Contact_PostalCode_lbl_line.Visible = True

30450       If lngRecsCur = 0& Then
30460         .opgLocContact_lbl_dim_hi.Visible = True
30470         .opgLocContact_lbl_line_dim_hi.Visible = True
30480         .Contact_Country_lbl_dim_hi.Visible = True
30490         .Contact_Country_lbl_line_dim_hi.Visible = True
30500         .Contact_PostalCode_lbl_dim_hi.Visible = True
30510         .Contact_PostalCode_lbl_line_dim_hi.Visible = True
30520       End If

30530       .Width = lngNewForm_Width

30540     Case False
            ' ** Wide, without country.

            ' ** lngForm_Width is accountno only, with country.
30550       If blnAcctNo = True And blnShortName = True Then
30560         lngNewForm_Width = ((lngForm_Width + lngShortName_Width + lngCtlSep) - (lngCountry_Diff + lngOpgLoc_Diff))
30570         lngTmp01 = (lngCountry_Diff + lngOpgLoc_Diff) - (lngShortName_Width + lngCtlSep)
30580       Else
30590         If blnAcctNo = True Then
30600           lngNewForm_Width = (lngForm_Width - (lngCountry_Diff + lngOpgLoc_Diff))
30610           lngTmp01 = (lngCountry_Diff + lngOpgLoc_Diff)
30620         ElseIf blnShortName = True Then
30630           lngNewForm_Width = ((lngForm_Width + lngShortName_Diff) - (lngCountry_Diff + lngOpgLoc_Diff))
30640           lngTmp01 = ((lngCountry_Diff + lngOpgLoc_Diff) - lngShortName_Diff)
30650         End If
30660       End If
30670       .Width = lngNewForm_Width
            ' ** The numbers above are just overall changes, not individual control changes.

30680       If .Sort_line.Left = .opgLocContact_lbl.Left Then blnSortHere = True
30690       .opgLocContact.Visible = False
30700       .opgLocContact_box.Visible = False
30710       .opgLocContact_optUSA_box.Visible = False
30720       .opgLocContact_optUSA_lbl_on.Visible = False
30730       .opgLocContact_optUSA_lbl_off.Visible = False
30740       .opgLocContact_optOther_box.Visible = False
30750       .opgLocContact_optOther_lbl_on.Visible = False
30760       .opgLocContact_optOther_lbl_off.Visible = False
30770       .opgLocContact_lbl.Visible = False
30780       .opgLocContact_lbl_dim_hi.Visible = False
30790       .opgLocContact_lbl_line.Visible = False
30800       .opgLocContact_lbl_line_dim_hi.Visible = False
30810       If blnSortHere = True Then
30820         blnResort = True
30830         .Sort_line.Visible = False
30840         .Sort_lbl.Visible = False
30850       End If
30860       blnSortHere = False
30870       If .Sort_line.Left = .Contact_State_lbl.Left Then blnSortHere = True
30880       lngTmp02 = ((.Contact_City.Left + .Contact_City.Width) + lngCtlSep)
30890       .Contact_State.Left = lngTmp02
30900       .Contact_State_box.Left = (.Contact_State.Left - lngTpp)
30910       .Contact_State_lbl.Left = .Contact_State.Left
30920       .Contact_State_lbl_dim_hi.Left = (.Contact_State_lbl.Left + lngTpp)
30930       .Contact_State_lbl_line.Left = .Contact_State_lbl.Left
30940       .Contact_State_lbl_line_dim_hi.Left = (.Contact_State_lbl_line.Left + lngTpp)
30950       If blnSortHere = True Then
30960         .Sort_line.Left = .Contact_State_lbl.Left
30970         .Sort_lbl.Left = ((.Contact_State_lbl.Left + .Contact_State_lbl.Width) - lngSortLbl_Width)
30980       End If
30990       blnSortHere = False
31000       If .Sort_line.Left = .Contact_Zip_lbl.Left Then blnSortHere = True
31010       lngTmp02 = ((.Contact_State.Left + .Contact_State.Width) + lngCtlSep)
31020       .Contact_Zip.Left = lngTmp02
31030       .Contact_Zip_Format.Left = .Contact_Zip.Left
31040       .Contact_Zip_box.Left = (.Contact_Zip.Left - lngTpp)
31050       .Contact_Zip_lbl.Left = .Contact_Zip.Left
31060       .Contact_Zip_lbl_dim_hi.Left = (.Contact_Zip_lbl.Left + lngTpp)
31070       .Contact_Zip_lbl_line.Left = .Contact_Zip_lbl.Left
31080       .Contact_Zip_lbl_line_dim_hi.Left = (.Contact_Zip_lbl_line.Left + lngTpp)
31090       If blnSortHere = True Then
31100         .Sort_line.Left = .Contact_Zip_lbl.Left
31110         .Sort_lbl.Left = ((.Contact_Zip_lbl.Left + .Contact_Zip_lbl.Width) - lngSortLbl_Width)
31120       End If
31130       blnSortHere = False
31140       If .Sort_line.Left = .Contact_Country_lbl.Left Then blnSortHere = True
31150       .Contact_Country.Visible = False
31160       .Contact_Country_box.Visible = False
31170       .Contact_Country_lbl.Visible = False
31180       .Contact_Country_lbl_line.Visible = False
31190       If lngRecsCur = 0& Then
31200         .Contact_Country_lbl_dim_hi.Visible = False
31210         .Contact_Country_lbl_line_dim_hi.Visible = False
31220       End If
31230       If blnSortHere = True Then
31240         blnResort = True
31250         .Sort_lbl.Visible = False
31260         .Sort_lbl.Left = 0&
31270         .Sort_line.Visible = False
31280         .Sort_line.Left = 0&
31290       End If
31300       blnSortHere = False
31310       If .Sort_line.Left = .Contact_PostalCode_lbl.Left Then blnSortHere = True
31320       .Contact_PostalCode.Visible = False
31330       .Contact_PostalCode_box.Visible = False
31340       .Contact_PostalCode_lbl.Visible = False
31350       .Contact_PostalCode_lbl_line.Visible = False
31360       If lngRecsCur = 0& Then
31370         .Contact_PostalCode_lbl_dim_hi.Visible = False
31380         .Contact_PostalCode_lbl_line_dim_hi.Visible = False
31390       End If
31400       If blnSortHere = True Then
31410         blnResort = True
31420         .Sort_lbl.Visible = False
31430         .Sort_lbl.Left = 0&
31440         .Sort_line.Visible = False
31450         .Sort_line.Left = 0&
31460       End If
31470       blnSortHere = False
31480       If .Sort_line.Left = .Contact_Phone1_lbl.Left Then blnSortHere = True
31490       lngTmp02 = ((.Contact_Zip.Left + .Contact_Zip.Width) + lngCtlSep)
31500       .Contact_Phone1.Left = lngTmp02 '(lngPhone1_Left - lngTmp01)
31510       .Contact_Phone1_Format.Left = lngTmp02 '(lngPhone1_Left - lngTmp01)
31520       .Contact_Phone1_lbl.Left = .Contact_Phone1.Left
31530       .Contact_Phone1_lbl_dim_hi.Left = (.Contact_Phone1_lbl.Left + lngTpp)
31540       .Contact_Phone1_lbl_line.Left = .Contact_Phone1_lbl.Left
31550       .Contact_Phone1_lbl_line_dim_hi.Left = (.Contact_Phone1_lbl_line.Left + lngTpp)
31560       If blnSortHere = True Then
31570         .Sort_lbl.Left = ((.Contact_Phone1_lbl.Left + .Contact_Phone1_lbl.Width) - lngSortLbl_Width)
31580         .Sort_line.Left = .Contact_Phone1_lbl.Left
31590       End If
31600       blnSortHere = False
31610       If .Sort_line.Left = .Contact_Phone2_lbl.Left Then blnSortHere = True
31620       lngTmp02 = ((.Contact_Phone1.Left + .Contact_Phone1.Width) + lngCtlSep)
31630       .Contact_Phone2.Left = lngTmp02 '(lngPhone2_Left - lngTmp01)
31640       .Contact_Phone2_Format.Left = lngTmp02 '(lngPhone2_Left - lngTmp01)
31650       .Contact_Phone2_lbl.Left = .Contact_Phone2.Left
31660       .Contact_Phone2_lbl_dim_hi.Left = (.Contact_Phone2_lbl.Left + lngTpp)
31670       .Contact_Phone2_lbl_line.Left = .Contact_Phone2_lbl.Left
31680       .Contact_Phone2_lbl_line_dim_hi.Left = (.Contact_Phone2_lbl_line.Left + lngTpp)
31690       If blnSortHere = True Then
31700         .Sort_lbl.Left = ((.Contact_Phone2_lbl.Left + .Contact_Phone2_lbl.Width) - lngSortLbl_Width)
31710         .Sort_line.Left = .Contact_Phone2_lbl.Left
31720       End If
31730       blnSortHere = False
31740       If .Sort_line.Left = .Contact_Fax_lbl.Left Then blnSortHere = True
31750       lngTmp02 = ((.Contact_Phone2.Left + .Contact_Phone2.Width) + lngCtlSep)
31760       .Contact_Fax.Left = lngTmp02 '(lngFax_Left - lngTmp01)
31770       .Contact_Fax_Format.Left = lngTmp02 '(lngFax_Left - lngTmp01)
31780       .Contact_Fax_lbl.Left = .Contact_Fax.Left
31790       .Contact_Fax_lbl_dim_hi.Left = (.Contact_Fax_lbl.Left + lngTpp)
31800       .Contact_Fax_lbl_line.Left = .Contact_Fax_lbl.Left
31810       .Contact_Fax_lbl_line_dim_hi.Left = (.Contact_Fax_lbl_line.Left + lngTpp)
31820       If blnSortHere = True Then
31830         .Sort_lbl.Left = ((.Contact_Fax_lbl.Left + .Contact_Fax_lbl.Width) - lngSortLbl_Width)
31840         .Sort_line.Left = .Contact_Fax_lbl.Left
31850       End If
31860       blnSortHere = False
31870       If .Sort_line.Left = .Contact_Email_lbl.Left Then blnSortHere = True
31880       lngTmp02 = ((.Contact_Fax.Left + .Contact_Fax.Width) + lngCtlSep)
31890       .Contact_Email.Left = lngTmp02 '(lngEmail_Left - lngTmp01)
31900       .Contact_Email_cmd.Left = (.Contact_Email.Left - (3& * lngTpp))
31910       .Contact_Email_lbl.Left = .Contact_Email.Left
31920       .Contact_Email_lbl2.Left = ((.Contact_Email_lbl.Left + .Contact_Email_lbl.Width) - .Contact_Email_lbl2.Width)
31930       .Contact_Email_lbl_dim_hi.Left = (.Contact_Email_lbl.Left + lngTpp)
31940       .Contact_Email_lbl_line.Left = .Contact_Email_lbl.Left
31950       .Contact_Email_lbl_line_dim_hi.Left = (.Contact_Email_lbl_line.Left + lngTpp)
31960       If blnSortHere = True Then
31970         .Sort_lbl.Left = ((.Contact_Email_lbl.Left + .Contact_Email_lbl.Width) - lngSortLbl_Width)
31980         .Sort_line.Left = .Contact_Email_lbl.Left
31990       End If
32000       blnSortHere = False
32010       DoEvents

32020       If blnResort = True Then
32030         If blnAcctNo = True And blnShortName = True Then
32040           SortNow "shortname_lbl_DblClick"  ' ** Procedure: Below
32050         Else
32060           If blnAcctNo = True Then
32070             SortNow "accountno_lbl_DblClick"  ' ** Procedure: Below
32080           ElseIf blnShortName = True Then
32090             SortNow "shortname_lbl_DblClick"  ' ** Procedure: Below
32100           End If
32110         End If
32120         DoEvents
32130       End If

32140       .Width = (lngNewForm_Width - lngTpp)

32150     End Select
32160     DoEvents

32170   End With

EXITP:
32180   Exit Sub

ERRH:
32190   Select Case ERR.Number
        Case Else
32200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32210   End Select
32220   Resume EXITP

End Sub

Public Sub ResetChanged()

32300 On Error GoTo ERRH

        Const THIS_PROC As String = "ResetChanged"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long

32310   cmdSave_Click  ' ** Procedure: Above.
32320   DoEvents

32330   Set dbs = CurrentDb
32340   With dbs
32350     Set rst = .OpenRecordset("tblAccount_Contact_Staging", dbOpenDynaset, dbConsistent)
32360     With rst
32370       .MoveLast
32380       lngRecs = .RecordCount
32390       .MoveFirst
32400       For lngX = 1& To lngRecs
32410         If ![Contact_Changed] = True Then
32420           .Edit
32430           ![Contact_Changed] = False
32440           ![Contact_DateModified] = Now()
32450           .Update
32460         End If
32470         If lngX < lngRecs Then .MoveNext
32480       Next
32490       .Close
32500     End With
32510     Set rst = Nothing
32520   End With  ' ** dbs.
32530   Set dbs = Nothing

32540   With Me
32550     .Refresh
32560     .Requery
32570   End With

EXITP:
32580   Set rst = Nothing
32590   Set dbs = Nothing
32600   Exit Sub

ERRH:
32610   Select Case ERR.Number
        Case Else
32620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32630   End Select
32640   Resume EXITP

End Sub

Public Function DelRec_AC(lngContactID As Long) As Boolean

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "DelRec_AC"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnRetVal As Boolean

32710   blnRetVal = True

32720   With Me
32730     Set dbs = CurrentDb
32740     With dbs
            ' ** Delete tblAccount_Contact_Staging, by specified [conid].
32750       Set qdf = .QueryDefs("qryAccount_Contacts_19")
32760       With qdf.Parameters
32770         ![conid] = lngContactID
32780       End With
32790       qdf.Execute
32800       .Close
32810     End With
32820     .Requery
32830   End With

EXITP:
32840   Set qdf = Nothing
32850   Set dbs = Nothing
32860   DelRec_AC = blnRetVal
32870   Exit Function

ERRH:
32880   blnRetVal = False
32890   Select Case ERR.Number
        Case Else
32900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
32910   End Select
32920   Resume EXITP

End Function

Public Function RecCnt() As Long

33000 On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

33010   lngRetVal = 0&

33020   With Me
33030     Set rst = .RecordsetClone
33040     With rst
33050       If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
33060       Else
33070         .MoveLast
33080         lngRetVal = .RecordCount
33090       End If
33100       .Close
33110     End With
33120   End With

EXITP:
33130   Set rst = Nothing
33140   RecCnt = lngRetVal
33150   Exit Function

ERRH:
33160   lngRetVal = 0&
33170   Select Case ERR.Number
        Case Else
33180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33190   End Select
33200   Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

33300 On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

33310   Select Case intWhere
        Case acCmdRecordsGoToFirst
33320     DoCmd.RunCommand acCmdRecordsGoToFirst
33330   Case acCmdRecordsGoToLast
33340     DoCmd.RunCommand acCmdRecordsGoToLast
33350   Case acCmdRecordsGoToPrevious
33360     DoCmd.RunCommand acCmdRecordsGoToPrevious
33370   Case acCmdRecordsGoToNext
33380     DoCmd.RunCommand acCmdRecordsGoToNext
33390   Case acCmdRecordsGoToNew
33400     DoCmd.RunCommand acCmdRecordsGoToNew
33410   Case Else
33420     If IsMissing(varID) = False Then
33430       With Me
33440         Set rst = .RecordsetClone
33450         With rst
33460           .FindFirst "[Contact_ID] = " & CStr(varID)
33470           If .NoMatch = False Then
33480             Me.Bookmark = .Bookmark
33490           End If
33500           .Close
33510         End With
33520       End With
33530     End If
33540   End Select

EXITP:
33550   Set rst = Nothing
33560   Exit Sub

ERRH:
33570   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
33580   Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
33590   Case Else
33600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
33610   End Select
33620   Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

33700 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim ctl As Access.Control
        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [alphasort], [Contact_Number]"

33710   With Me
33720     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
33730       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
33740     End If
33750     .Controls(strSortLbl).Visible = False
33760     .Controls(strSortLine).Visible = False
33770     If strProc = "Form_Load" Then
33780       strCalled = "accountno"
33790       Select Case .Controls(strCalled).Visible
            Case True
33800         strSortNow = strSortOrig
33810       Case False
33820         strCalled = "shortname"
33830         strSortNow = "[shortname], [Contact_Number]"
33840       End Select
33850       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
33860       .Controls(strSortLbl).Left = lngSortLbl_Left
33870       .Controls(strSortLbl).Caption = strArwUp
33880       .Controls(strSortLbl).ForeColor = CLR_DKBLU
33890     Else
33900       strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
33910       lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
33920       .Controls(strSortLbl).Left = lngSortLbl_Left
33930       If strCalled = "accountno" Or strCalled = "shortname" Then
33940         Select Case strCalled
              Case "accountno"
33950           strSortAsc = strSortOrig
33960         Case "shortname"
33970           strSortAsc = "[shortname], [Contact_Number]"
33980         End Select
33990         .Controls(strSortLbl).Left = lngSortLbl_Left
34000         If strSortNow = strSortAsc Then
34010           Select Case strCalled
                Case "accountno"
34020             strSortNow = "[alphasort] DESC, [Contact_Number] DESC"
34030           Case "shortname"
34040             strSortNow = "[shortname] DESC, [Contact_Number] DESC"
34050           End Select
34060           .Controls(strSortLbl).Caption = strArwDn
34070           .Controls(strSortLbl).ForeColor = CLR_DKRED
34080         Else
34090           strSortNow = strSortAsc
34100           .Controls(strSortLbl).Caption = strArwUp
34110           .Controls(strSortLbl).ForeColor = CLR_DKBLU
34120         End If
34130       Else
34140         .Controls(strSortLbl).Caption = strArwUp
34150         .Controls(strSortLbl).ForeColor = CLR_DKBLU
34160         Select Case strCalled
              Case "Contact_Name"
34170           Select Case .Controls(strCalled).Visible
                Case True
34180             strSortAsc = "[Contact_Name]" & strStdAsc
34190           Case False
34200             strSortAsc = "[Contact_Name_LastFirst]" & strStdAsc
34210           End Select
34220         Case "Contact_Address1"
34230           strSortAsc = "[Contact_Address1]" & strStdAsc
34240         Case "Contact_Address2"
34250           strSortAsc = "[Contact_Address2]" & strStdAsc
34260         Case "Contact_City"
34270           strSortAsc = "[Contact_City]" & strStdAsc
34280         Case "opgLocContact"
34290           strSortAsc = "[opgLocContact]" & strStdAsc
34300         Case "Contact_State"
34310           strSortAsc = "[Contact_State]" & strStdAsc
34320         Case "Contact_Zip"
34330           Select Case .Controls(strCalled).Visible
                Case True
34340             strSortAsc = "[Contact_Zip]" & strStdAsc
34350           Case False
34360             strSortAsc = "[Contact_Zip_Format]" & strStdAsc
34370           End Select
34380         Case "Contact_Country"
34390           strSortAsc = "[Contact_Country]" & strStdAsc
34400         Case "Contact_PostalCode"
34410           strSortAsc = "[Contact_PostalCode]" & strStdAsc
34420           lngSortLbl_Left = (lngSortLbl_Left + (6& * lngTpp))
34430           .Controls(strSortLbl).Left = lngSortLbl_Left
34440         Case "Contact_Phone1"
34450           Select Case .Controls(strCalled).Visible
                Case True
34460             strSortAsc = "[Contact_Phone1]" & strStdAsc
34470           Case False
34480             strSortAsc = "[Contact_Phone1_Format]" & strStdAsc
34490           End Select
34500         Case "Contact_Phone1"
34510           Select Case .Controls(strCalled).Visible
                Case True
34520             strSortAsc = "[Contact_Phone1]" & strStdAsc
34530           Case False
34540             strSortAsc = "[Contact_Phone1_Format]" & strStdAsc
34550           End Select
34560         Case "Contact_Fax"
34570           Select Case .Controls(strCalled).Visible
                Case True
34580             strSortAsc = "[Contact_Fax]" & strStdAsc
34590           Case False
34600             strSortAsc = "[Contact_Fax_Format]" & strStdAsc
34610           End Select
34620         Case "Contact_Email"
34630           strSortAsc = "[Contact_Email]" & strStdAsc
34640         End Select
34650         If strSortNow = strSortAsc Then
34660           intCnt = (CharCnt(strSortAsc, ",") + 1)  ' ** Module Function: modStringFuncs.
34670           Select Case intCnt
                Case 1
                  ' ** [accounttype_description] DESC.
34680             strSortNow = strSortNow & " DESC"
34690           Case 2
                  ' ** [accounttypegroup_name] DESC, [accounttype] DESC.
34700             intPos01 = InStr(strSortAsc, ",")
34710             strSortNow = Left(strSortAsc, (intPos01 - 1)) & " DESC" & Mid(strSortAsc, intPos01) & " DESC"
34720           Case 3
34730             intPos01 = InStr(strSortAsc, ",")
34740             strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
34750             strTmp02 = Mid(strSortAsc, intPos01)
34760             intPos01 = InStr(2, strSortAsc, ",")
34770             strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC" & Mid(strTmp02, intPos01) & " DESC"
34780             strSortNow = strTmp01
34790           End Select
34800           .Controls(strSortLbl).Caption = strArwDn
34810           .Controls(strSortLbl).ForeColor = CLR_DKRED
34820         Else
34830           strSortNow = strSortAsc
34840         End If
34850       End If
34860     End If
34870     .Controls(strSortLbl).Visible = True
34880     .Controls(strSortLine).Width = lngTpp  ' ** So it doesn't push of the right side of the form.
34890     .Controls(strSortLine).Left = .Controls(strCalled & "_lbl").Left
34900     .Controls(strSortLine).Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
34910     .Controls(strSortLine).Visible = True
34920     .OrderBy = strSortNow
34930     .OrderByOn = True
34940   End With

EXITP:
34950   Set ctl = Nothing
34960   Exit Sub

ERRH:
34970   Select Case ERR.Number
        Case Else
34980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
34990   End Select
35000   Resume EXITP

End Sub

Public Function SortNow_Get() As String

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow_Get"

        Dim strRetVal As String

35110   If strSortNow = vbNullString Then
35120     strRetVal = strSortOrig
35130   Else
35140     strRetVal = strSortNow
35150   End If

EXITP:
35160   SortNow_Get = strRetVal
35170   Exit Function

ERRH:
35180   strRetVal = RET_ERR
35190   Select Case ERR.Number
        Case Else
35200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
35210   End Select
35220   Resume EXITP

End Function
