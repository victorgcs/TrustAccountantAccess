VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLicense_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLicense_Edit"

'VGC 11/21/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Calc Reg Code:    A {cmdCalculateReg}
' **   Cancel:           C {cmdCancel}
' **   Calc Price Code:  D {cmdCalculatePricing}
' **   Reset:            E {cmdReset}
' **   Save To File:     F {cmdSaveToFile}
' **   Price - Reg Code: G {PricingCode}
' **   Demo Mode:        H {chkDemo}
' **   Expires:          I {RegistrationExpires}
' **   Calc Lic Code:    L {cmdCalculateLisc}
' **   Account Limit:    M {AccountLimit}
' **   License Code:     N {LicenseCode}
' **   OK:               O {cmdOK}
' **   Price - Expires:  P {PricingExpires}
' **   Reg Code:         R {RegistrationCode}
' **   On-Screen Notice: S {WarningScreen}
' **   Licensed To:      T {LicensedTo}
' **   Pop-Up Notice:    U {WarningMessage}
' **   Validate Codes:   V {cmdValidate}
' **   Password Cycle:   W {PasswordCycle}
' **   Exit:             X {cmdCancel}
' **   Copy Codes:       Y {cmdCopyCodes}

'Available Keys:
'B J K Q Z

' ** Shortcut Ctrl keys responsive from this form:
' **   Save To File:     S {cmdSaveToFile}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnReset_Focus As Boolean, blnReset_MouseDown As Boolean
Private blnClearPricing_Focus As Boolean, blnClearPricing_MouseDown As Boolean

Private lngAccts As Long, lngPW_Cycle As Long, lngPW_Msg As Long, lngPW_Screen As Long
Private blnPasswordOnly As Boolean, blnCodesSaved As Boolean
Private strCallingForm As String, strCallersCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = Trim(.OpenArgs)
150           intPos01 = InStr(strTmp01, "~")
160           If intPos01 > 0 Then

170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strCallersCallingForm = Mid(strTmp01, (intPos01 + 1))

190             blnPasswordOnly = True: blnCodesSaved = False
200             blnReset_Focus = False: blnReset_MouseDown = False
210             blnClearPricing_Focus = False: blnClearPricing_MouseDown = False

220             CLR_DISABLED_FG = CLR_DKGRY
230             CLR_DISABLED_BG = CLR_LTTEAL

240             Set dbs = CurrentDb
250             With dbs
                  ' ** Account, grouped, with AccountCount.
260               Set qdf = .QueryDefs("qrySecurity_License_08")
270               Set rst = qdf.OpenRecordset
280               With rst
290                 .MoveFirst
300                 lngAccts = ![AccountCount]
310                 .Close
320               End With
330             End With
340             .AccountCount = lngAccts

350             .cmdOK.Enabled = False

360             .LicensedTo = DecodeString(IniFile_Get("License", "Firm", EncodeString("Call Delta Data for Licensing info."), _
                  gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Function: modCodeUtilities, modStartupFuncs.
370             .RegistrationExpires = DecodeString(IniFile_Get("License", "Expires", "", _
                  gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Function: modCodeUtilities, modStartupFuncs.
380             .AccountLimit = DecodeString(IniFile_Get("License", "Limit", "", _
                  gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Function: modCodeUtilities, modStartupFuncs.
390             .PricingExpires = DecodeString(IniFile_Get("License", "Pricing", vbNullString, _
                  gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Function: modCodeUtilities, modStartupFuncs.

400             With dbs
410               Set rst = .OpenRecordset("tblSecurity_License", dbOpenDynaset, dbConsistent)
420               With rst
430                 If .BOF = True And .EOF = True Then
                      ' ** Catastrophe! Devastation!
440                   Cancel = -1
450                 Else
460                   .MoveFirst
470                   If IsNull(![seclic_cycle]) = False Then
480                     lngPW_Cycle = CLng(DecodeString(![seclic_cycle]))  ' ** Module Function: modCodeUtilities.
490                   Else
500                     .Edit
510                     ![seclic_cycle] = EncodeString("0")  ' ** Module Function: modCodeUtilities. DEF_CYCLE
520                     ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
530                     ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
540                     ![seclic_datemodified] = Now()
550                     .Update
560                     lngPW_Cycle = 0& 'DEF_CYCLE
570                   End If
580                   If IsNull(![seclic_cycle_screen]) = False Then
590                     lngPW_Screen = CLng(DecodeString(![seclic_cycle_screen]))  ' ** Module Function: modCodeUtilities.
600                   Else
610                     .Edit
620                     ![seclic_cycle_screen] = EncodeString("0")  ' ** Module Function: modCodeUtilities. DEF_SCREEN
630                     ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
640                     ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
650                     ![seclic_datemodified] = Now()
660                     .Update
670                     lngPW_Screen = 0& 'DEF_SCREEN
680                   End If
690                   If IsNull(![seclic_cycle_message]) = False Then
700                     lngPW_Msg = CLng(DecodeString(![seclic_cycle_message]))  ' ** Module Function: modCodeUtilities.
710                   Else
720                     .Edit
730                     ![seclic_cycle_message] = EncodeString("0")  ' ** Module Function: modCodeUtilities. DEF_MSGBOX
740                     ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
750                     ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
760                     ![seclic_datemodified] = Now()
770                     .Update
780                     lngPW_Msg = 0& 'DEF_MSGBOX
790                   End If
800                 End If
810                 .Close
820               End With
830             End With
840             dbs.Close
850             .PasswordCycle = lngPW_Cycle
860             .WarningMessage = lngPW_Msg
870             .WarningScreen = lngPW_Screen

880             .cmdReset_raised_img.Visible = True
890             .cmdReset_raised_semifocus_dots_img.Visible = False
900             .cmdReset_raised_focus_img.Visible = False
910             .cmdReset_raised_focus_dots_img.Visible = False
920             .cmdReset_sunken_focus_dots_img.Visible = False
930             .cmdReset_raised_img_dis.Visible = False

940           Else
950             Cancel = -1
960           End If
970         Else
980           Cancel = -1
990         End If
1000      Else
1010        Cancel = -1
1020      End If
1030    End With

1040    If Cancel = -1 Then
1050      Beep
1060      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1070    End If

EXITP:
1080    Set rst = Nothing
1090    Set qdf = Nothing
1100    Set dbs = Nothing
1110    Exit Sub

ERRH:
1120    Select Case ERR.Number
        Case Else
1130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1140    End Select
1150    Resume EXITP

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

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyEscape
1270        intRetVal = 0
1280        cmdCancel_Click  ' ** Procedure: Below.
1290      End Select
1300    End If

        ' ** Alt keys.
1310    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1320      Select Case intRetVal
          Case vbKeyX
1330        intRetVal = 0
1340        cmdCancel_Click  ' ** Procedure: Below.
1350      End Select
1360    End If

        ' ** Ctrl keys.
1370    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1380      Select Case intRetVal
          Case vbKeyS
1390        intRetVal = 0
1400        cmdSaveToFile_Click  ' ** Procedure: Below.
1410      End Select
1420    End If

EXITP:
1430    KeyCode = intRetVal
1440    Exit Sub

ERRH:
1450    intRetVal = 0
1460    Select Case ERR.Number
        Case Else
1470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1480    End Select
1490    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1510    If DataErr <> 0 Then
1520      Select Case DataErr
          Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
1530        Select Case Me.ActiveControl.Name
            Case "RegistrationExpires"
1540          MsgBox "Please enter a valid Registration Expires date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
1550        Case "PricingExpires"
1560          MsgBox "Please enter a valid Pricing Registration Expires date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
1570        End Select
1580        Response = acDataErrContinue
1590      Case Else
1600        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1610        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1620      End Select
1630    End If

EXITP:
1640    Exit Sub

ERRH:
1650    Select Case ERR.Number
        Case Else
1660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1670    End Select
1680    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** Will always be frmLicense; it's strCallersCallingForm that changes.
1710    DoCmd.SelectObject acForm, strCallingForm, False

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1810    With Me
1820      If .cmdReset_raised_focus_dots_img.Visible = True Or .cmdReset_raised_focus_img.Visible = True Then
1830        Select Case blnReset_Focus
            Case True
1840          .cmdReset_raised_semifocus_dots_img.Visible = True
1850          .cmdReset_raised_img.Visible = False
1860        Case False
1870          .cmdReset_raised_img.Visible = True
1880          .cmdReset_raised_semifocus_dots_img.Visible = False
1890        End Select
1900        .cmdReset_raised_focus_img.Visible = False
1910        .cmdReset_raised_focus_dots_img.Visible = False
1920        .cmdReset_sunken_focus_dots_img.Visible = False
1930        .cmdReset_raised_img_dis.Visible = False
1940      End If
1950      If .cmdClearPricing_raised_focus_dots_img.Visible = True Or .cmdClearPricing_raised_focus_img.Visible = True Then
1960        Select Case blnClearPricing_Focus
            Case True
1970          .cmdClearPricing_raised_semifocus_dots_img.Visible = True
1980          .cmdClearPricing_raised_img.Visible = False
1990        Case False
2000          .cmdClearPricing_raised_img.Visible = True
2010          .cmdClearPricing_raised_semifocus_dots_img.Visible = False
2020        End Select
2030        .cmdClearPricing_raised_focus_img.Visible = False
2040        .cmdClearPricing_raised_focus_dots_img.Visible = False
2050        .cmdClearPricing_sunken_focus_dots_img.Visible = False
2060        .cmdClearPricing_raised_img_dis.Visible = False
2070      End If
2080    End With

EXITP:
2090    Exit Sub

ERRH:
2100    Select Case ERR.Number
        Case Else
2110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2120    End Select
2130    Resume EXITP

End Sub

Private Sub cmdOK_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

2210    If blnCodesSaved = False Then
          ' ** The fact that cmdOK is accessible infers that changes have been made.
2220      If MsgBox("You have made changes, but have not saved them." & vbCrLf & vbCrLf & _
              "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved") = vbYes Then
2230        DoCmd.Close acForm, THIS_NAME
2240      End If
2250    Else
2260      DoCmd.Close acForm, THIS_NAME
2270    End If

EXITP:
2280    Exit Sub

ERRH:
2290    Select Case ERR.Number
        Case Else
2300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2310    End Select
2320    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyTab
2470        With Me
2480          intRetVal = 0
2490          .cmdCancel.SetFocus
2500        End With
2510      End Select
2520    End If

        ' ** Shift keys.
2530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2540      Select Case intRetVal
          Case vbKeyTab
2550        With Me
2560          intRetVal = 0
2570          .cmdSaveToFile.SetFocus
2580        End With
2590      End Select
2600    End If

EXITP:
2610    KeyCode = intRetVal
2620    Exit Sub

ERRH:
2630    intRetVal = 0
2640    Select Case ERR.Number
        Case Else
2650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2660    End Select
2670    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2710    gblnClosing = False
2720    DoCmd.Close acForm, THIS_NAME

EXITP:
2730    Exit Sub

ERRH:
2740    Select Case ERR.Number
        Case Else
2750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2760    End Select
2770    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab
2870        With Me
2880          intRetVal = 0
2890          .LicensedTo.SetFocus
2900        End With
2910      End Select
2920    End If

        ' ** Shift keys.
2930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2940      Select Case intRetVal
          Case vbKeyTab
2950        With Me
2960          intRetVal = 0
2970          .cmdOK.SetFocus
2980        End With
2990      End Select
3000    End If

EXITP:
3010    KeyCode = intRetVal
3020    Exit Sub

ERRH:
3030    intRetVal = 0
3040    Select Case ERR.Number
        Case Else
3050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3060    End Select
3070    Resume EXITP

End Sub

Private Sub chkDemo_AfterUpdate()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkDemo_AfterUpdate"

3110    With Me
3120      blnPasswordOnly = False
3130      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
3140      Select Case .chkDemo
          Case True
3150        .LicensedTo = "Call Delta Data for Licensing info."
3160        .chkDemo_lbl.FontBold = True
3170        .chkDemo_lbl2.FontBold = True
3180        .chkDemo_lbl3.FontBold = True
3190      Case False
3200        .LicensedTo = vbNullString
3210        .chkDemo_lbl.FontBold = False
3220        .chkDemo_lbl2.FontBold = False
3230        .chkDemo_lbl3.FontBold = False
3240      End Select
3250    End With

EXITP:
3260    Exit Sub

ERRH:
3270    Select Case ERR.Number
        Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

End Sub

Private Sub LicensedTo_AfterUpdate()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "LicensedTo_AfterUpdate"

3410    With Me
3420      blnPasswordOnly = False
3430      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
3440    End With

EXITP:
3450    Exit Sub

ERRH:
3460    Select Case ERR.Number
        Case Else
3470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3480    End Select
3490    Resume EXITP

End Sub

Private Sub LicensedTo_GotFocus()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "LicensedTo_GotFocus"

3510    With Me
3520      .LicensedTo.SelLength = 0
3530      .LicensedTo.SelStart = 99
3540    End With

EXITP:
3550    Exit Sub

ERRH:
3560    Select Case ERR.Number
        Case Else
3570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3580    End Select
3590    Resume EXITP

End Sub

Private Sub LicensedTo_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "LicensedTo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3620    intShiftDown = (Shift And acShiftMask) > 0
3630    intAltDown = (Shift And acAltMask) > 0
3640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3670        With Me
3680          intRetVal = 0
3690          .RegistrationExpires.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3750        With Me
3760          intRetVal = 0
3770          .cmdCancel.SetFocus
3780        End With
3790      End Select
3800    End If

EXITP:
3810    KeyCode = intRetVal
3820    Exit Sub

ERRH:
3830    intRetVal = 0
3840    Select Case ERR.Number
        Case Else
3850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3860    End Select
3870    Resume EXITP

End Sub

Private Sub RegistrationExpires_Change()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "RegistrationExpires_Change"

3910    With Me
3920      blnPasswordOnly = False
3930      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
3940      .RegistrationCode = vbNullString
3950    End With

EXITP:
3960    Exit Sub

ERRH:
3970    Select Case ERR.Number
        Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Sub

Private Sub RegistrationExpires_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "RegistrationExpires_MouseUp"

4110    With Me
4120      .RegistrationExpires.SelStart = 0
4130      .RegistrationExpires.SelLength = 0
4140    End With

EXITP:
4150    Exit Sub

ERRH:
4160    Select Case ERR.Number
        Case Else
4170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4180    End Select
4190    Resume EXITP

End Sub

Private Sub RegistrationCode_AfterUpdate()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "RegistrationCode_AfterUpdate"

4210    With Me
4220      blnPasswordOnly = False
4230      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
4240    End With

EXITP:
4250    Exit Sub

ERRH:
4260    Select Case ERR.Number
        Case Else
4270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4280    End Select
4290    Resume EXITP

End Sub

Private Sub cmdCalculateReg_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalculateReg_Click"

4310    With Me
4320      blnPasswordOnly = False
4330      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
4340      If IsNull(.RegistrationExpires) = True Then
4350        MsgBox "Must specify a Registration Expires date.", vbExclamation + vbOKOnly, "Entry Required"
4360        .RegistrationExpires.SetFocus
4370      Else
4380        If IsDate(.RegistrationExpires) = False Then
4390          MsgBox "Must specify a valid Registration Expires date.", vbExclamation + vbOKOnly, "Invalid Date"
4400          .RegistrationExpires.SetFocus
4410        Else
4420          If CLng(CDate(.RegistrationExpires)) <= CLng(Date) Then
4430            MsgBox "Must specify a future date for Registration Expires.", vbExclamation + vbOKOnly, "Invalid Date"
4440            .RegistrationExpires.SetFocus
4450          Else
4460            .RegistrationExpires = Format(.RegistrationExpires, "mm/dd/yyyy")
4470            .RegistrationCode = CalcRegCode(.RegistrationExpires)  ' ** Module Function: modCodeUtilities.
4480          End If
4490        End If
4500      End If
4510    End With

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub AccountLimit_Change()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "AccountLimit_Change"

4610    With Me
4620      blnPasswordOnly = False
4630      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
4640      .LicenseCode = vbNullString
4650    End With

EXITP:
4660    Exit Sub

ERRH:
4670    Select Case ERR.Number
        Case Else
4680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4690    End Select
4700    Resume EXITP

End Sub

Private Sub LicenseCode_AfterUpdate()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "LicenseCode_AfterUpdate"

4810    With Me
4820      blnPasswordOnly = False
4830      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
4840    End With

EXITP:
4850    Exit Sub

ERRH:
4860    Select Case ERR.Number
        Case Else
4870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4880    End Select
4890    Resume EXITP

End Sub

Private Sub cmdCalculateLisc_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalculateLisc_Click"

4910    With Me
4920      blnPasswordOnly = False
4930      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
4940      If IsNull(.AccountLimit) = True Then
4950        MsgBox "An account Limit must be specified.", vbExclamation + vbOKOnly, "Entry Required"
4960        .AccountLimit.SetFocus
4970      Else
4980        If Val(.AccountLimit) = 0 Then
4990          MsgBox "Invalid account Limit.", vbExclamation + vbOKOnly, "Entry Required"
5000          .AccountLimit.SetFocus
5010          .AccountLimit.SelStart = 0
5020          .AccountLimit.SelLength = Len(.AccountLimit)
5030        Else
5040          If Val(.AccountLimit) < 0 Then .AccountLimit = Abs(Val(.AccountLimit))
5050          If Val(.AccountLimit) > 2000000 Then
5060            MsgBox "The account Limit too high!", vbExclamation + vbOKOnly, "Invalid Entry"
5070            .AccountLimit.SetFocus
5080            .AccountLimit.SelStart = 0
5090            .AccountLimit.SelLength = Len(.AccountLimit)
5100          Else
5110            .LicenseCode = CalcLimitCode(.AccountLimit)  ' ** Module Function: modCodeUtilities.
5120          End If
5130        End If
5140      End If
5150    End With

EXITP:
5160    Exit Sub

ERRH:
5170    Select Case ERR.Number
        Case Else
5180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5190    End Select
5200    Resume EXITP

End Sub

Private Sub Password_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "Password_box_MouseMove"

5310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub PasswordCycle_AfterUpdate()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "PasswordCycle_AfterUpdate"

5410    With Me
5420      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
5430      If IsNull(.PasswordCycle) = False Then
5440        If CLng(.PasswordCycle) < 0 Then
5450          MsgBox "The Password Cycle cannot be a negative number.", vbInformation + vbOKOnly, "Invalid Entry"
5460          .PasswordCycle = 0& 'DEF_CYCLE
5470        ElseIf CLng(.PasswordCycle) = 0& Then
5480          .WarningScreen = 0&
5490          .WarningMessage = 0&
5500        ElseIf CLng(.PasswordCycle) < 7& Then
5510          If MsgBox("Are you sure you want passwords to expire every " & _
                  CStr(CLng(.PasswordCycle)) & " days?", vbQuestion + vbYesNo + vbDefaultButton2, _
                  "Short Password Cycle") <> vbYes Then
5520            .PasswordCycle = 0& 'DEF_CYCLE
5530          End If
5540        ElseIf CLng(.PasswordCycle) > 1825& Then  ' ** 5 years.
5550          MsgBox "Your entry will be changed to 0, for no password expiration.", _
                vbInformation + vbOKOnly, "Passwords Will Not Expire"
5560          .PasswordCycle = 0&
5570          .WarningScreen = 0&
5580          .WarningMessage = 0&
5590        End If
5600      Else
5610        .PasswordCycle = 0&
5620        .WarningScreen = 0&
5630        .WarningMessage = 0&
5640      End If
5650    End With

EXITP:
5660    Exit Sub

ERRH:
5670    THAT_PROC = THIS_PROC
5680    That_Erl = Erl
5690    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5700    Resume EXITP

End Sub

Private Sub WarningScreen_AfterUpdate()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "WarningScreen_AfterUpdate"

5810    With Me
5820      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
5830      If CLng(.PasswordCycle) > 0& Then
5840        If IsNull(.WarningScreen) = False Then
5850          If CLng(.WarningScreen) < 0 Then
5860            MsgBox "The On-Screen Notice cannot be a negative number.", vbInformation + vbOKOnly, "Invalid Entry"
5870            If CLng(.PasswordCycle) > DEF_SCREEN Then
5880              .WarningScreen = DEF_SCREEN
5890            ElseIf CLng(.PasswordCycle) > 3& Then
5900              .WarningScreen = Int(CLng(.PasswordCycle) * 0.25)
5910            Else
5920              .WarningScreen = 1&
5930            End If
5940          ElseIf CLng(.WarningScreen) > CLng(.PasswordCycle) Then
5950            MsgBox "The On-Screen Notice cannot be greater than the Password Cycle.", _
                  vbInformation + vbOKOnly, "Invalid Entry"
5960            If CLng(.PasswordCycle) > DEF_SCREEN Then
5970              .WarningScreen = DEF_SCREEN
5980            ElseIf CLng(.PasswordCycle) > 3& Then
5990              .WarningScreen = Int(CLng(.PasswordCycle) * 0.25)
6000            Else
6010              .WarningScreen = 1&
6020            End If
6030          End If
6040        Else
6050          .WarningScreen = 0&
6060        End If
6070      Else
6080        .WarningScreen = 0&
6090      End If
6100    End With

EXITP:
6110    Exit Sub

ERRH:
6120    THAT_PROC = THIS_PROC
6130    That_Erl = Erl
6140    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6150    Resume EXITP

End Sub

Private Sub WarningMessage_AfterUpdate()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "WarningMessage_AfterUpdate"

6210    With Me
6220      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
6230      If CLng(.PasswordCycle) > 0& Then
6240        If IsNull(.WarningMessage) = False Then
6250          If CLng(.WarningMessage) < 0 Then
6260            MsgBox "The Pop-Up Notice cannot be a negative number.", vbInformation + vbOKOnly, "Invalid Entry"
6270            If CLng(.PasswordCycle) > DEF_MSGBOX Then
6280              .WarningMessage = DEF_MSGBOX
6290            ElseIf CLng(.PasswordCycle) > 3& Then
6300              .WarningMessage = Int(CLng(.PasswordCycle) * 0.25)
6310            Else
6320              .WarningMessage = 1&
6330            End If
6340          ElseIf CLng(.WarningMessage) > CLng(.PasswordCycle) Then
6350            MsgBox "The Pop-Up Notice cannot be greater than the Password Cycle.", _
                  vbInformation + vbOKOnly, "Invalid Entry"
6360            If CLng(.PasswordCycle) > DEF_MSGBOX Then
6370              .WarningMessage = DEF_MSGBOX
6380            ElseIf CLng(.PasswordCycle) > 3& Then
6390              .WarningMessage = Int(CLng(.PasswordCycle) * 0.25)
6400            Else
6410              .WarningMessage = 1&
6420            End If
6430          End If
6440        Else
6450          .WarningMessage = 0&
6460        End If
6470      Else
6480        .WarningMessage = 0&
6490      End If
6500    End With

EXITP:
6510    Exit Sub

ERRH:
6520    THAT_PROC = THIS_PROC
6530    That_Erl = Erl
6540    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6550    Resume EXITP

End Sub

Private Sub WarningMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "WarningMessage_MouseMove"

6610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6620    Exit Sub

ERRH:
6630    Select Case ERR.Number
        Case Else
6640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6650    End Select
6660    Resume EXITP

End Sub

Private Sub cmdReset_GotFocus()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_GotFocus"

6710    With Me
6720      blnReset_Focus = True
6730      .cmdReset_raised_semifocus_dots_img.Visible = True
6740      .cmdReset_raised_img.Visible = False
6750      .cmdReset_raised_focus_img.Visible = False
6760      .cmdReset_raised_focus_dots_img.Visible = False
6770      .cmdReset_sunken_focus_dots_img.Visible = False
6780      .cmdReset_raised_img_dis.Visible = False
6790    End With

EXITP:
6800    Exit Sub

ERRH:
6810    Select Case ERR.Number
        Case Else
6820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6830    End Select
6840    Resume EXITP

End Sub

Private Sub cmdReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseDown"

6910    With Me
6920      blnReset_MouseDown = True
6930      .cmdReset_sunken_focus_dots_img.Visible = True
6940      .cmdReset_raised_img.Visible = False
6950      .cmdReset_raised_semifocus_dots_img.Visible = False
6960      .cmdReset_raised_focus_img.Visible = False
6970      .cmdReset_raised_focus_dots_img.Visible = False
6980      .cmdReset_raised_img_dis.Visible = False
6990    End With

EXITP:
7000    Exit Sub

ERRH:
7010    Select Case ERR.Number
        Case Else
7020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7030    End Select
7040    Resume EXITP

End Sub

Private Sub cmdReset_Click()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_Click"

7110    With Me
7120      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
7130      .PasswordCycle = 0& 'DEF_CYCLE
7140      .WarningScreen = 0& 'DEF_SCREEN
7150      .WarningMessage = 0& 'DEF_MSGBOX
7160    End With

EXITP:
7170    Exit Sub

ERRH:
7180    Select Case ERR.Number
        Case Else
7190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7200    End Select
7210    Resume EXITP

End Sub

Private Sub cmdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseMove"

7310    With Me
7320      If blnReset_MouseDown = False Then
7330        Select Case blnReset_Focus
            Case True
7340          .cmdReset_raised_focus_dots_img.Visible = True
7350          .cmdReset_raised_focus_img.Visible = False
7360        Case False
7370          .cmdReset_raised_focus_img.Visible = True
7380          .cmdReset_raised_focus_dots_img.Visible = False
7390        End Select
7400        .cmdReset_raised_img.Visible = False
7410        .cmdReset_raised_semifocus_dots_img.Visible = False
7420        .cmdReset_sunken_focus_dots_img.Visible = False
7430        .cmdReset_raised_img_dis.Visible = False
7440      End If

7450    End With

EXITP:
7460    Exit Sub

ERRH:
7470    Select Case ERR.Number
        Case Else
7480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7490    End Select
7500    Resume EXITP

End Sub

Private Sub cmdReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseUp"

7610    With Me
7620      .cmdReset_raised_focus_dots_img.Visible = True
7630      .cmdReset_raised_img.Visible = False
7640      .cmdReset_raised_semifocus_dots_img.Visible = False
7650      .cmdReset_raised_focus_img.Visible = False
7660      .cmdReset_sunken_focus_dots_img.Visible = False
7670      .cmdReset_raised_img_dis.Visible = False
7680      blnReset_MouseDown = False
7690    End With

EXITP:
7700    Exit Sub

ERRH:
7710    Select Case ERR.Number
        Case Else
7720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7730    End Select
7740    Resume EXITP

End Sub

Private Sub cmdReset_LostFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_LostFocus"

7810    With Me
7820      .cmdReset_raised_img.Visible = True
7830      .cmdReset_raised_semifocus_dots_img.Visible = False
7840      .cmdReset_raised_focus_img.Visible = False
7850      .cmdReset_raised_focus_dots_img.Visible = False
7860      .cmdReset_sunken_focus_dots_img.Visible = False
7870      .cmdReset_raised_img_dis.Visible = False
7880      blnReset_Focus = False
7890    End With

EXITP:
7900    Exit Sub

ERRH:
7910    Select Case ERR.Number
        Case Else
7920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7930    End Select
7940    Resume EXITP

End Sub

Private Sub Pricing_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "Pricing_box_MouseMove"

8010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8020    Exit Sub

ERRH:
8030    Select Case ERR.Number
        Case Else
8040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8050    End Select
8060    Resume EXITP

End Sub

Private Sub PricingExpires_Change()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "PricingExpires_Change"

8110    With Me
8120      blnPasswordOnly = False
8130      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
8140      .PricingCode = vbNullString
8150    End With

EXITP:
8160    Exit Sub

ERRH:
8170    Select Case ERR.Number
        Case Else
8180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8190    End Select
8200    Resume EXITP

End Sub

Private Sub PricingExpires_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "PricingExpires_MouseUp"

8310    With Me
8320      .PricingExpires.SelStart = 0
8330      .PricingExpires.SelLength = 0
8340    End With

EXITP:
8350    Exit Sub

ERRH:
8360    Select Case ERR.Number
        Case Else
8370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8380    End Select
8390    Resume EXITP

End Sub

Private Sub PricingCode_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "PricingCode_AfterUpdate"

8410    With Me
8420      blnPasswordOnly = False
8430      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
8440    End With

EXITP:
8450    Exit Sub

ERRH:
8460    Select Case ERR.Number
        Case Else
8470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8480    End Select
8490    Resume EXITP

End Sub

Private Sub PricingCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "PricingCode_MouseMove"

8510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8520    Exit Sub

ERRH:
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Sub

Private Sub cmdCalculatePricing_Click()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalculatePricing_Click"

8610    With Me
8620      blnPasswordOnly = False
8630      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
8640      If IsNull(.PricingExpires) = True Then
8650        MsgBox "Must specify a Pricing Registration Expires date.", vbExclamation + vbOKOnly, "Date Required"
8660        .PricingExpires.SetFocus
8670      Else
8680        If IsDate(.PricingExpires) = False Then
8690          MsgBox "Must specify a valid Pricing Registration Expires date.", vbExclamation + vbOKOnly, "Invalid Date"
8700          .PricingExpires.SetFocus
8710        Else
8720          If CLng(CDate(.PricingExpires)) <= CLng(Date) Then
8730            MsgBox "Must specify a future date for Pricing Registration Expires.", vbExclamation + vbOKOnly, "Invalid Date"
8740            .PricingExpires.SetFocus
8750          Else
8760            .PricingExpires = Format(.PricingExpires, "mm/dd/yyyy")
8770            .PricingCode = CalcPriceCode(.PricingExpires)  ' ** Module Function: modCodeUtilities.
8780          End If
8790        End If
8800      End If
8810    End With

EXITP:
8820    Exit Sub

ERRH:
8830    Select Case ERR.Number
        Case Else
8840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8850    End Select
8860    Resume EXITP

End Sub

Private Sub cmdCalculatePricing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalculatePricing_MouseMove"

8910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8920    Exit Sub

ERRH:
8930    Select Case ERR.Number
        Case Else
8940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8950    End Select
8960    Resume EXITP

End Sub

Private Sub cmdClearPricing_GotFocus()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearPricing_GotFocus"

9010    With Me
9020      blnClearPricing_Focus = True
9030      .cmdClearPricing_raised_semifocus_dots_img.Visible = True
9040      .cmdClearPricing_raised_img.Visible = False
9050      .cmdClearPricing_raised_focus_img.Visible = False
9060      .cmdClearPricing_raised_focus_dots_img.Visible = False
9070      .cmdClearPricing_sunken_focus_dots_img.Visible = False
9080      .cmdClearPricing_raised_img_dis.Visible = False
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

Private Sub cmdClearPricing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearPricing_MouseDown"

9210    With Me
9220      blnClearPricing_MouseDown = True
9230      .cmdClearPricing_sunken_focus_dots_img.Visible = True
9240      .cmdClearPricing_raised_img.Visible = False
9250      .cmdClearPricing_raised_semifocus_dots_img.Visible = False
9260      .cmdClearPricing_raised_focus_img.Visible = False
9270      .cmdClearPricing_raised_focus_dots_img.Visible = False
9280      .cmdClearPricing_raised_img_dis.Visible = False
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

Private Sub cmdClearPricing_Click()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearPricing_Click"

9410    With Me
9420      .PricingExpires = Null
9430      .PricingCode = Null
9440      If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
9450    End With

EXITP:
9460    Exit Sub

ERRH:
9470    Select Case ERR.Number
        Case Else
9480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9490    End Select
9500    Resume EXITP

End Sub

Private Sub cmdClearPricing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearPricing_MouseMove"

9610    With Me
9620      If blnClearPricing_MouseDown = False Then
9630        Select Case blnClearPricing_Focus
            Case True
9640          .cmdClearPricing_raised_focus_dots_img.Visible = True
9650          .cmdClearPricing_raised_focus_img.Visible = False
9660        Case False
9670          .cmdClearPricing_raised_focus_img.Visible = True
9680          .cmdClearPricing_raised_focus_dots_img.Visible = False
9690        End Select
9700        .cmdClearPricing_raised_img.Visible = False
9710        .cmdClearPricing_raised_semifocus_dots_img.Visible = False
9720        .cmdClearPricing_sunken_focus_dots_img.Visible = False
9730        .cmdClearPricing_raised_img_dis.Visible = False
9740      End If
9750    End With

EXITP:
9760    Exit Sub

ERRH:
9770    Select Case ERR.Number
        Case Else
9780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9790    End Select
9800    Resume EXITP

End Sub

Private Sub cmdClearPricing_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearPricing_MouseUp"

9910    With Me
9920      .cmdClearPricing_raised_focus_dots_img.Visible = True
9930      .cmdClearPricing_raised_img.Visible = False
9940      .cmdClearPricing_raised_semifocus_dots_img.Visible = False
9950      .cmdClearPricing_raised_focus_img.Visible = False
9960      .cmdClearPricing_sunken_focus_dots_img.Visible = False
9970      .cmdClearPricing_raised_img_dis.Visible = False
9980      blnClearPricing_MouseDown = False
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

Private Sub cmdClearPricing_LostFocus()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClearPricing_LostFocus"

10110   With Me
10120     .cmdClearPricing_raised_img.Visible = True
10130     .cmdClearPricing_raised_semifocus_dots_img.Visible = False
10140     .cmdClearPricing_raised_focus_img.Visible = False
10150     .cmdClearPricing_raised_focus_dots_img.Visible = False
10160     .cmdClearPricing_sunken_focus_dots_img.Visible = False
10170     .cmdClearPricing_raised_img_dis.Visible = False
10180     blnClearPricing_Focus = False
10190   End With

EXITP:
10200   Exit Sub

ERRH:
10210   Select Case ERR.Number
        Case Else
10220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10230   End Select
10240   Resume EXITP

End Sub

Private Sub cmdCopyCodes_Click()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCopyCodes_Click"

        Dim strTmp01 As String

10310   With Me
10320     .LicensedTo = DecodeString(IniFile_Get("License", "Firm", EncodeString("Call Delta Data for Licensing info."), _
            gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
10330     .RegistrationExpires = DecodeString(IniFile_Get("License", "Expires", "", _
            gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
10340     .RegistrationCode = vbNullString
10350     cmdCalculateReg_Click  ' ** Procedure: Above.
10360     .AccountLimit = DecodeString(IniFile_Get("License", "Limit", "", _
            gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
10370     .LicenseCode = vbNullString
10380     cmdCalculateLisc_Click  ' ** Procedure: Above.
10390     strTmp01 = Trim(DecodeString(IniFile_Get("License", "Pricing", "", _
            gstrTrustDataLocation & gstrFile_LIC)))    ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
10400     If strTmp01 = vbNullString Then
10410       .PricingExpires = Null
10420     Else
10430       .PricingExpires = CDate(strTmp01)
10440     End If
10450     .PricingCode = vbNullString
10460     If Security_PricingChk = True Then  ' ** Module Function: modSecurityFunctions.
10470       cmdCalculatePricing_Click  ' ** Procedure: Above.
10480     End If
10490   End With

EXITP:
10500   Exit Sub

ERRH:
10510   Select Case ERR.Number
        Case Else
10520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10530   End Select
10540   Resume EXITP

End Sub

Private Sub cmdValidate_Click()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdValidate_Click"

10610   If CalcCodes = True Then  ' ** Function: Below.
10620     MsgBox "License Information Validated.", vbInformation + vbOKOnly, "License Codes Validated"
10630   End If

EXITP:
10640   Exit Sub

ERRH:
10650   Select Case ERR.Number
        Case Else
10660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10670   End Select
10680   Resume EXITP

End Sub

Private Sub cmdValidate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdValidate_MouseMove"

10710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10720   Exit Sub

ERRH:
10730   Select Case ERR.Number
        Case Else
10740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10750   End Select
10760   Resume EXITP

End Sub

Private Sub cmdSaveToFile_Click()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaveToFile_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strLicense As String, strVersion As String, dblVersion As Double
        Dim blnRetVal As Boolean

10810   With Me
10820     If blnPasswordOnly = True Then
10830       Set dbs = CurrentDb
10840       With dbs
10850         Set rst = .OpenRecordset("tblSecurity_License", dbOpenDynaset, dbConsistent)
10860         With rst
10870           If .BOF = True And .EOF = True Then
                  ' ** Unthinkable! Impossible!
10880             .AddNew
10890             ![seclic_licensedto] = Me.LicensedTo
10900             ![seclic_clientpath_ta] = CurrentAppPath  ' ** Module Function: modFileUtilities.
10910             ![seclic_datapath_ta] = CurrentBackendPath  ' ** Module Function: modFileUtilities.
10920             ![seclic_auxiliarypath] = CurrentBackendPath  ' ** Module Function: modFileUtilities.
10930             ![seclic_cycle] = EncodeString(CStr(Me.PasswordCycle))  ' ** Module Function: modCodeUtilities.
10940             ![seclic_cycle_screen] = EncodeString(CStr(Me.WarningScreen))  ' ** Module Function: modCodeUtilities.
10950             ![seclic_cycle_message] = EncodeString(CStr(Me.WarningMessage))  ' ** Module Function: modCodeUtilities.
10960             ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
10970             ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
10980             ![seclic_datemodified] = Now()
10990             .Update
11000           Else
11010             .MoveFirst
11020             .Edit
11030             ![seclic_cycle] = EncodeString(CStr(Me.PasswordCycle))  ' ** Module Function: modCodeUtilities.
11040             ![seclic_cycle_screen] = EncodeString(CStr(Me.WarningScreen))  ' ** Module Function: modCodeUtilities.
11050             ![seclic_cycle_message] = EncodeString(CStr(Me.WarningMessage))  ' ** Module Function: modCodeUtilities.
11060             ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
11070             ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
11080             ![seclic_datemodified] = Now()
11090             .Update
11100           End If
11110           .Close
11120         End With
11130         .Close
11140       End With
11150       MsgBox "Password Cycle Information Saved", vbInformation + vbOKOnly, "Password Cycle Entries Saved"
11160       blnCodesSaved = True
11170     Else
11180       If CalcCodes = True Then  ' ** Function: Below.
11190         DoCmd.Hourglass True
11200         DoEvents
              ' ** Licensed To.
11210         blnRetVal = IniFile_Set("License", "Firm", EncodeString(.LicensedTo), _
                gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
11220         Set dbs = CurrentDb
11230         With dbs
                ' ** Update License Name, by specified [licnam].
11240           Set qdf = .QueryDefs("qrySecurity_License_09")
11250           With qdf.Parameters
11260             ![licnam] = Me.LicensedTo
11270           End With
11280           qdf.Execute
                ' ** Update tblSecurity_License, by specified [licnam].
11290           Set qdf = .QueryDefs("qrySecurity_License_10")
11300           With qdf.Parameters
11310             ![licnam] = Me.LicensedTo
11320           End With
11330           qdf.Execute
11340           .Close
11350         End With
              ' ** Limit code.
11360         blnRetVal = IniFile_Set("License", "Limit", EncodeString(.AccountLimit), _
                gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
              ' ** Expires.
11370         blnRetVal = IniFile_Set("License", "Expires", EncodeString(.RegistrationExpires), _
                gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
              ' ** Password Cycle.
11380         Set dbs = CurrentDb
11390         With dbs
                ' ** Update tblSecurity_License, by specified [liccyc], [licscr], [licmsg].
11400           Set qdf = .QueryDefs("qrySecurity_License_11")
11410           With qdf.Parameters
11420             ![liccyc] = EncodeString(CStr(Me.PasswordCycle))  ' ** Module Function: modCodeUtilities.
11430             ![licscr] = EncodeString(CStr(Me.WarningScreen))  ' ** Module Function: modCodeUtilities.
11440             ![licmsg] = EncodeString(CStr(Me.WarningMessage))  ' ** Module Function: modCodeUtilities.
11450           End With
11460           qdf.Execute
11470           .Close
11480         End With
              ' ** Pricing.
11490         If IsNull(.PricingCode) = False Then
11500           If Len(.PricingCode) <> 0 Then
11510             blnRetVal = IniFile_Set("License", "Pricing", EncodeString(.PricingExpires), _
                    gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
11520             gblnPricingAllowed = True
11530           Else
11540             blnRetVal = IniFile_Del("License", "Pricing", _
                    gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
11550             gblnPricingAllowed = False
11560           End If
11570         Else
11580           blnRetVal = IniFile_Del("License", "Pricing", _
                  gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
11590           gblnPricingAllowed = False
11600         End If

11610         dblVersion = Nz(DLookup("Version", "[License Name]"), 0)
11620         strVersion = CStr(dblVersion)
11630         If InStr(strVersion, ".") > 0 Then
11640           If Len(Mid(strVersion, (InStr(strVersion, ".") + 1))) > 1 Then
11650             strVersion = Left(strVersion, InStr(strVersion, ".")) & _
                    Mid(strVersion, (InStr(strVersion, ".") + 1), 1) & "." & Mid(strVersion, (InStr(strVersion, ".") + 2))
11660           Else
11670             strVersion = strVersion & ".00"
11680           End If
11690         Else
11700           strVersion = strVersion & ".0.00"
11710         End If

11720         strLicense = DecodeString(IniFile_Get("License", "Firm", EncodeString("Call Delta Data, Inc., for Licensing info."), _
                gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
11730         strLicense = "Licensed to " & strLicense & "    Version " & _
                IIf(Left(Right(strVersion, 2), 1) = ".", strVersion & "0", strVersion)
11740         If IsLoaded("frmMenu_Title", acForm) = True Then  ' ** Module Function: modFileUtilities.
11750           Forms("frmMenu_Title").LicenseInfo = strLicense
11760         End If

11770         DoCmd.Hourglass False
11780         MsgBox "License Information Saved", vbInformation + vbOKOnly, "License Entries Saved"
11790         blnCodesSaved = True
11800       End If
11810     End If  ' ** blnPasswordOnly.
11820   End With

EXITP:
11830   Set rst = Nothing
11840   Set qdf = Nothing
11850   Set dbs = Nothing
11860   Exit Sub

ERRH:
11870   Select Case ERR.Number
        Case Else
11880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11890   End Select
11900   Resume EXITP

End Sub

Private Sub cmdSaveToFile_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaveToFile_KeyDown"

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
          Case vbKeyTab
12070       With Me
12080         intRetVal = 0
12090         .cmdOK.SetFocus
12100       End With
12110     End Select
12120   End If

        ' ** Shift keys.
12130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12140     Select Case intRetVal
          Case vbKeyTab
12150       With Me
12160         intRetVal = 0
12170         .cmdValidate.SetFocus
12180       End With
12190     End Select
12200   End If

EXITP:
12210   KeyCode = intRetVal
12220   Exit Sub

ERRH:
12230   intRetVal = 0
12240   Select Case ERR.Number
        Case Else
12250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12260   End Select
12270   Resume EXITP

End Sub

Private Sub cmdSaveToFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaveToFile_MouseMove"

12310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12320   Exit Sub

ERRH:
12330   Select Case ERR.Number
        Case Else
12340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12350   End Select
12360   Resume EXITP

End Sub

Private Function CalcCodes() As Boolean

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "CalcCodes"

        Dim blnHasPricingExpires As Boolean, blnHasPricingCode As Boolean
        Dim blnRetVal As Boolean

12410   blnRetVal = True  ' ** Unless proven otherwise.

12420   With Me
12430     If IsNull(.LicensedTo) Then
12440       blnRetVal = False
12450       MsgBox "Must specify a Licensed To name.", vbExclamation + vbOKOnly, "Entry Required"
12460       .LicensedTo.SetFocus
12470     Else
12480       If Len(.LicensedTo) = 0 Then
12490         blnRetVal = False
12500         MsgBox "Must specify a Licensed To name.", vbExclamation + vbOKOnly, "Entry Required"
12510         .LicensedTo.SetFocus
12520       Else
12530         If Len(.LicensedTo) > 200 Then
12540           blnRetVal = False
12550           MsgBox "The Licensed To name is too long!", vbExclamation + vbOKOnly, "Invalid Entry"
12560           .LicensedTo.SetFocus
12570         Else
12580           If IsNull(.RegistrationExpires) = True Then
12590             blnRetVal = False
12600             MsgBox "Must specify a Registration Expires date.", vbExclamation + vbOKOnly, "Date Required"
12610             .RegistrationExpires.SetFocus
12620           Else
12630             If Len(.RegistrationExpires) = 0 Then
12640               blnRetVal = False
12650               MsgBox "Must specify a Registration Expires date.", vbExclamation + vbOKOnly, "Date Required"
12660               .RegistrationExpires.SetFocus
12670             Else
12680               If IsDate(.RegistrationExpires) = False Then
12690                 blnRetVal = False
12700                 MsgBox "Must specify a valid date for Registration Expires.", vbExclamation + vbOKOnly, "Invalid Date"
12710                 .RegistrationExpires.SetFocus
12720               Else
12730                 If CLng(CDate(.RegistrationExpires)) <= CLng(Date) Then
12740                   blnRetVal = False
12750                   MsgBox "Must specify a future date for Registration Expires.", vbExclamation + vbOKOnly, "Invalid Date"
12760                   .RegistrationExpires.SetFocus
12770                 Else
12780                   If IsNull(.RegistrationCode) = True Then
12790                     blnRetVal = False
12800                     MsgBox "Must specify a Registration Code.", vbExclamation + vbOKOnly, "Entry Required"
12810                     .RegistrationCode.SetFocus
12820                   Else
12830                     If Len(.RegistrationCode) = 0 Then
12840                       blnRetVal = False
12850                       MsgBox "Must specify a Registration Code.", vbExclamation + vbOKOnly, "Entry Required"
12860                       .RegistrationCode.SetFocus
12870                     End If
12880                   End If
12890                 End If
12900               End If
12910             End If
12920           End If
12930         End If
12940       End If
12950     End If

12960     If blnRetVal = True Then

12970       .RegistrationExpires = Format(.RegistrationExpires, "mm/dd/yyyy")

            ' ** Ignore the validation if both values are missing.
12980       blnHasPricingExpires = True: blnHasPricingCode = True
12990       If IsNull(.PricingExpires) = True Then
13000         blnHasPricingExpires = False
13010       Else
13020         If Len(.PricingExpires) = 0 Then
13030           blnHasPricingExpires = False
13040         End If
13050       End If
13060       If blnHasPricingExpires = True Then
13070         If IsNull(.PricingCode) = True Then
13080           blnHasPricingCode = False
13090         Else
13100           If Len(.PricingCode) = 0 Then
13110             blnHasPricingCode = False
13120           End If
13130         End If
13140         If blnHasPricingExpires = True And blnHasPricingCode = True Then
13150           If IsDate(.PricingExpires) = False Then
13160             MsgBox "Must specify a valid date for Pricing Registration Expires.", vbExclamation + vbOKOnly, "Invalid Date"
13170             .PricingExpires.SetFocus
13180             blnRetVal = False
13190           Else
13200             If CLng(CDate(.PricingExpires)) <= CLng(Date) Then
13210               MsgBox "Must specify a future date for Pricing Registration Expires.", vbExclamation + vbOKOnly, "Invalid Date"
13220               .PricingExpires.SetFocus
13230               blnRetVal = False
13240             Else
13250               .PricingExpires = Format(.PricingExpires, "mm/dd/yyyy")
13260             End If
13270           End If
13280         Else
13290           .PricingCode = Null
13300         End If
13310       Else
13320         .PricingCode = Null
13330       End If

13340     End If

13350     If blnRetVal = True Then
13360       If blnHasPricingExpires = True And blnHasPricingCode = True Then
13370         If Val(.PricingCode) <> CalcPriceCode(.PricingExpires) Then  ' ** Module Function: modCodeUtilities.
13380           MsgBox "Invalid Pricing Registration Code.", vbExclamation + vbOKOnly, "Invalid Entry"
13390           .PricingCode.SetFocus
13400           .PricingCode.SelStart = 0
13410           .PricingCode.SelLength = Len(.PricingCode)
13420           blnRetVal = False
13430         End If
13440       End If
13450     End If

13460     If blnRetVal = True Then
13470       If IsNull(.AccountLimit) = True Then
13480         blnRetVal = False
13490         MsgBox "Must specify an account Limit quantity.", vbExclamation + vbOKOnly, "Entry Required"
13500         .AccountLimit.SetFocus
13510       Else
13520         If Val(.AccountLimit) = 0 Then
13530           blnRetVal = False
13540           MsgBox "Invalid account Limit quantity.", vbExclamation + vbOKOnly, "Entry Required"
13550           .AccountLimit.SetFocus
13560           .AccountLimit.SelStart = 0
13570           .AccountLimit.SelLength = Len(.AccountLimit)
13580         Else
13590           If Val(.AccountLimit) < 0 Then .AccountLimit = Abs(Val(.AccountLimit))
13600           If Val(.AccountLimit) > 2000000 Then
13610             blnRetVal = False
13620             MsgBox "The account Limit quantity is too high!", vbExclamation + vbOKOnly, "Invalid Entry"
13630             .AccountLimit.SetFocus
13640             .AccountLimit.SelStart = 0
13650             .AccountLimit.SelLength = Len(.AccountLimit)
13660           Else
13670             If IsNull(.LicenseCode) = True Then
13680               blnRetVal = False
13690               MsgBox "Must specify a License Code.", vbExclamation + vbOKOnly, "Entry Required"
13700               .LicenseCode.SetFocus
13710             Else
13720               If Len(.LicenseCode) = 0 Then
13730                 blnRetVal = False
13740                 MsgBox "Must specify a License Code.", vbExclamation + vbOKOnly, "Entry Required"
13750                 .LicenseCode.SetFocus
13760               Else
13770                 If Val(.RegistrationCode) <> CalcRegCode(.RegistrationExpires) Then
13780                   blnRetVal = False
13790                   MsgBox "Invalid Registration Code.", vbExclamation + vbOKOnly, "Invalid Entry"
13800                   .RegistrationCode.SetFocus
13810                   .RegistrationCode.SelStart = 0
13820                   .RegistrationCode.SelLength = Len(.RegistrationCode)
13830                 Else
13840                   If .AccountLimit <> CalcLimitNumber(.LicenseCode) Then
13850                     blnRetVal = False
13860                     MsgBox "Invalid License Code.", vbExclamation + vbOKOnly, "Invalid Entry"
13870                     .LicenseCode.SetFocus
13880                     .LicenseCode.SelStart = 0
13890                     .LicenseCode.SelLength = Len(.LicenseCode)
13900                   End If
13910                 End If
13920               End If
13930             End If
13940           End If
13950         End If
13960       End If
13970     End If

13980   End With

EXITP:
13990   CalcCodes = blnRetVal
14000   Exit Function

ERRH:
14010   blnRetVal = False
14020   Select Case ERR.Number
        Case Else
14030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14040   End Select
14050   Resume EXITP

End Function
