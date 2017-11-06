VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLicense"

'VGC 09/05/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Pricing Reg:      A {cmdPricing}
' **   Cancel:           C {cmdCancel}
' **   Reset:            E {cmdReset}
' **   Price - Reg Code: G {PricingCode}
' **   Expires:          I {RegistrationExpires}
' **   Account Limit:    M {AccountLimit}
' **   License Code:     N {LicenseCode}
' **   OK:               O {cmdOK}
' **   Price - Expires:  P {PricingExpires}
' **   Quit:             Q {cmdQuit}
' **   Reg Code:         R {RegistrationCode}
' **   On-Screen Notice: S {WarningScreen}
' **   Licensed To:      T {LicensedTo}
' **   Pop-Up Notice:    U {WarningMessage}
' **   Password Cycle:   W {PasswordCycle}
' **   Exit:             X {cmdCancel}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ***********************************************
' ** Easter Egg (Secret Key):
' ** cmdKey_MouseUp
' ***********************************************

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnPricing_Focus As Boolean, blnPricing_MouseDown As Boolean
Private blnKey_Focus As Boolean, blnKey_MouseDown As Boolean

Private blnDone As Boolean, blnCancel As Boolean
Private lngAccts As Long, lngPW_Cycle As Long, lngPW_Msg As Long, lngPW_Screen As Long
Private strCallingForm As String
Private blnFromUtilityMenu As Boolean, blnQuitInvoked As Boolean, blnPasswordOnly As Boolean
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strTmp01 As String

110     DoCmd.Restore
120     blnDone = False: blnCancel = False

130     With Me
140       If IsNull(.OpenArgs) = False Then
150         If Trim(.OpenArgs) <> vbNullString Then

160           strCallingForm = Trim(.OpenArgs)
              ' ** frmMenu_Title   : acDialog       {modStartupFuncs.InitializeTables()}
              ' ** frmMenu_Utility : acWindowNormal {Default}

170           blnFromUtilityMenu = False: blnQuitInvoked = False
180           blnPasswordOnly = True
190           blnPricing_Focus = False: blnPricing_MouseDown = False
200           blnKey_Focus = False: blnKey_MouseDown = False

210           CLR_DISABLED_FG = CLR_DKGRY
220           CLR_DISABLED_BG = CLR_LTTEAL

230           .cmdPricing2_raised_img.Visible = False
240           .cmdPricing2_raised_semifocus_dots_img.Visible = False
250           .cmdPricing2_raised_focus_img.Visible = False
260           .cmdPricing2_raised_focus_dots_img.Visible = False
270           .cmdPricing2_sunken_focus_dots_img.Visible = False
280           .Pricing_Box.Visible = False
290           .Pricing_Box2.Visible = False
300           .PricingExpires.Enabled = False
310           .PricingExpires.Visible = False
320           .PricingCode.Enabled = False
330           .PricingCode.Visible = False
340           .Pricing_hline01.Visible = False
350           .Pricing_hline02.Visible = False
360           .Pricing_hline03.Visible = False
370           .Pricing_vline01.Visible = False
380           .Pricing_vline02.Visible = False
390           .Pricing_vline03.Visible = False
400           .Pricing_vline04.Visible = False
410           .cmdPricing.Visible = True
420           .cmdPricing.Enabled = True
430           .cmdPricing_raised_img.Visible = True
440           .cmdPricing_raised_semifocus_dots_img.Visible = False
450           .cmdPricing_raised_focus_img.Visible = False
460           .cmdPricing_raised_focus_dots_img.Visible = False
470           .cmdPricing_sunken_focus_dots_img.Visible = False
480           .cmdPricing_raised_img_dis.Visible = False
490           DoEvents

500           Set dbs = CurrentDb
510           With dbs
                ' ** Account, grouped, with AccountCount.
520             Set qdf = .QueryDefs("qrySecurity_License_08")
530             Set rst = qdf.OpenRecordset
540             With rst
550               If .BOF = True And .EOF = True Then
                    ' ** Seriously wrong!
560                 lngAccts = 0&
570               Else
580                 .MoveFirst
590                 lngAccts = rst![AccountCount]
600               End If
610               .Close
620             End With
630           End With
640           .AccountCount = lngAccts

650           Select Case strCallingForm
              Case "frmMenu_Title"
660             .cmdOK.Visible = True
670             .cmdOK.Enabled = True
680             .cmdCancel.Visible = False
690             .cmdCancel.Enabled = False
700             .cmdQuit.Visible = True
710             .cmdQuit.Enabled = True
720             .Modal = True
730           Case "frmMenu_Utility"
740             blnFromUtilityMenu = True
750             .cmdOK.Visible = True
760             .cmdOK.Enabled = False  ' ** Unless they change something.
770             .cmdCancel.Visible = True
780             .cmdCancel.Enabled = True
790             .cmdQuit.Visible = False
800             .cmdQuit.Enabled = False
810             If Security_LicenseChk = True Then  ' ** Module Function: modSecurityFunctions.
820               .Registration_lbl2.Visible = True
830             End If
840           End Select

850           .LicensedTo = DecodeString(IniFile_Get("License", "Firm", EncodeString("Call Delta Data for Licensing info."), _
                gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Function: modCodeUtilities, modStartupFuncs.
860           .RegistrationExpires = DecodeString(IniFile_Get("License", "Expires", "", _
                gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Function: modCodeUtilities, modStartupFuncs.
870           strTmp01 = DecodeString(IniFile_Get("License", "Pricing", "", _
                gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Function: modCodeUtilities, modStartupFuncs.
880           .PricingExpires = strTmp01
890           If Security_PricingChk = True Then  ' ** Module Function: modSecurityFunctions.
900             cmdPricing_Click  ' ** Procedure: Below.
910             DoEvents
920             .LicensedTo.SetFocus
930           End If
940           .AccountLimit = DecodeString(IniFile_Get("License", "Limit", "", _
                gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Function: modCodeUtilities, modStartupFuncs.
950           If .LicensedTo = "Call Delta Data for Licensing info." Then
960             .LicensedTo.Enabled = True
970           End If

980           lngPW_Cycle = 0&: lngPW_Msg = 0&: lngPW_Screen = 0&
990           With dbs
1000            Set rst = .OpenRecordset("tblSecurity_License", dbOpenDynaset, dbConsistent)
1010            With rst
1020              If .BOF = True And .EOF = True Then
                    ' ** Catastrophe! Devastation!
1030                Cancel = -1
1040              Else
1050                .MoveFirst
1060                If IsNull(![seclic_cycle]) = False Then
1070                  lngPW_Cycle = CLng(DecodeString(![seclic_cycle]))  ' ** Module Function: modCodeUtilities.
1080                Else
1090                  .Edit
1100                  ![seclic_cycle] = EncodeString("0")  ' ** Module Function: modCodeUtilities. DEF_CYCLE
1110                  ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
1120                  ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1130                  ![seclic_datemodified] = Now()
1140                  .Update
1150                  lngPW_Cycle = 0& 'DEF_CYCLE
1160                End If
1170                If IsNull(![seclic_cycle_screen]) = False Then
1180                  lngPW_Screen = CLng(DecodeString(![seclic_cycle_screen]))  ' ** Module Function: modCodeUtilities.
1190                Else
1200                  .Edit
1210                  ![seclic_cycle_screen] = EncodeString("0")  ' ** Module Function: modCodeUtilities. DEF_SCREEN
1220                  ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
1230                  ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1240                  ![seclic_datemodified] = Now()
1250                  .Update
1260                  lngPW_Screen = 0& 'DEF_SCREEN
1270                End If
1280                If IsNull(![seclic_cycle_message]) = False Then
1290                  lngPW_Msg = CLng(DecodeString(![seclic_cycle_message]))  ' ** Module Function: modCodeUtilities.
1300                Else
1310                  .Edit
1320                  ![seclic_cycle_message] = EncodeString("0")  ' ** Module Function: modCodeUtilities. DEF_MSGBOX
1330                  ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
1340                  ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1350                  ![seclic_datemodified] = Now()
1360                  .Update
1370                  lngPW_Msg = 0& 'DEF_MSGBOX
1380                End If
1390              End If
1400              .Close
1410            End With
1420          End With
1430          dbs.Close
1440          .PasswordCycle = lngPW_Cycle
1450          .WarningMessage = lngPW_Msg
1460          .WarningScreen = lngPW_Screen

1470        Else
1480          Cancel = -1
1490        End If
1500      Else
1510        Cancel = -1
1520      End If
1530    End With

1540    If Cancel = -1 Then
1550      DoCmd.Hourglass False
1560      Beep
1570      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1580    End If

EXITP:
1590    Set rst = Nothing
1600    Set qdf = Nothing
1610    Set dbs = Nothing
1620    Exit Sub

ERRH:
1630    DoCmd.Hourglass False
1640    Select Case ERR.Number
        Case Else
1650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1660    End Select
1670    Resume EXITP

End Sub

Private Sub Form_Load()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1710    With Me
1720      .AccountCount.ForeColor = CLR_DISABLED_FG
1730      .AccountCount.BackColor = CLR_DISABLED_BG
1740      DoCmd.Hourglass False
1750    End With

EXITP:
1760    Exit Sub

ERRH:
1770    DoCmd.Hourglass False
1780    Select Case ERR.Number
        Case Else
1790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1800    End Select
1810    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1920    intShiftDown = (Shift And acShiftMask) > 0
1930    intAltDown = (Shift And acAltMask) > 0
1940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyEscape
1970        intRetVal = 0
1980        cmdCancel_Click  ' ** Procedure: Below.
1990      End Select
2000    End If

        ' ** Alt keys.
2010    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2020      Select Case intRetVal
          Case vbKeyX
2030        intRetVal = 0
2040        cmdCancel_Click  ' ** Procedure: Below.
2050      End Select
2060    End If

        ' ** Ctrl keys.
2070    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2080      Select Case intRetVal
          Case vbKeyS
2090        intRetVal = 0
2100        cmdSave_Click  ' ** Procedure: Below.
2110      End Select
2120    End If

EXITP:
2130    KeyCode = intRetVal
2140    Exit Sub

ERRH:
2150    intRetVal = 0
2160    Select Case ERR.Number
        Case Else
2170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2180    End Select
2190    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2210    If DataErr <> 0 Then
2220      Select Case DataErr
          Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2230        Select Case Me.ActiveControl.Name
            Case "RegistrationExpires"
2240          MsgBox "Please enter a valid Registration Expires date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
2250        Case "PricingExpires"
2260          MsgBox "Please enter a valid Pricing Registration Expires date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
2270        End Select
2280        Response = acDataErrContinue
2290      Case 7753  ' ** The value you entered doesn't meet the validation rule defined for the field or control.
            ' ** Ignore, let messages inform.
2300        Response = acDataErrContinue
2310      Case Else
2320        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2330        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2340      End Select
2350    End If

EXITP:
2360    Exit Sub

ERRH:
2370    Select Case ERR.Number
        Case Else
2380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2390    End Select
2400    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** From Cancel_Click():
        ' **   blnCancel = True
        ' **   gblnClosing = False
        ' ** From Quick_Click():
        ' **   blnQuitInvoked = True
        ' **   gblnClosing = False

        ' ** MAKE SURE A FAILED LICENSE CHECK EXITS!
2510    Select Case blnQuitInvoked
        Case True
          ' ** If coming from startup, make sure this is false.
2520      gblnClosing = False
2530      gblnMessage = True
2540    Case False

2550      If blnCancel = True Then
            ' ** If coming from startup, make sure this is false.
2560        gblnClosing = False
2570      End If
2580      If Security_LicenseChk = False Then  ' ** Module Function: modSecurityFunctions
            ' ** Make sure someone knows about this.
2590        gblnClosing = False
2600      End If

2610      gstrReturningForm = THIS_NAME
2620      gblnSetFocus = True
2630      Select Case strCallingForm
          Case "frmMenu_Utility"
            ' ** Gets here, irrespective of blnDone or blnCancel.
2640  On Error Resume Next
2650        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2660  On Error GoTo ERRH
2670      Case "frmMenu_Title"
            ' ** This should mean it came from (Beyond The Grave!):
            ' **   modStartupFuncs
            ' **     InitializeTables()
2680        If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Functions: modFileUtilities.
              ' ** Just let the startup sequence continue.
              ' ** If canceled (so gblnClosing will be False), InitializeTables() will do QuitNow().
2690          DoCmd.SelectObject acForm, strCallingForm, False
2700        Else
              ' ** It's called as acDialog, so this shouldn't ever happen.
2710  On Error Resume Next
              ' ** Error: 2486  You can't carry out this action at the present time.
2720          DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
2730  On Error GoTo ERRH
2740        End If
2750      Case Else
            ' ** Nothing at the moment.
2760      End Select

2770    End Select  ' ** blnQuitInvoked.

EXITP:
2780    Exit Sub

ERRH:
2790    Select Case ERR.Number
        Case 2486  ' ** You can't carry out this action at the present time.
          ' ** Ignore.
2800    Case Else
2810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2820    End Select
2830    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2910    With Me
2920      If .cmdPricing_raised_focus_img.Visible = True Or .cmdPricing_raised_focus_dots_img.Visible = True Then
2930        Select Case blnPricing_Focus
            Case True
2940          .cmdPricing_raised_semifocus_dots_img.Visible = True
2950          .cmdPricing_raised_img.Visible = False
2960        Case False
2970          .cmdPricing_raised_img.Visible = True
2980          .cmdPricing_raised_semifocus_dots_img.Visible = False
2990        End Select
3000        .cmdPricing_raised_focus_img.Visible = False
3010        .cmdPricing_raised_focus_dots_img.Visible = False
3020        .cmdPricing_sunken_focus_dots_img.Visible = False
3030        .cmdPricing_raised_img_dis.Visible = False
3040      End If
3050      If .cmdPricing2_raised_focus_img.Visible = True Or .cmdPricing2_raised_focus_dots_img.Visible = True Then
3060        Select Case blnPricing_Focus
            Case True
3070          .cmdPricing2_raised_semifocus_dots_img.Visible = True
3080          .cmdPricing2_raised_img.Visible = False
3090        Case False
3100          .cmdPricing2_raised_img.Visible = True
3110          .cmdPricing2_raised_semifocus_dots_img.Visible = False
3120        End Select
3130        .cmdPricing2_raised_focus_img.Visible = False
3140        .cmdPricing2_raised_focus_dots_img.Visible = False
3150        .cmdPricing2_sunken_focus_dots_img.Visible = False
3160      End If
3170      If .cmdKey_raised_focus_dots_img.Visible = True Or .cmdKey_raised_focus_img.Visible = True Then
3180        Select Case blnKey_Focus
            Case True
3190          .cmdKey_raised_semifocus_dots_img.Visible = True
3200          .cmdKey_raised_img.Visible = False
3210        Case False
3220          .cmdKey_raised_img.Visible = True
3230          .cmdKey_raised_semifocus_dots_img.Visible = False
3240        End Select
3250        .cmdKey_raised_focus_img.Visible = False
3260        .cmdKey_raised_focus_dots_img.Visible = False
3270        .cmdKey_sunken_focus_dots_img.Visible = False
3280      End If
3290    End With

EXITP:
3300    Exit Sub

ERRH:
3310    Select Case ERR.Number
        Case Else
3320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3330    End Select
3340    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub cmdOK_Click()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strLicense As String, strVersion As String, dblVersion As Double
        Dim blnRetVal As Boolean

3510    blnRetVal = True

3520    With Me
3530      If blnPasswordOnly = True Then
3540        Set dbs = CurrentDb
3550        With dbs
3560          Set rst = .OpenRecordset("tblSecurity_License", dbOpenDynaset, dbConsistent)
3570          With rst
3580            If .BOF = True And .EOF = True Then
                  ' ** Unthinkable! Impossible!
3590              .AddNew
3600              ![seclic_licensedto] = Me.LicensedTo
3610              ![seclic_clientpath_ta] = CurrentAppPath  ' ** Module Function: modFileUtilities.
3620              ![seclic_datapath_ta] = CurrentBackendPath  ' ** Module Function: modFileUtilities.
3630              ![seclic_auxiliarypath] = CurrentBackendPath  ' ** Module Function: modFileUtilities.
3640              ![seclic_cycle] = EncodeString(CStr(Me.PasswordCycle))  ' ** Module Function: modCodeUtilities.
3650              ![seclic_cycle_screen] = EncodeString(CStr(Me.WarningScreen))  ' ** Module Function: modCodeUtilities.
3660              ![seclic_cycle_message] = EncodeString(CStr(Me.WarningMessage))  ' ** Module Function: modCodeUtilities.
3670              ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
3680              ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
3690              ![seclic_datemodified] = Now()
3700              .Update
3710            Else
3720              .MoveFirst
3730              .Edit
3740              ![seclic_cycle] = EncodeString(CStr(Me.PasswordCycle))  ' ** Module Function: modCodeUtilities.
3750              ![seclic_cycle_screen] = EncodeString(CStr(Me.WarningScreen))  ' ** Module Function: modCodeUtilities.
3760              ![seclic_cycle_message] = EncodeString(CStr(Me.WarningMessage))  ' ** Module Function: modCodeUtilities.
3770              ![seclic_user] = GetUserName  ' ** Module Function: modFileUtilities.
3780              ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
3790              ![seclic_datemodified] = Now()
3800              .Update
3810            End If
3820            .Close
3830          End With
3840          .Close
3850        End With
3860        MsgBox "Password Cycle Information Saved", vbInformation + vbOKOnly, "Password Cycle Entries Saved"
3870        .cmdQuit.Enabled = False
3880        .cmdQuit.Visible = False
3890        DoEvents
            ' ** Just close, don't wait to be to asked.
3900        blnDone = True
3910        DoCmd.Close acForm, THIS_NAME
3920      Else
3930        If IsNull(.LicensedTo) = True Then
3940          blnRetVal = False
3950          MsgBox "Must specify a Licensed To name.", vbExclamation + vbOKOnly, "Entry Required"
3960          .LicensedTo.SetFocus
3970        Else
3980          If Len(.LicensedTo) = 0 Then
3990            blnRetVal = False
4000            MsgBox "Must specify a Licensed To name.", vbExclamation + vbOKOnly, "Entry Required"
4010            .LicensedTo.SetFocus
4020          Else
4030            If Len(.LicensedTo) > 200 Then
4040              blnRetVal = False
4050              MsgBox "Licensed To name is too long!", vbExclamation + vbOKOnly, "Invalid Entry"
4060              .LicensedTo.SetFocus
4070            Else
4080              If IsNull(.RegistrationExpires) = True Then
4090                blnRetVal = False
4100                MsgBox "Must specify a Registration Expires date.", vbExclamation + vbOKOnly, "Date Required"
4110                .RegistrationExpires.SetFocus
4120              Else
4130                If Len(.RegistrationExpires) = 0 Then
4140                  blnRetVal = False
4150                  MsgBox "Must specify a Registration Expires Date.", vbExclamation + vbOKOnly, "Date Required"
4160                  .RegistrationExpires.SetFocus
4170                Else
4180                  If IsDate(.RegistrationExpires) = False Then
4190                    blnRetVal = False
4200                    MsgBox "Must specify a valid Registration Expires date.", vbExclamation + vbOKOnly, "Invalid Date"
4210                    .RegistrationExpires.SetFocus
4220                  Else
4230                    If CLng(CDate(.RegistrationExpires)) <= CLng(Date) Then
4240                      blnRetVal = False
4250                      MsgBox "Must specify a future date for Registration Expires.", vbExclamation + vbOKOnly, "Invalid Date"
4260                      .RegistrationExpires.SetFocus
4270                    Else
4280                      If IsNull(.RegistrationCode) = True Then
4290                        blnRetVal = False
4300                        MsgBox "Must specify a Registration Code.", vbExclamation + vbOKOnly, "Entry Required"
4310                        .RegistrationCode.SetFocus
4320                      Else
4330                        If Len(.RegistrationCode) = 0 Then
4340                          blnRetVal = False
4350                          MsgBox "Must specify a Registration Code.", vbExclamation + vbOKOnly, "Entry Required"
4360                          .RegistrationCode.SetFocus
4370                        End If
4380                      End If
4390                    End If
4400                  End If
4410                End If
4420              End If
4430            End If
4440          End If
4450        End If  ' ** LicensedTo.

4460        If blnRetVal = True Then
4470          .RegistrationExpires = Format(.RegistrationExpires, "mm/dd/yyyy")
              ' ** Because pricing is optional, make sure there are some codes to process, otherwise ignore these fields.
4480          If IsNull(.PricingCode) = True And IsNull(.PricingExpires) = True Then
                ' ** Skip them.
4490          Else
4500            If Nz(.PricingCode, vbNullString) = vbNullString And Nz(.PricingExpires, vbNullString) = vbNullString Then
                  ' ** Skip them.
4510            Else
4520              If IsNull(.PricingExpires) = True Then
4530                blnRetVal = False
4540                MsgBox "Must specify a Pricing Registration Expires date." & vbCrLf & vbCrLf & _
                      "If you did not purchase the Pricing Option," & vbCrLf & _
                      "clear both the Pricing Registration fields to continue.", vbExclamation + vbOKOnly, "Date Required"
4550                .PricingExpires.SetFocus
4560              Else
4570                If Len(.PricingExpires) = 0 Then
4580                  blnRetVal = False
4590                  MsgBox "Must specify a Pricing Registration Expires date.", vbExclamation + vbOKOnly, "Date Required"
4600                  .PricingExpires.SetFocus
4610                Else
4620                  If IsDate(.PricingExpires) = False Then
4630                    blnRetVal = False
4640                    MsgBox "Must specify a valid date for Pricing Registration Expires.", vbExclamation + vbOKOnly, "Invalid Date"
4650                    .PricingExpires.SetFocus
4660                  Else
4670                    If CLng(CDate(.PricingExpires)) <= CLng(Date) Then
4680                      blnRetVal = False
4690                      MsgBox "Must specify a future date for Pricing Registration Expires.", vbExclamation + vbOKOnly, "Invalid Date"
4700                      .PricingExpires.SetFocus
4710                    Else
4720                      If IsNull(.PricingCode) = True Then
4730                        blnRetVal = False
4740                        MsgBox "Must specify Pricing a Registration Code'." & vbCrLf & vbCrLf & _
                              "If you did not purchase the Pricing Option," & vbCrLf & _
                              "clear both the Pricing Registration fields to continue.", vbExclamation + vbOKOnly, "Entry Required"
4750                        .PricingCode.SetFocus
4760                      Else
4770                        If Len(.PricingCode) = 0 Then
4780                          blnRetVal = False
4790                          MsgBox "Must specify a Pricing Registration Code.", vbExclamation + vbOKOnly, "Entry Required"
4800                          .PricingCode.SetFocus
4810                        Else
4820                          .PricingExpires = Format(.PricingExpires, "mm/dd/yyyy")
4830                          If Val(.PricingCode) <> CalcPriceCode(.PricingExpires) Then
4840                            blnRetVal = False
4850                            MsgBox "Invalid Pricing Registration Code.", vbExclamation + vbOKOnly, "Invalid Entry"
4860                            .PricingCode.SetFocus
4870                            .PricingCode.SelStart = 0
4880                            .PricingCode.SelLength = Len(.PricingCode)
4890                          End If
4900                        End If
4910                      End If
4920                    End If
4930                  End If
4940                End If
4950              End If
4960            End If
4970          End If
4980        End If  ' ** blnRetVal.

4990        If blnRetVal = True Then
5000          If IsNull(.AccountLimit) = True Then
5010            blnRetVal = False
5020            MsgBox "Must specify an account Limit quantity.", vbExclamation + vbOKOnly, "Entry Required"
5030            .AccountLimit.SetFocus
5040          Else
5050            If Val(.AccountLimit) = 0 Then
5060              blnRetVal = False
5070              MsgBox "Invalid account Limit quantity.", vbExclamation + vbOKOnly, "Entry Required"
5080              .AccountLimit.SetFocus
5090              .AccountLimit.SelStart = 0
5100              .AccountLimit.SelLength = Len(.AccountLimit)
5110            Else
5120              If Val(.AccountLimit) < 0 Then .AccountLimit = Abs(Val(.AccountLimit))
5130              If Val(.AccountLimit) > 2000000 Then
5140                blnRetVal = False
5150                MsgBox "The account Limit quantity is too high!", vbExclamation + vbOKOnly, "Invalid Entry"
5160                .AccountLimit.SetFocus
5170                .AccountLimit.SelStart = 0
5180                .AccountLimit.SelLength = Len(.AccountLimit)
5190              Else
5200                If IsNull(.LicenseCode) = True Then
5210                  blnRetVal = False
5220                  MsgBox "Must specify a License Code.", vbExclamation + vbOKOnly, "Entry Required"
5230                  .LicenseCode.SetFocus
5240                Else
5250                  If Len(.LicenseCode) = 0 Then
5260                    blnRetVal = False
5270                    MsgBox "Must specify a License Code.", vbExclamation + vbOKOnly, "Entry Required"
5280                    .LicenseCode.SetFocus
5290                  Else
5300                    If Val(.RegistrationCode) <> CalcRegCode(.RegistrationExpires) Then
5310                      blnRetVal = False
5320                      MsgBox "Invalid Registration Code.", vbExclamation + vbOKOnly, "Invalid Entry"
5330                      .RegistrationCode.SetFocus
5340                      .RegistrationCode.SelStart = 0
5350                      .RegistrationCode.SelLength = Len(.RegistrationCode)
5360                    Else
5370                      If .AccountLimit <> CalcLimitNumber(.LicenseCode) Then
5380                        blnRetVal = False
5390                        MsgBox "Invalid License Code.", vbExclamation + vbOKOnly, "Invalid Entry"
5400                        .LicenseCode.SetFocus
5410                        .LicenseCode.SelStart = 0
5420                        .LicenseCode.SelLength = Len(.LicenseCode)
5430                      End If
5440                    End If
5450                  End If
5460                End If
5470              End If
5480            End If
5490          End If
5500        End If  ' ** blnRetVal.

5510        If blnRetVal = True Then
              ' ** Save settings.

              ' ** Licensed To.
5520          If .LicensedTo.Enabled = True Then
5530            blnRetVal = IniFile_Set("License", "Firm", EncodeString(.LicensedTo), _
                  gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modStartupFuncs.
5540            Set dbs = CurrentDb
5550            With dbs
                  ' ** Update License Name, by specified [licnam].
5560              Set qdf = .QueryDefs("qrySecurity_License_09")
5570              With qdf.Parameters
5580                ![licnam] = Me.LicensedTo
5590              End With
5600              qdf.Execute
                  ' ** Update tblSecurity_License, by specified [licnam].
5610              Set qdf = .QueryDefs("qrySecurity_License_10")
5620              With qdf.Parameters
5630                ![licnam] = Me.LicensedTo
5640              End With
5650              qdf.Execute
5660              .Close
5670            End With
5680          End If
              ' ** Limit code.
5690          If Len(.AccountLimit) > 0 And Len(.LicenseCode) > 0 Then
5700            If Len(.AccountLimit) > 0 Then
5710              blnRetVal = IniFile_Set("License", "Limit", EncodeString(.AccountLimit), _
                    gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modStartupFuncs.
5720            End If
5730            If .AccountLimit = "0" Then
5740              blnRetVal = IniFile_Del("License", "Limit", _
                    gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modStartupFuncs.
5750            End If
5760          End If
              ' ** Expires.
5770          blnRetVal = IniFile_Set("License", "Expires", EncodeString(.RegistrationExpires), _
                gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modStartupFuncs.
              ' ** Password Cycle.
5780          Set dbs = CurrentDb
5790          With dbs
                ' ** Update tblSecurity_License, by specified [liccyc], [licscr], [licmsg].
5800            Set qdf = .QueryDefs("qrySecurity_License_11")
5810            With qdf.Parameters
5820              ![liccyc] = EncodeString(CStr(Me.PasswordCycle))  ' ** Module Function: modCodeUtilities.
5830              ![licscr] = EncodeString(CStr(Me.WarningScreen))  ' ** Module Function: modCodeUtilities.
5840              ![licmsg] = EncodeString(CStr(Me.WarningMessage))  ' ** Module Function: modCodeUtilities.
5850            End With
5860            qdf.Execute
5870            .Close
5880          End With
              ' ** Pricing Expires.
5890          If Len(.PricingExpires) <> 0 Then
5900            blnRetVal = IniFile_Set("License", "Pricing", EncodeString(.PricingExpires), _
                  gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modStartupFuncs.
5910          Else
5920            blnRetVal = IniFile_Del("License", "Pricing", _
                  gstrTrustDataLocation & gstrFile_LIC)  ' ** Module Procedure: modStartupFuncs.
5930          End If

5940          dblVersion = Nz(DLookup("Version", "[License Name]"), 0)
5950          strVersion = CStr(dblVersion)
5960          If InStr(strVersion, ".") > 0 Then
5970            If Len(Mid(strVersion, (InStr(strVersion, ".") + 1))) > 1 Then
5980              strVersion = Left(strVersion, InStr(strVersion, ".")) & _
                    Mid(strVersion, (InStr(strVersion, ".") + 1), 1) & "." & Mid(strVersion, (InStr(strVersion, ".") + 2))
5990            Else
6000              strVersion = strVersion & ".00"
6010            End If
6020          Else
6030            strVersion = strVersion & ".0.00"
6040          End If

6050          strLicense = DecodeString(IniFile_Get("License", "Firm", EncodeString("Call Delta Data, Inc., for Licensing info."), _
                gstrTrustDataLocation & gstrFile_LIC))  ' ** Module Procedure: modCodeUtilities, modStartupFuncs.
6060          strLicense = "Licensed to " & strLicense & "    Version " & _
                IIf(Left(Right(strVersion, 2), 1) = ".", strVersion & "0", strVersion)
6070          If IsLoaded("frmMenu_Title", acForm) = True Then  ' ** Module Function: modFileUtilities.
6080            Forms("frmMenu_Title").LicenseInfo = strLicense
6090            DoEvents
6100          End If

6110          MsgBox "License Information Saved", vbInformation + vbOKOnly, "License Entries Saved"

6120          If Security_LicenseChk = True Then  ' ** Module Procedure: modSecurityFunctions.
6130            .cmdQuit.Enabled = False
6140            .cmdQuit.Visible = False
6150            DoEvents
                ' ** Just close, don't wait to be to asked.
6160            blnDone = True
6170            DoCmd.Close acForm, THIS_NAME
6180          Else
6190            Beep
6200            MsgBox "The license information provided did" & vbCrLf & _
                  "not pass a subsequent license check." & vbCrLf & vbCrLf & _
                  "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "License Check Failed"
6210            cmdQuit_Click  ' ** Procedure: Below.
6220          End If

6230        End If  ' ** blnRetVal.

6240      End If  ' ** blnPasswordOnly.

6250    End With  ' ** Me.

EXITP:
6260    Set rst = Nothing
6270    Set qdf = Nothing
6280    Set dbs = Nothing
6290    Exit Sub

ERRH:
6300    Select Case ERR.Number
        Case Else
6310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6320    End Select
6330    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6420    intShiftDown = (Shift And acShiftMask) > 0
6430    intAltDown = (Shift And acAltMask) > 0
6440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6460      Select Case intRetVal
          Case vbKeyTab
6470        With Me
6480          intRetVal = 0
6490          If .cmdCancel.Enabled = True And .cmdCancel.Visible = True Then
6500            .cmdCancel.SetFocus
6510          ElseIf .cmdQuit.Enabled = True And .cmdQuit.Visible = True Then
6520            .cmdQuit.SetFocus
6530          Else
6540            .LicensedTo.SetFocus
6550          End If
6560        End With
6570      End Select
6580    End If

        ' ** Shift keys.
6590    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6600      Select Case intRetVal
          Case vbKeyTab
6610        With Me
6620          intRetVal = 0
6630          If .PricingCode.Enabled = True Then
6640            .PricingCode.SetFocus
6650          Else
6660            .WarningMessage.SetFocus
6670          End If
6680        End With
6690      End Select
6700    End If

EXITP:
6710    KeyCode = intRetVal
6720    Exit Sub

ERRH:
6730    intRetVal = 0
6740    Select Case ERR.Number
        Case Else
6750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6760    End Select
6770    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

6810    blnCancel = True
6820    gblnClosing = False
6830    DoCmd.Close acForm, THIS_NAME

EXITP:
6840    Exit Sub

ERRH:
6850    Select Case ERR.Number
        Case Else
6860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6870    End Select
6880    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
6990          .LicensedTo.SetFocus
7000        End With
7010      End Select
7020    End If

        ' ** Shift keys.
7030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7040      Select Case intRetVal
          Case vbKeyTab
7050        With Me
7060          intRetVal = 0
7070          If .cmdOK.Enabled = True Then
7080            .cmdOK.SetFocus
7090          ElseIf .PricingCode.Enabled = True Then
7100            .PricingCode.SetFocus
7110          Else
7120            .WarningMessage.SetFocus
7130          End If
7140        End With
7150      End Select
7160    End If

EXITP:
7170    KeyCode = intRetVal
7180    Exit Sub

ERRH:
7190    intRetVal = 0
7200    Select Case ERR.Number
        Case Else
7210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7220    End Select
7230    Resume EXITP

End Sub

Private Sub cmdQuit_Click()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdQuit_Click"

7310    blnQuitInvoked = True
7320    gblnClosing = False
7330    QuitNow  ' ** Module Procedure: modStartupFuncs.

EXITP:
7340    Exit Sub

ERRH:
7350    Select Case ERR.Number
        Case Else
7360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7370    End Select
7380    Resume EXITP

End Sub

Private Sub cmdQuit_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdQuit_KeyDown"

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
          Case vbKeyTab
7470        With Me
7480          intRetVal = 0
7490          .LicensedTo.SetFocus
7500        End With
7510      End Select
7520    End If

        ' ** Shift keys.
7530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7540      Select Case intRetVal
          Case vbKeyTab
7550        With Me
7560          intRetVal = 0
7570          If .cmdOK.Enabled = True Then
7580            .cmdOK.SetFocus
7590          ElseIf .PricingCode.Enabled = True Then
7600            .PricingCode.SetFocus
7610          Else
7620            .WarningMessage.SetFocus
7630          End If
7640        End With
7650      End Select
7660    End If

EXITP:
7670    KeyCode = intRetVal
7680    Exit Sub

ERRH:
7690    intRetVal = 0
7700    Select Case ERR.Number
        Case Else
7710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7720    End Select
7730    Resume EXITP

End Sub

Private Sub cmdKey_GotFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdKey_GotFocus"

7810    With Me
7820      blnKey_Focus = True
7830      .cmdKey_raised_semifocus_dots_img.Visible = True
7840      .cmdKey_raised_img.Visible = False
7850      .cmdKey_raised_focus_img.Visible = False
7860      .cmdKey_raised_focus_dots_img.Visible = False
7870      .cmdKey_sunken_focus_dots_img.Visible = False
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

Private Sub cmdKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdKey_MouseDown"

8010    With Me
8020      blnKey_MouseDown = True
8030      .cmdKey_sunken_focus_dots_img.Visible = True
8040      .cmdKey_raised_img.Visible = False
8050      .cmdKey_raised_semifocus_dots_img.Visible = False
8060      .cmdKey_raised_focus_img.Visible = False
8070      .cmdKey_raised_focus_dots_img.Visible = False
8080    End With

EXITP:
8090    Exit Sub

ERRH:
8100    Select Case ERR.Number
        Case Else
8110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8120    End Select
8130    Resume EXITP

End Sub

Private Sub cmdKey_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdKey_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8220    intShiftDown = (Shift And acShiftMask) > 0
8230    intAltDown = (Shift And acAltMask) > 0
8240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8260      Select Case intRetVal
          Case vbKeyTab
8270        With Me
8280          intRetVal = 0
8290          If .cmdOK.Enabled = True Then
8300            .cmdOK.SetFocus
8310          ElseIf .cmdCancel.Enabled = True And .cmdCancel.Visible = True Then
8320            .cmdCancel.SetFocus
8330          ElseIf .cmdQuit.Enabled = True And .cmdQuit.Visible = True Then
8340            .cmdQuit.SetFocus
8350          Else
8360            .LicensedTo.SetFocus
8370          End If
8380        End With
8390      End Select
8400    End If

        ' ** Shift keys.
8410    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8420      Select Case intRetVal
          Case vbKeyTab
8430        With Me
8440          intRetVal = 0
8450          If .PricingCode.Enabled = True Then
8460            .PricingCode.SetFocus
8470          Else
8480            .WarningMessage.SetFocus
8490          End If
8500        End With
8510      End Select
8520    End If

EXITP:
8530    KeyCode = intRetVal
8540    Exit Sub

ERRH:
8550    intRetVal = 0
8560    Select Case ERR.Number
        Case Else
8570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8580    End Select
8590    Resume EXITP

End Sub

Private Sub cmdKey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdKey_MouseMove"

8610    With Me
8620      If blnKey_MouseDown = False Then
8630        Select Case blnKey_Focus
            Case True
8640          .cmdKey_raised_focus_dots_img.Visible = True
8650          .cmdKey_raised_focus_img.Visible = False
8660        Case False
8670          .cmdKey_raised_focus_img.Visible = True
8680          .cmdKey_raised_focus_dots_img.Visible = False
8690        End Select
8700        .cmdKey_raised_img.Visible = False
8710        .cmdKey_raised_semifocus_dots_img.Visible = False
8720        .cmdKey_sunken_focus_dots_img.Visible = False
8730      End If
8740    End With

EXITP:
8750    Exit Sub

ERRH:
8760    Select Case ERR.Number
        Case Else
8770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8780    End Select
8790    Resume EXITP

End Sub

Private Sub cmdKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ** Easter Egg (Secret Key) to open frmLicense_Edit.

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdKey_MouseUp"

8810    With Me

8820      .cmdKey_raised_focus_dots_img.Visible = True
8830      .cmdKey_raised_img.Visible = False
8840      .cmdKey_raised_semifocus_dots_img.Visible = False
8850      .cmdKey_raised_focus_img.Visible = False
8860      .cmdKey_sunken_focus_dots_img.Visible = False

8870      If Button = acRightButton And Shift = acCtrlMask + acAltMask Then
8880        DoCmd.OpenForm "frmLicense_Edit", , , , , acDialog, THIS_NAME & "~" & strCallingForm
            ' ** If this originally came from InitializeTables(), and frmLicense_Edit
            ' ** was canceled, gblnClosing will have been set to False.
8890        If gblnClosing = False Then
8900          blnCancel = True
8910        End If
8920        DoCmd.Close acForm, THIS_NAME
8930      End If

8940      blnKey_MouseDown = False
8950    End With

EXITP:
8960    Exit Sub

ERRH:
8970    Select Case ERR.Number
        Case Else
8980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8990    End Select
9000    Resume EXITP

End Sub

Private Sub cmdKey_LostFocus()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdKey_LostFocus"

9110    With Me
9120      .cmdKey_raised_img.Visible = True
9130      .cmdKey_raised_semifocus_dots_img.Visible = False
9140      .cmdKey_raised_focus_img.Visible = False
9150      .cmdKey_raised_focus_dots_img.Visible = False
9160      .cmdKey_sunken_focus_dots_img.Visible = False
9170      blnKey_Focus = False
9180    End With

EXITP:
9190    Exit Sub

ERRH:
9200    Select Case ERR.Number
        Case Else
9210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9220    End Select
9230    Resume EXITP

End Sub

Private Sub cmdSave_Click()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

9310    cmdOK_Click  ' ** Procedure: Above.

EXITP:
9320    Exit Sub

ERRH:
9330    Select Case ERR.Number
        Case Else
9340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9350    End Select
9360    Resume EXITP

End Sub

Private Sub LicensedTo_AfterUpdate()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "LicensedTo_AfterUpdate"

9410    With Me
9420      blnPasswordOnly = False
9430      If blnFromUtilityMenu = True Then
9440        If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
9450      End If
9460    End With

EXITP:
9470    Exit Sub

ERRH:
9480    THAT_PROC = THIS_PROC
9490    That_Erl = Erl
9500    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9510    Resume EXITP

End Sub

Private Sub LicensedTo_GotFocus()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "LicensedTo_GotFocus"

9610    With Me
9620      .LicensedTo.SelLength = 0
9630      .LicensedTo.SelStart = 99
9640    End With

EXITP:
9650    Exit Sub

ERRH:
9660    THAT_PROC = THIS_PROC
9670    That_Erl = Erl
9680    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9690    Resume EXITP

End Sub

Private Sub LicensedTo_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "LicensedTo_KeyDown"

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
9790          .RegistrationExpires.SetFocus
9800        End With
9810      End Select
9820    End If

        ' ** Shift keys.
9830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9850        With Me
9860          intRetVal = 0
9870          If .cmdCancel.Enabled = True And .cmdCancel.Visible = True Then
9880            .cmdCancel.SetFocus
9890          ElseIf .cmdQuit.Enabled = True And .cmdQuit.Visible = True Then
9900            .cmdQuit.SetFocus
9910          Else
9920            .cmdOK.SetFocus
9930          End If
9940        End With
9950      End Select
9960    End If

EXITP:
9970    KeyCode = intRetVal
9980    Exit Sub

ERRH:
9990    intRetVal = 0
10000   Select Case ERR.Number
        Case Else
10010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10020   End Select
10030   Resume EXITP

End Sub

Private Sub RegistrationExpires_AfterUpdate()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "RegistrationExpires_AfterUpdate"

10110   With Me
10120     blnPasswordOnly = False
10130     If blnFromUtilityMenu = True Then
10140       If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
10150     End If
10160   End With

EXITP:
10170   Exit Sub

ERRH:
10180   THAT_PROC = THIS_PROC
10190   That_Erl = Erl
10200   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10210   Resume EXITP

End Sub

Private Sub RegistrationExpires_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "RegistrationExpires_MouseUp"

10310   With Me
10320     .RegistrationExpires.SelStart = 0
10330     .RegistrationExpires.SelLength = 0
10340   End With

EXITP:
10350   Exit Sub

ERRH:
10360   THAT_PROC = THIS_PROC
10370   That_Erl = Erl
10380   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10390   Resume EXITP

End Sub

Private Sub RegistrationCode_AfterUpdate()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "RegistrationCode_AfterUpdate"

10410   With Me
10420     blnPasswordOnly = False
10430     If blnFromUtilityMenu = True Then
10440       If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
10450     End If
10460   End With

EXITP:
10470   Exit Sub

ERRH:
10480   THAT_PROC = THIS_PROC
10490   That_Erl = Erl
10500   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10510   Resume EXITP

End Sub

Private Sub AccountLimit_AfterUpdate()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "AccountLimit_AfterUpdate"

10610   With Me
10620     blnPasswordOnly = False
10630     If blnFromUtilityMenu = True Then
10640       If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
10650     End If
10660   End With

EXITP:
10670   Exit Sub

ERRH:
10680   THAT_PROC = THIS_PROC
10690   That_Erl = Erl
10700   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10710   Resume EXITP

End Sub

Private Sub LicenseCode_AfterUpdate()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "LicenseCode_AfterUpdate"

10810   With Me
10820     blnPasswordOnly = False
10830     If blnFromUtilityMenu = True Then
10840       If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
10850     End If
10860   End With

EXITP:
10870   Exit Sub

ERRH:
10880   THAT_PROC = THIS_PROC
10890   That_Erl = Erl
10900   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10910   Resume EXITP

End Sub

Private Sub Password_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "Password_box_MouseMove"

11010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11020   Exit Sub

ERRH:
11030   Select Case ERR.Number
        Case Else
11040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11050   End Select
11060   Resume EXITP

End Sub

Private Sub PasswordCycle_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "PasswordCycle_AfterUpdate"

11110   With Me
11120     If blnFromUtilityMenu = True Then
11130       If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
11140     End If
11150     If IsNull(.PasswordCycle) = False Then
11160       If CLng(.PasswordCycle) < 0 Then
11170         MsgBox "The Password Cycle cannot be a negative number.", vbInformation + vbOKOnly, "Invalid Entry"
11180         .PasswordCycle = 0& 'DEF_CYCLE
11190       ElseIf CLng(.PasswordCycle) = 0& Then
11200         .WarningScreen = 0&
11210         .WarningMessage = 0&
11220       ElseIf CLng(.PasswordCycle) < 7& Then
11230         If MsgBox("Are you sure you want passwords to expire every " & _
                  CStr(CLng(.PasswordCycle)) & " days?", vbQuestion + vbYesNo + vbDefaultButton2, _
                  "Short Password Cycle") <> vbYes Then
11240           .PasswordCycle = 0& 'DEF_CYCLE
11250         End If
11260       ElseIf CLng(.PasswordCycle) > 1825& Then  ' ** 5 years.
11270         MsgBox "Your entry will be changed to 0, for no password expiration.", _
                vbInformation + vbOKOnly, "Passwords Will Not Expire"
11280         .PasswordCycle = 0&
11290         .WarningScreen = 0&
11300         .WarningMessage = 0&
11310       End If
11320     Else
11330       .PasswordCycle = 0&
11340       .WarningScreen = 0&
11350       .WarningMessage = 0&
11360     End If
11370   End With

EXITP:
11380   Exit Sub

ERRH:
11390   THAT_PROC = THIS_PROC
11400   That_Erl = Erl
11410   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11420   Resume EXITP

End Sub

Private Sub WarningScreen_AfterUpdate()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "WarningScreen_AfterUpdate"

11510   With Me
11520     If blnFromUtilityMenu = True Then
11530       If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
11540     End If
11550     If CLng(.PasswordCycle) > 0& Then
11560       If IsNull(.WarningScreen) = False Then
11570         If CLng(.WarningScreen) < 0 Then
11580           MsgBox "The On-Screen Notice cannot be a negative number.", vbInformation + vbOKOnly, "Invalid Entry"
11590           If CLng(.PasswordCycle) > DEF_SCREEN Then
11600             .WarningScreen = DEF_SCREEN
11610           ElseIf CLng(.PasswordCycle) > 3& Then
11620             .WarningScreen = Int(CLng(.PasswordCycle) * 0.25)
11630           Else
11640             .WarningScreen = 1&
11650           End If
11660         ElseIf CLng(.WarningScreen) > CLng(.PasswordCycle) Then
11670           MsgBox "The On-Screen Notice cannot be greater than the Password Cycle.", _
                  vbInformation + vbOKOnly, "Invalid Entry"
11680           If CLng(.PasswordCycle) > DEF_SCREEN Then
11690             .WarningScreen = DEF_SCREEN
11700           ElseIf CLng(.PasswordCycle) > 3& Then
11710             .WarningScreen = Int(CLng(.PasswordCycle) * 0.25)
11720           Else
11730             .WarningScreen = 1&
11740           End If
11750         End If
11760       Else
11770         .WarningScreen = 0&
11780       End If
11790     Else
11800       .WarningScreen = 0&
11810     End If
11820   End With

EXITP:
11830   Exit Sub

ERRH:
11840   THAT_PROC = THIS_PROC
11850   That_Erl = Erl
11860   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11870   Resume EXITP

End Sub

Private Sub WarningMessage_AfterUpdate()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "WarningMessage_AfterUpdate"

11910   With Me
11920     If blnFromUtilityMenu = True Then
11930       If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
11940     End If
11950     If CLng(.PasswordCycle) > 0& Then
11960       If IsNull(.WarningMessage) = False Then
11970         If CLng(.WarningMessage) < 0 Then
11980           MsgBox "The Pop-Up Notice cannot be a negative number.", vbInformation + vbOKOnly, "Invalid Entry"
11990           If CLng(.PasswordCycle) > DEF_MSGBOX Then
12000             .WarningMessage = DEF_MSGBOX
12010           ElseIf CLng(.PasswordCycle) > 3& Then
12020             .WarningMessage = Int(CLng(.PasswordCycle) * 0.25)
12030           Else
12040             .WarningMessage = 1&
12050           End If
12060         ElseIf CLng(.WarningMessage) > CLng(.PasswordCycle) Then
12070           MsgBox "The Pop-Up Notice cannot be greater than the Password Cycle.", _
                  vbInformation + vbOKOnly, "Invalid Entry"
12080           If CLng(.PasswordCycle) > DEF_MSGBOX Then
12090             .WarningMessage = DEF_MSGBOX
12100           ElseIf CLng(.PasswordCycle) > 3& Then
12110             .WarningMessage = Int(CLng(.PasswordCycle) * 0.25)
12120           Else
12130             .WarningMessage = 1&
12140           End If
12150         End If
12160       Else
12170         .WarningMessage = 0&
12180       End If
12190     Else
12200       .WarningMessage = 0&
12210     End If
12220   End With

EXITP:
12230   Exit Sub

ERRH:
12240   THAT_PROC = THIS_PROC
12250   That_Erl = Erl
12260   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12270   Resume EXITP

End Sub

Private Sub WarningMessage_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "WarningMessage_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12320   intShiftDown = (Shift And acShiftMask) > 0
12330   intAltDown = (Shift And acAltMask) > 0
12340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12370       With Me
12380         intRetVal = 0
12390         If .PricingExpires.Enabled = True Then
12400           .PricingExpires.SetFocus
12410         ElseIf .cmdOK.Enabled = True Then
12420           .cmdOK.SetFocus
12430         ElseIf .cmdCancel.Enabled = True And .cmdCancel.Visible = True Then
12440           .cmdCancel.SetFocus
12450         Else
12460           .cmdQuit.SetFocus
12470         End If
12480       End With
12490     End Select
12500   End If

        ' ** Shift keys.
12510   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12520     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12530       With Me
12540         intRetVal = 0
12550         .WarningScreen.SetFocus
12560       End With
12570     End Select
12580   End If

EXITP:
12590   KeyCode = intRetVal
12600   Exit Sub

ERRH:
12610   intRetVal = 0
12620   Select Case ERR.Number
        Case Else
12630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12640   End Select
12650   Resume EXITP

End Sub

Private Sub cmdReset_Click()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_Click"

12710   With Me
12720     .PasswordCycle = 0& 'DEF_CYCLE
12730     .WarningScreen = 0& 'DEF_SCREEN
12740     .WarningMessage = 0& 'DEF_MSGBOX
12750     If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
12760   End With

EXITP:
12770   Exit Sub

ERRH:
12780   THAT_PROC = THIS_PROC
12790   That_Erl = Erl
12800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12810   Resume EXITP

End Sub

Private Sub Pricing_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "Pricing_Box_MouseMove"

12910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12920   Exit Sub

ERRH:
12930   Select Case ERR.Number
        Case Else
12940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12950   End Select
12960   Resume EXITP

End Sub

Private Sub cmdPricing_GotFocus()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPricing_GotFocus"

13010   With Me
13020     blnPricing_Focus = True
13030     Select Case .Pricing_Box.Visible
          Case True
13040       .cmdPricing2_raised_semifocus_dots_img.Visible = True
13050       .cmdPricing2_raised_img.Visible = False
13060       .cmdPricing2_raised_focus_img.Visible = False
13070       .cmdPricing2_raised_focus_dots_img.Visible = False
13080       .cmdPricing2_sunken_focus_dots_img.Visible = False
13090     Case False
13100       .cmdPricing_raised_semifocus_dots_img.Visible = True
13110       .cmdPricing_raised_img.Visible = False
13120       .cmdPricing_raised_focus_img.Visible = False
13130       .cmdPricing_raised_focus_dots_img.Visible = False
13140       .cmdPricing_sunken_focus_dots_img.Visible = False
13150       .cmdPricing_raised_img_dis.Visible = False
13160     End Select
13170   End With

EXITP:
13180   Exit Sub

ERRH:
13190   Select Case ERR.Number
        Case Else
13200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13210   End Select
13220   Resume EXITP

End Sub

Private Sub cmdPricing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPricing_MouseDown"

13310   With Me
13320     blnPricing_MouseDown = True
13330     Select Case .Pricing_Box.Visible
          Case True
13340       .cmdPricing2_sunken_focus_dots_img.Visible = True
13350       .cmdPricing2_raised_img.Visible = False
13360       .cmdPricing2_raised_semifocus_dots_img.Visible = False
13370       .cmdPricing2_raised_focus_img.Visible = False
13380       .cmdPricing2_raised_focus_dots_img.Visible = False
13390     Case False
13400       .cmdPricing_sunken_focus_dots_img.Visible = True
13410       .cmdPricing_raised_img.Visible = False
13420       .cmdPricing_raised_semifocus_dots_img.Visible = False
13430       .cmdPricing_raised_focus_img.Visible = False
13440       .cmdPricing_raised_focus_dots_img.Visible = False
13450       .cmdPricing_raised_img_dis.Visible = False
13460     End Select
13470   End With

EXITP:
13480   Exit Sub

ERRH:
13490   Select Case ERR.Number
        Case Else
13500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13510   End Select
13520   Resume EXITP

End Sub

Private Sub cmdPricing_Click()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPricing_Click"

13610   With Me
13620     Select Case .Pricing_Box.Visible
          Case True
13630       .PricingExpires.Enabled = False
13640       .PricingExpires.Visible = False
13650       .PricingCode.Enabled = False
13660       .PricingCode.Visible = False
13670       .Pricing_hline01.Visible = False
13680       .Pricing_hline02.Visible = False
13690       .Pricing_hline03.Visible = False
13700       .Pricing_vline01.Visible = False
13710       .Pricing_vline02.Visible = False
13720       .Pricing_vline03.Visible = False
13730       .Pricing_vline04.Visible = False
13740       .Pricing_Box.Visible = False
13750       .Pricing_Box2.Visible = False
13760       .cmdPricing2_raised_img.Visible = False
13770       .cmdPricing2_raised_semifocus_dots_img.Visible = False
13780       .cmdPricing2_raised_focus_img.Visible = False
13790       .cmdPricing2_raised_focus_dots_img.Visible = False
13800       .cmdPricing2_sunken_focus_dots_img.Visible = False
13810       .cmdPricing_raised_focus_dots_img.Visible = True
13820       .cmdPricing_raised_img.Visible = False
13830       .cmdPricing_raised_semifocus_dots_img.Visible = False
13840       .cmdPricing_raised_focus_img.Visible = False
13850       .cmdPricing_sunken_focus_dots_img.Visible = False
13860       .cmdPricing_raised_img_dis.Visible = False
13870       .cmdPricing.Width = .cmdPricing_raised_img.Width
13880       If .cmdOK.Enabled = True Then
13890         .cmdOK.SetFocus
13900       Else
13910         .cmdCancel.SetFocus
13920       End If
13930     Case False
13940       .cmdPricing_raised_img.Visible = False
13950       .cmdPricing_raised_semifocus_dots_img.Visible = False
13960       .cmdPricing_raised_focus_img.Visible = False
13970       .cmdPricing_raised_focus_dots_img.Visible = False
13980       .cmdPricing_sunken_focus_dots_img.Visible = False
13990       .cmdPricing_raised_img_dis.Visible = False
14000       .Pricing_Box.Visible = True
14010       .Pricing_Box2.Visible = True
14020       .cmdPricing2_raised_img.Visible = True
14030       .cmdPricing.Width = .cmdPricing2_raised_img.Width
14040       .PricingExpires.Enabled = True
14050       .PricingExpires.Visible = True
14060       .PricingCode.Enabled = True
14070       .PricingCode.Visible = True
14080       .Pricing_hline01.Visible = True
14090       .Pricing_hline02.Visible = True
14100       .Pricing_hline03.Visible = True
14110       .Pricing_vline01.Visible = True
14120       .Pricing_vline02.Visible = True
14130       .Pricing_vline03.Visible = True
14140       .Pricing_vline04.Visible = True
14150       .PricingExpires.SetFocus
14160     End Select
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

Private Sub cmdPricing_KeyDown(KeyCode As Integer, Shift As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPricing_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14320   intShiftDown = (Shift And acShiftMask) > 0
14330   intAltDown = (Shift And acAltMask) > 0
14340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14360     Select Case intRetVal
          Case vbKeyTab
14370       With Me
14380         intRetVal = 0
14390         If .PricingExpires.Enabled = True Then
14400           .PricingExpires.SetFocus
14410         ElseIf .cmdOK.Enabled = True Then
14420           .cmdOK.SetFocus
14430         ElseIf .cmdCancel.Enabled = True And .cmdCancel.Visible = True Then
14440           .cmdCancel.SetFocus
14450         Else
14460           .cmdQuit.SetFocus
14470         End If
14480       End With
14490     End Select
14500   End If

        ' ** Shift keys.
14510   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14520     Select Case intRetVal
          Case vbKeyTab
14530       With Me
14540         intRetVal = 0
14550         .cmdReset.SetFocus
14560       End With
14570     End Select
14580   End If

EXITP:
14590   KeyCode = intRetVal
14600   Exit Sub

ERRH:
14610   intRetVal = 0
14620   Select Case ERR.Number
        Case Else
14630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14640   End Select
14650   Resume EXITP

End Sub

Private Sub cmdPricing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPricing_MouseMove"

14710   With Me
14720     If blnPricing_MouseDown = False Then
14730       Select Case .Pricing_Box.Visible
            Case True
14740         Select Case blnPricing_Focus
              Case True
14750           .cmdPricing2_raised_focus_dots_img.Visible = True
14760           .cmdPricing2_raised_focus_img.Visible = False
14770         Case False
14780           .cmdPricing2_raised_focus_img.Visible = True
14790           .cmdPricing2_raised_focus_dots_img.Visible = False
14800         End Select
14810         .cmdPricing2_raised_img.Visible = False
14820         .cmdPricing2_raised_semifocus_dots_img.Visible = False
14830         .cmdPricing2_sunken_focus_dots_img.Visible = False
14840       Case False
14850         Select Case blnPricing_Focus
              Case True
14860           .cmdPricing_raised_focus_dots_img.Visible = True
14870           .cmdPricing_raised_focus_img.Visible = False
14880         Case False
14890           .cmdPricing_raised_focus_img.Visible = True
14900           .cmdPricing_raised_focus_dots_img.Visible = False
14910         End Select
14920         .cmdPricing_raised_img.Visible = False
14930         .cmdPricing_raised_semifocus_dots_img.Visible = False
14940         .cmdPricing_sunken_focus_dots_img.Visible = False
14950         .cmdPricing_raised_img_dis.Visible = False
14960       End Select
14970     End If
14980   End With

EXITP:
14990   Exit Sub

ERRH:
15000   Select Case ERR.Number
        Case Else
15010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15020   End Select
15030   Resume EXITP

End Sub

Private Sub cmdPricing_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPricing_MouseUp"

15110   With Me
15120     Select Case .Pricing_Box.Visible
          Case True
15130       .cmdPricing2_raised_focus_dots_img.Visible = True
15140       .cmdPricing2_raised_img.Visible = False
15150       .cmdPricing2_raised_semifocus_dots_img.Visible = False
15160       .cmdPricing2_raised_focus_img.Visible = False
15170       .cmdPricing2_sunken_focus_dots_img.Visible = False
15180     Case False
15190       .cmdPricing_raised_focus_dots_img.Visible = True
15200       .cmdPricing_raised_img.Visible = False
15210       .cmdPricing_raised_semifocus_dots_img.Visible = False
15220       .cmdPricing_raised_focus_img.Visible = False
15230       .cmdPricing_sunken_focus_dots_img.Visible = False
15240       .cmdPricing_raised_img_dis.Visible = False
15250     End Select
15260     blnPricing_MouseDown = False
15270   End With

EXITP:
15280   Exit Sub

ERRH:
15290   Select Case ERR.Number
        Case Else
15300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15310   End Select
15320   Resume EXITP

End Sub

Private Sub cmdPricing_LostFocus()

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPricing_LostFocus"

15410   With Me
15420     Select Case .Pricing_Box.Visible
          Case True
15430       .cmdPricing2_raised_img.Visible = True
15440       .cmdPricing2_raised_semifocus_dots_img.Visible = False
15450       .cmdPricing2_raised_focus_img.Visible = False
15460       .cmdPricing2_raised_focus_dots_img.Visible = False
15470       .cmdPricing2_sunken_focus_dots_img.Visible = False
15480     Case False
15490       .cmdPricing_raised_img.Visible = True
15500       .cmdPricing_raised_semifocus_dots_img.Visible = False
15510       .cmdPricing_raised_focus_img.Visible = False
15520       .cmdPricing_raised_focus_dots_img.Visible = False
15530       .cmdPricing_sunken_focus_dots_img.Visible = False
15540       .cmdPricing_raised_img_dis.Visible = False
15550     End Select
15560     blnPricing_Focus = False
15570   End With

EXITP:
15580   Exit Sub

ERRH:
15590   Select Case ERR.Number
        Case Else
15600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15610   End Select
15620   Resume EXITP

End Sub

Private Sub PricingExpires_AfterUpdate()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "PricingExpires_AfterUpdate"

15710   With Me
15720     blnPasswordOnly = False
15730     If blnFromUtilityMenu = True Then
15740       If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
15750     End If
15760   End With

EXITP:
15770   Exit Sub

ERRH:
15780   THAT_PROC = THIS_PROC
15790   That_Erl = Erl
15800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15810   Resume EXITP

End Sub

Private Sub PricingExpires_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "PricingExpires_MouseMove"

15910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15920   Exit Sub

ERRH:
15930   Select Case ERR.Number
        Case Else
15940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15950   End Select
15960   Resume EXITP

End Sub

Private Sub PricingCode_AfterUpdate()

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "PricingCode_AfterUpdate"

16010   With Me
16020     blnPasswordOnly = False
16030     If blnFromUtilityMenu = True Then
16040       If .cmdOK.Enabled = False Then .cmdOK.Enabled = True
16050     End If
16060   End With

EXITP:
16070   Exit Sub

ERRH:
16080   THAT_PROC = THIS_PROC
16090   That_Erl = Erl
16100   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16110   Resume EXITP

End Sub

Private Sub PricingCode_KeyDown(KeyCode As Integer, Shift As Integer)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "PricingCode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16220   intShiftDown = (Shift And acShiftMask) > 0
16230   intAltDown = (Shift And acAltMask) > 0
16240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16270       With Me
16280         intRetVal = 0
16290         If .cmdOK.Enabled = True Then
16300           .cmdOK.SetFocus
16310         ElseIf .cmdCancel.Enabled = True And .cmdCancel.Visible = True Then
16320           .cmdCancel.SetFocus
16330         Else
16340           .cmdQuit.SetFocus
16350         End If
16360       End With
16370     End Select
16380   End If

        ' ** Shift keys.
16390   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16400     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16410       With Me
16420         intRetVal = 0
16430         .PricingExpires.SetFocus
16440       End With
16450     End Select
16460   End If

EXITP:
16470   KeyCode = intRetVal
16480   Exit Sub

ERRH:
16490   intRetVal = 0
16500   Select Case ERR.Number
        Case Else
16510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16520   End Select
16530   Resume EXITP

End Sub

Private Sub PricingCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "PricingCode_MouseMove"

16610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16620   Exit Sub

ERRH:
16630   Select Case ERR.Number
        Case Else
16640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16650   End Select
16660   Resume EXITP

End Sub
