VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Main"

'VGC 10/11/2017: CHANGES!
' ** TimerInverval: 100 (100 set by GoToReport)

' ** Shortcut Alt keys responsive from this form:
' **   Accounts:         A {cmdAccount}
' **   About:            B {cmdAbout}
' **   Close:            C {cmdClose}
' **   Load Timer:       D {chkLoadTimer}
' **   Site Map:         M {cmdSiteMap}
' **   Posting/Journal:  P {cmdPost}
' **   Reports:          R {cmdReport}
' **   Assets:           S {cmdAsset}
' **   Utilities:        U {cmdUtility}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

Private blnAbout_Focus As Boolean, blnAbout_MouseDown As Boolean
Private blnSiteMap_Focus As Boolean, blnSiteMap_MouseDown As Boolean

Private lngAbout_Left As Long, lngSiteMap_Left As Long
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long, blnHourglass As Boolean
Private blnJustClose As Boolean, blnFormOpening As Boolean, blnIsOpen As Boolean
Private strDocName As String, strLastFocus As String, lngTpp As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim msgResponse As VbMsgBoxResult
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           If gstrReturningForm = vbNullString Then
160             blnIsOpen = True
170           End If

180           gblnSignal = False

190           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

200           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

210           blnJustClose = False: blnFormOpening = False
220           blnSiteMap_Focus = False: blnSiteMap_MouseDown = False
230           blnAbout_Focus = False: blnAbout_MouseDown = False
240           blnGTR_Emblem = False: lngGTR_Stat = 0&

250           lngAbout_Left = .cmdAbout.Left
260           lngSiteMap_Left = .cmdSiteMap.Left

270           TAReports_SetZero False  ' ** Module Function: modReportFunctions.
280           glngInstance = -1&  ' ** Don't run conversion loop.
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
290           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
300           DoEvents

310           If Right(gstrTrustDataLocation, 1) <> LNK_SEP Then gstrTrustDataLocation = gstrTrustDataLocation & LNK_SEP

              ' ** Make sure MouseHook.dll is present and accounted for.
320           blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.
              ' ** Default to mouse wheel on.
330           If blnRetVal = True Then
340             MouseWheelON  ' ** Module Function: modMouseWheel.
350           End If
360           DoEvents

370           PathCheck True  ' ** Module Function: modXAdminFuncs.
380           DoEvents
390           JrnlPrefCheck  ' ** Module Function: modXAdminFuncs.
400           DoEvents

410           Select Case .chkLoadTimer_vis
              Case True
420             .chkLoadTimer.Visible = True
430             .chkLoadTimer_lbl1.Visible = True
440             .chkLoadTimer_lbl2.Visible = True
450             .chkLoadTimer_lbl3.Visible = True
460             .chkLoadTimer_lbl4.Visible = True
470             .chkLoadTimer_lbl5.Visible = True
480           Case False
490             .chkLoadTimer.Visible = False
500             .chkLoadTimer_lbl1.Visible = False
510             .chkLoadTimer_lbl2.Visible = False
520             .chkLoadTimer_lbl3.Visible = False
530             .chkLoadTimer_lbl4.Visible = False
540             .chkLoadTimer_lbl5.Visible = False
550           End Select
560           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
570           DoEvents

580           If IsLoaded("frmMenu_Utility", acForm) = True Then  ' ** Module Function: modFileUtilities.
590   On Error Resume Next
600             DoCmd.Close acForm, "frmMenu_Utility", acSaveNo
610   On Error GoTo ERRH
620           End If

630           If CoOptions_Read = False Then  ' ** Module Function: modStartupFuncs.
640             msgResponse = MsgBox("Unable to read company information and options!", vbCritical + vbOKOnly, "Error")
650           Else

                ' ** Build the company info for reports.
660             gstrCo_InfoBlock = vbNullString
670             If gstrCo_Name <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & gstrCo_Name & vbCrLf
680             If gstrCo_Address1 <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & gstrCo_Address1 & vbCrLf
690             If gstrCo_Address2 <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & gstrCo_Address2 & vbCrLf
700             If gstrCo_Country = vbNullString Then
710               If gstrCo_City <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & gstrCo_City & ", "
720               If gstrCo_State <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & gstrCo_State & "  "
730               If gstrCo_Zip <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & FormatZip9(gstrCo_Zip) & vbCrLf
740             Else
750               If gstrCo_City <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & gstrCo_City & vbCrLf
760               If gstrCo_Country <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & gstrCo_Country
770               If gstrCo_PostalCode <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & "  " & gstrCo_PostalCode
780             End If
790             If gstrCo_Phone <> vbNullString Then gstrCo_InfoBlock = gstrCo_InfoBlock & FormatPhoneNum(gstrCo_Phone)

                ' ** Start with windows not maximized.
800             DoCmd.Restore

810           End If

820         Else
830           Cancel = -1
840         End If
850       Else
860         Cancel = -1
870       End If
880     End With

890     If Cancel = -1 Then
900       Beep
910       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
920     End If

930     DoCmd.Hourglass False

EXITP:
940     Exit Sub

ERRH:
950     DoCmd.Hourglass False
960     Select Case ERR.Number
        Case Else
970       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
980     End Select
990     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer, Cancel As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyEscape
1070        intRetVal = 0
1080        cmdClose_Click  ' ** Procedure: Below.
1090      End Select
1100    End If

        ' ** Alt keys.
1110    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1120      Select Case intRetVal
          Case vbKeyX
1130        intRetVal = 0
1140        cmdClose_Click  ' ** Procedure: Below.
1150      End Select
1160    End If

        ' ** Ctrl-Shift keys.
1170    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1180      Select Case intRetVal
          Case vbKeyF
1190        With Me
1200          intRetVal = 0
1210          .FocusHolder.SetFocus
1220        End With
1230      Case vbKeyT
1240        With Me
1250          intRetVal = 0
1260          If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
1270            FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
1280          ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
1290            Select Case .chkLoadTimer.Visible
                Case True
1300              gblnTimer = False
1310              .FocusHolder.SetFocus
1320              DoEvents
1330              .chkLoadTimer.Visible = False
1340              .chkLoadTimer_lbl1.Visible = False
1350              .chkLoadTimer_lbl2.Visible = False
1360              .chkLoadTimer_lbl3.Visible = False
1370              .chkLoadTimer_lbl4.Visible = False
1380              .chkLoadTimer_lbl5.Visible = False
1390            Case False
1400              .chkLoadTimer.Visible = True
1410              .chkLoadTimer_lbl1.Visible = True
1420              .chkLoadTimer_lbl2.Visible = True
1430              .chkLoadTimer_lbl3.Visible = True
1440              .chkLoadTimer_lbl4.Visible = True
1450              .chkLoadTimer_lbl5.Visible = True
1460              chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
1470            End Select
1480            .chkLoadTimer_vis = .chkLoadTimer.Visible
1490          End If
1500        End With
1510      Case vbKeyX
1520        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1530          intRetVal = 0
1540          blnJustClose = True
1550          SetOption_Dev  ' ** Module Function: modUtilities.
1560          gblnSignal = True
1570          DoCmd.Close acForm, THIS_NAME
1580          DoCmd.Hourglass False
1590        End If
1600      End Select
1610    End If

EXITP:
1620    KeyCode = intRetVal
1630    Exit Sub

ERRH:
1640    intRetVal = 0
1650    Select Case ERR.Number
        Case Else
1660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1670    End Select
1680    Resume EXITP

End Sub

Private Sub Form_Timer()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long

1710    With Me
1720      .TimerInterval = 0&
1730      If blnGTR_Emblem = True Then
1740        lngGTR_Stat = lngGTR_Stat + 1&
1750        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1760        lngTmp01 = lngTmp01 + 1&
1770        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1780        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1790        If lngTmp01 > 1& Then
1800          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1810        End If
1820        If lngTmp01 = 24& Then
1830          lngTmp01 = 0&
1840          lngTmp02 = lngTmp02 + 1&
1850          If lngTmp02 < 5& Then
1860            .TimerInterval = 50&
1870          Else
1880            blnGTR_Emblem = False
1890          End If
1900        Else
1910          .TimerInterval = 50&
1920        End If
1930      End If  ' ** blnGTR_Emblem.
1940      If gblnSetFocus = True Then
1950        gblnSetFocus = False
1960        lngTmp03 = 0&
1970        strFormName = vbNullString
1980  On Error Resume Next
1990        Do Until strFormName = THIS_NAME
2000          lngTmp03 = lngTmp03 + 1&
2010          DoCmd.SelectObject acForm, THIS_NAME
2020          strFormName = Screen.ActiveForm.Name
2030          If lngTmp03 >= 100& Then Exit Do
2040        Loop
2050  On Error GoTo ERRH
2060        DoEvents
2070      End If
2080      If blnHourglass = True Then
2090        blnHourglass = False
2100        DoCmd.Hourglass False
2110      End If
2120      If blnGoingToReport = False Then
2130        Select Case gstrReturningForm
            Case "frmMenu_Account"
2140          .cmdAccount.SetFocus
2150        Case "frmMenu_Post"
2160          .cmdPost.SetFocus
2170        Case "frmMenu_Report"
2180          .cmdReport.SetFocus
2190        Case "frmMenu_Asset"
2200          .cmdAsset.SetFocus
2210        Case "frmMenu_Utility"
2220          .cmdUtility.SetFocus
2230        Case "frmSiteMap"
2240          .cmdSiteMap.SetFocus
2250          blnHourglass = True
2260          .TimerInterval = 500&
2270        Case "frmAbout"
2280          .cmdAbout.SetFocus
2290          blnHourglass = True
2300          .TimerInterval = 500&
2310        Case Else
2320          .cmdAccount.SetFocus
2330        End Select
2340        gstrReturningForm = vbNullString
2350        DoCmd.Hourglass False
2360      End If
2370      If gblnGoToReport = True Then
2380        If (garr_varGoToReport(GTR_FRM1) <> "frmMenu_Asset" And ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&))) Then
              ' ** Just let the emblem go.
2390        ElseIf (garr_varGoToReport(GTR_FRM1) = "frmMenu_Asset" And ((lngGTR_Stat > 0& And lngGTR_Stat < 60&) Or (lngGTR_Stat > 60&))) Then
              ' ** Just let the emblem go a little longer.
2400        Else
2410          DoCmd.Hourglass True  ' ** Make sure it's still running.
2420          If blnGTR_Emblem = False Then
2430            blnGTR_Emblem = True
2440            lngTmp01 = 0&: lngTmp02 = 0&
2450          End If
2460          DoEvents
2470          Select Case blnGoingToReport
              Case True
2480            Select Case garr_varGoToReport(GTR_FRM1)
                Case "frmMenu_Account"
2490              .TimerInterval = 50&
2500              cmdAccount_Click  ' ** Procedure: Below.
2510            Case "frmMenu_Post"
2520              .TimerInterval = 50&
2530              cmdPost_Click  ' ** Procedure: Below.
2540            Case "frmMenu_Report"
2550              .TimerInterval = 50&
2560              cmdReport_Click  ' ** Procedure: Below.
2570            Case "frmMenu_Asset"
2580              .TimerInterval = 50&
2590              cmdAsset_Click  ' ** Procedure: Below.
2600            Case "frmMenu_Utility"
2610              .TimerInterval = 50&
2620              cmdUtility_Click  ' ** Procedure: Below.
2630            Case "frmMenu_Title"
2640              .TimerInterval = 50&
2650              cmdClose_Click  ' ** Procedure: Below.
2660            Case "frmMenu_Main"
2670              .TimerInterval = 50&
2680              cmdSiteMap_Click  ' ** Procedure: Below.
2690            Case Else
                  ' ** Nothing else.
2700              blnGoingToReport = False
2710              DoCmd.Hourglass False
2720              gblnGoToReport = False
2730              blnGTR_Emblem = False
2740              GTREmblem_Off  ' ** Procedure: Below.
2750              Beep
2760            End Select
2770          Case False
2780            blnGoingToReport = True
2790            Select Case garr_varGoToReport(GTR_FRM1)
                Case "frmMenu_Account"
2800              .cmdAccount.SetFocus
2810              .GoToReport_arw_menu_acct_img.Visible = True
2820              DoEvents
2830              .TimerInterval = 50&
2840            Case "frmMenu_Post"
2850              .cmdPost.SetFocus
2860              .GoToReport_arw_menu_post_img.Visible = True
2870              DoEvents
2880              .TimerInterval = 50&
2890            Case "frmMenu_Asset"
2900              .cmdAsset.SetFocus
2910              .GoToReport_arw_menu_asset_img.Visible = True
2920              DoEvents
2930              .TimerInterval = 50&
2940            Case "frmMenu_Utility"
2950              .cmdUtility.SetFocus
2960              .GoToReport_arw_menu_util_img.Visible = True
2970              DoEvents
2980              .TimerInterval = 50&
2990            Case "frmMenu_Title"
3000              .cmdClose.SetFocus
3010              DoEvents
3020              .TimerInterval = 50&
3030            Case "frmMenu_Main"
3040              .cmdSiteMap.SetFocus
3050              .GoToReport_arw_site_img.Visible = True
3060              DoEvents
3070              .TimerInterval = 50&
3080            Case Else
                  ' ** Nothing else.
3090              blnGoingToReport = False
3100              DoCmd.Hourglass False
3110              gblnGoToReport = False
3120              blnGTR_Emblem = False
3130              GTREmblem_Off  ' ** Procedure: Below.
3140              Beep
3150            End Select
3160          End Select
3170        End If
3180      End If  ' ** gblnGoToReport.
3190    End With

EXITP:
3200    Exit Sub

ERRH:
3210    DoCmd.Hourglass False
3220    blnGoingToReport = False
3230    gblnGoToReport = False
3240    Select Case ERR.Number
        Case Else
3250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3260    End Select
3270    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3310    With Me

3320      If gblnSignal = False Then  ' ** Borrowing this variable.
3330        DoCmd.Hourglass True
3340        DoEvents
3350      End If

          ' ** Make sure all reports are closed.
3360      Do While Reports.Count > 0
3370        DoCmd.Close acReport, Reports(0).Name
3380      Loop

3390      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3400      blnGoingToReport = False

3410    End With

EXITP:
3420    Exit Sub

ERRH:
3430    DoCmd.Hourglass False
3440    Select Case ERR.Number
        Case 2450  ' ** Microsoft Access can't find the form 'frmMenu_Title' referred to in a macro expression or Visual Basic code.
          ' ** Sometimes during Quit from the license screen, Access gets confused.
3450    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
3460    Case Else
3470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3480    End Select
3490    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3510    With Me
3520      If .cmdAbout_raised_focus_dots_img.Visible = True Or .cmdAbout_raised_focus_img.Visible = True Then
3530        Select Case blnAbout_Focus
            Case True
3540          .cmdAbout_raised_semifocus_dots_img.Visible = True
3550          .cmdAbout_raised_img.Visible = False
3560        Case False
3570          .cmdAbout_raised_img.Visible = True
3580          .cmdAbout_raised_semifocus_dots_img.Visible = False
3590        End Select
3600        .cmdAbout_raised_focus_img.Visible = False
3610        .cmdAbout_raised_focus_dots_img.Visible = False
3620        .cmdAbout_sunken_focus_dots_img.Visible = False
3630        .cmdAbout_raised_img_dis.Visible = False
3640      End If
3650      If .cmdSiteMap_raised_focus_dots_img.Visible = True Or .cmdSiteMap_raised_focus_img.Visible = True Then
3660        Select Case blnSiteMap_Focus
            Case True
3670          .cmdSiteMap_raised_semifocus_dots_img.Visible = True
3680          .cmdSiteMap_raised_img.Visible = False
3690        Case False
3700          .cmdSiteMap_raised_img.Visible = True
3710          .cmdSiteMap_raised_semifocus_dots_img.Visible = False
3720        End Select
3730        .cmdSiteMap_raised_focus_img.Visible = False
3740        .cmdSiteMap_raised_focus_dots_img.Visible = False
3750        .cmdSiteMap_sunken_focus_dots_img.Visible = False
3760        .cmdSiteMap_raised_img_dis.Visible = False
3770      End If
3780      .cmdAccount_box01.Visible = False
3790      .cmdAccount_box02.Visible = False
3800      .cmdPost_box01.Visible = False
3810      .cmdPost_box02.Visible = False
3820      .cmdReport_box01.Visible = False
3830      .cmdReport_box02.Visible = False
3840      .cmdAsset_box01.Visible = False
3850      .cmdAsset_box02.Visible = False
3860      .cmdUtility_box01.Visible = False
3870      .cmdUtility_box02.Visible = False
3880      If .cmdClose.Caption <> "&Close" Then
3890        .cmdClose.Caption = "&Close"
3900      End If
3910      If strLastFocus <> "cmdClose" And strLastFocus <> "cmdSiteMap" And strLastFocus <> "cmdAbout" Then
3920        If strLastFocus <> vbNullString Then
3930          .Controls(strLastFocus & "_box01").Visible = True
3940          .Controls(strLastFocus & "_box02").Visible = True
3950        End If
3960      End If
3970    End With

EXITP:
3980    Exit Sub

ERRH:
3990    Select Case ERR.Number
        Case Else
4000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4010    End Select
4020    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

4110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub Header_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_box_MouseMove"

4210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

4310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub Header_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl2_MouseMove"

4410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub cmdMain_box2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMain_box2_MouseMove"

4510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub FocusHolder_GotFocus()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_GotFocus"

4610    With Me
4620      .cmdAccount_box01.Visible = False
4630      .cmdAccount_box02.Visible = False
4640      .cmdPost_box01.Visible = False
4650      .cmdPost_box02.Visible = False
4660      .cmdReport_box01.Visible = False
4670      .cmdReport_box02.Visible = False
4680      .cmdAsset_box01.Visible = False
4690      .cmdAsset_box02.Visible = False
4700      .cmdUtility_box01.Visible = False
4710      .cmdUtility_box02.Visible = False
4720      DoEvents
4730      .Repaint
4740    End With

EXITP:
4750    Exit Sub

ERRH:
4760    Select Case ERR.Number
        Case Else
4770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4780    End Select
4790    Resume EXITP

End Sub

Private Sub cmdClose_GotFocus()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_GotFocus"

4810    With Me
4820      strLastFocus = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
4830    End With

EXITP:
4840    Exit Sub

ERRH:
4850    Select Case ERR.Number
        Case Else
4860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4870    End Select
4880    Resume EXITP

End Sub

Private Sub cmdClose_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4910    With Me
4920      If blnFormOpening = False Then

4930        DoCmd.Hourglass True
4940        DoEvents

4950        .TimerInterval = 0&

4960        If strCallingForm = vbNullString Then strCallingForm = "frmMenu_Title"
4970        If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
              ' ** Quickly clicking back-and-forth sometimes leaves nothing!
4980          gblnSetFocus = True
4990          gstrReturningForm = THIS_NAME
5000          Select Case strCallingForm
              Case "frmMenu_Title"
5010            strDocName = "frmSiteMap"
5020            If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
5030              DoCmd.Close acForm, strDocName
5040              DoEvents
5050            End If
5060            DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
                ' ** Initializing message now has its own label!
5070            With Forms(strCallingForm)
5080              .InitMsg_lbl1.Visible = False
5090              DoEvents
5100  On Error Resume Next
                  ' ** One back too far, to frmMenu_Title, then quick 'Main Menu' back to here.
5110              .cmdMenu.Visible = True
5120              .cmdMenu_box1.Visible = True
5130              .cmdMenu_box2.Visible = True
5140              .cmdQuit.Visible = True
5150              .cmdQuit_box1.Visible = True
5160              .cmdQuit_box2.Visible = True
5170  On Error GoTo ERRH
5180            End With
5190          End Select
              'Else
              '  Cancel = -1
5200        End If

5210        blnJustClose = False
5220        DoCmd.Close acForm, THIS_NAME

5230        DoCmd.Hourglass False

5240      End If
5250    End With

EXITP:
5260    Exit Sub

ERRH:
5270    blnFormOpening = False
5280    DoCmd.Hourglass False
5290    Select Case ERR.Number
        Case 2450  ' ** Microsoft Access can't find the form 'frmMenu_Title' referred to in a macro expression or Visual Basic code.
          ' ** Sometimes during Quit from the license screen, Access gets confused.
5300    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
5310    Case Else
5320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5330    End Select
5340    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
5470        With Me
5480          intRetVal = 0
5490          .cmdAccount.SetFocus
5500        End With
5510      End Select
5520    End If

        ' ** Shift keys.
5530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5540      Select Case intRetVal
          Case vbKeyTab
5550        With Me
5560          intRetVal = 0
5570          .cmdAbout.SetFocus
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

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_MouseMove"

5710    With Me
5720      .cmdClose.Caption = "< Ba&ck"
5730    End With

EXITP:
5740    Exit Sub

ERRH:
5750    Select Case ERR.Number
        Case Else
5760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5770    End Select
5780    Resume EXITP

End Sub

Private Sub cmdAccount_GotFocus()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccount_GotFocus"

5810    With Me
5820      strLastFocus = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
5830      Select Case blnIsOpen
          Case True
5840        blnIsOpen = False
5850      Case False
5860        .cmdAccount_box01.Visible = True
5870        .cmdAccount_box02.Visible = True
5880        .cmdPost_box01.Visible = False
5890        .cmdPost_box02.Visible = False
5900        .cmdReport_box01.Visible = False
5910        .cmdReport_box02.Visible = False
5920        .cmdAsset_box01.Visible = False
5930        .cmdAsset_box02.Visible = False
5940        .cmdUtility_box01.Visible = False
5950        .cmdUtility_box02.Visible = False
5960      End Select
5970    End With

EXITP:
5980    Exit Sub

ERRH:
5990    Select Case ERR.Number
        Case Else
6000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6010    End Select
6020    Resume EXITP

End Sub

Private Sub cmdAccount_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccount_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

6110    With Me
6120      If blnFormOpening = False Then

6130        strDocName = "frmMenu_Account"

6140        If .chkLoadTimer.Visible = True Then
6150          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
6160          If gblnTimer = True Then
                'WE CAN'T USE THIS HERE, BECAUSE AS SOON AS IT'S TURNED ON, BELOW,
                'THE TIMER LOOP IN THIS FORM HERE WILL TURN IT OFF!!
                'gblnSetFocus = True
6170            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
6180          End If
6190        End If
6200        DoEvents

6210        DoCmd.Hourglass True
6220        DoEvents

6230        .TimerInterval = 0&

6240        blnFormOpening = True

6250        Set dbs = CurrentDb
            ' ** Account, just accountno.
6260        Set qdf = dbs.QueryDefs("qryAccountProfile_01b")
6270        Set rst = qdf.OpenRecordset
6280        If rst.BOF = True And rst.EOF = True Then
              ' ** There are no accounts,
6290          rst.Close
6300          Set rst = Nothing
6310          Set qdf = Nothing
6320          dbs.Close
6330          Set dbs = Nothing
6340          DoCmd.Hourglass False
6350          MsgBox "There are no accounts to view.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
6360          If gblnGoToReport = True Then
6370            blnGoingToReport = False
6380            GTREmblem_Off  ' ** Procedure: Below.
6390            gblnGoToReport = False
6400          End If
6410          If gblnTimer = True Then
6420            gblnTimer = False
6430            DoCmd.Close acForm, "frmLoadTimer"
6440            .chkLoadTimer = False
6450            chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
6460          End If
6470        Else
6480          rst.Close
6490          Set rst = Nothing
6500          Set qdf = Nothing
              ' ** Append qryAccountProfile_01c (qryAccountProfile_01b (Account, just accountno), with
              ' ** qryAccountProfile_01a (tblForm_Graphics, just frmAccountProfile_Sub); Cartesian) to tmpForm_Graphics_Account.
6510          Set qdf = dbs.QueryDefs("qryAccountProfile_01d")
6520          qdf.Execute
6530          Set qdf = Nothing
6540          dbs.Close
6550          Set dbs = Nothing
6560          blnJustClose = True
6570          gblnSetFocus = True
6580          gstrReturningForm = vbNullString
6590          gstrActNo = vbNullString
6600          gstrAccountNo = vbNullString
6610          gstrAccountName = vbNullString
6620          DoCmd.Close acForm, THIS_NAME
6630          DoCmd.OpenForm strDocName, , , , , , THIS_NAME
6640          If gblnGoToReport = True Then
6650            blnGoingToReport = False
6660            GTREmblem_Off  ' ** Procedure: Below.
6670          End If
6680        End If

6690      End If
6700    End With

EXITP:
6710    Set rst = Nothing
6720    Set qdf = Nothing
6730    Set dbs = Nothing
6740    Exit Sub

ERRH:
6750    blnFormOpening = False
6760    DoCmd.Hourglass False
6770    Select Case ERR.Number
        Case Else
6780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6790    End Select
6800    Resume EXITP

End Sub

Private Sub cmdAccount_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccount_KeyDown"

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
6990          .cmdPost.SetFocus
7000        End With
7010      End Select
7020    End If

        ' ** Shift keys.
7030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7040      Select Case intRetVal
          Case vbKeyTab
7050        With Me
7060          intRetVal = 0
7070          .cmdClose.SetFocus
7080        End With
7090      End Select
7100    End If

EXITP:
7110    KeyCode = intRetVal
7120    Exit Sub

ERRH:
7130    intRetVal = 0
7140    Select Case ERR.Number
        Case Else
7150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7160    End Select
7170    Resume EXITP

End Sub

Private Sub cmdAccount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccount_MouseMove"

7210    With Me
7220      If .cmdAccount_box01.Visible = False Then
7230        .cmdAccount_box01.Visible = True
7240        .cmdAccount_box02.Visible = True
7250      End If
7260      If .cmdPost_box01.Visible = True Then
7270        .cmdPost_box01.Visible = False
7280        .cmdPost_box02.Visible = False
7290      End If
7300      If .cmdReport_box01.Visible = True Then
7310        .cmdReport_box01.Visible = False
7320        .cmdReport_box02.Visible = False
7330      End If
7340      If .cmdAsset_box01.Visible = True Then
7350        .cmdAsset_box01.Visible = False
7360        .cmdAsset_box02.Visible = False
7370      End If
7380      If .cmdUtility_box01.Visible = True Then
7390        .cmdUtility_box01.Visible = False
7400        .cmdUtility_box02.Visible = False
7410      End If
7420    End With

EXITP:
7430    Exit Sub

ERRH:
7440    Select Case ERR.Number
        Case Else
7450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7460    End Select
7470    Resume EXITP

End Sub

Private Sub cmdAccount_LostFocus()

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccount_LostFocus"

7510    With Me
7520      .cmdAccount_box01.Visible = False
7530      .cmdAccount_box02.Visible = False
7540    End With

EXITP:
7550    Exit Sub

ERRH:
7560    Select Case ERR.Number
        Case Else
7570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7580    End Select
7590    Resume EXITP

End Sub

Private Sub cmdPost_GotFocus()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPost_GotFocus"

7610    With Me
7620      strLastFocus = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
7630      .cmdPost_box01.Visible = True
7640      .cmdPost_box02.Visible = True
7650      .cmdAccount_box01.Visible = False
7660      .cmdAccount_box02.Visible = False
7670      .cmdReport_box01.Visible = False
7680      .cmdReport_box02.Visible = False
7690      .cmdAsset_box01.Visible = False
7700      .cmdAsset_box02.Visible = False
7710      .cmdUtility_box01.Visible = False
7720      .cmdUtility_box02.Visible = False
7730    End With

EXITP:
7740    Exit Sub

ERRH:
7750    Select Case ERR.Number
        Case Else
7760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7770    End Select
7780    Resume EXITP

End Sub

Private Sub cmdPost_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPost_Click"

        Dim usr As DAO.User, grp As DAO.Group
        Dim blnFound As Boolean

7810    With Me
7820      If blnFormOpening = False Then

7830        strDocName = "frmMenu_Post"

7840        If .chkLoadTimer.Visible = True Then
7850          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
7860          If gblnTimer = True Then
                'gblnSetFocus = True
7870            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
7880          End If
7890        End If
7900        DoEvents

7910        DoCmd.Hourglass True
7920        DoEvents

7930        .TimerInterval = 0&

7940        blnFound = False

7950        For Each grp In DBEngine.Workspaces(0).Groups
7960          If grp.Name = "ViewOnly" Then
7970            For Each usr In grp.Users
7980              If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
7990                blnFound = True
8000                Exit For
8010              End If
8020            Next
8030          End If
8040        Next

8050        If blnFound = True Then
8060          DoCmd.Hourglass False
8070          MsgBox gstrNoPermission, vbInformation + vbOKOnly, "Access Denied"
8080          If gblnGoToReport = True Then
8090            blnGoingToReport = False
8100            GTREmblem_Off  ' ** Procedure: Below.
8110            gblnGoToReport = False
8120          End If
8130        Else
8140          blnJustClose = True: blnFormOpening = True
8150          gstrReturningForm = vbNullString
8160          gblnSetFocus = True
8170          DoCmd.Close acForm, THIS_NAME
8180          DoCmd.OpenForm strDocName, , , , , , THIS_NAME
8190          If gblnGoToReport = False Then
8200            blnGoingToReport = False
8210            GTREmblem_Off  ' ** Procedure: Below.
8220          End If
8230        End If

8240      End If
8250    End With

EXITP:
8260    Set usr = Nothing
8270    Set grp = Nothing
8280    Exit Sub

ERRH:
8290    blnFormOpening = False
8300    DoCmd.Hourglass False
8310    Select Case ERR.Number
        Case Else
8320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8330    End Select
8340    Resume EXITP

End Sub

Private Sub cmdPost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPost_MouseMove"

8410    With Me
8420      If .cmdPost_box01.Visible = False Then
8430        .cmdPost_box01.Visible = True
8440        .cmdPost_box02.Visible = True
8450      End If
8460      If .cmdAccount_box01.Visible = True Then
8470        .cmdAccount_box01.Visible = False
8480        .cmdAccount_box02.Visible = False
8490      End If
8500      If .cmdReport_box01.Visible = True Then
8510        .cmdReport_box01.Visible = False
8520        .cmdReport_box02.Visible = False
8530      End If
8540      If .cmdAsset_box01.Visible = True Then
8550        .cmdAsset_box01.Visible = False
8560        .cmdAsset_box02.Visible = False
8570      End If
8580      If .cmdUtility_box01.Visible = True Then
8590        .cmdUtility_box01.Visible = False
8600        .cmdUtility_box02.Visible = False
8610      End If
8620    End With

EXITP:
8630    Exit Sub

ERRH:
8640    Select Case ERR.Number
        Case Else
8650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8660    End Select
8670    Resume EXITP

End Sub

Private Sub cmdPost_LostFocus()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPost_LostFocus"

8710    With Me
8720      .cmdPost_box01.Visible = False
8730      .cmdPost_box02.Visible = False
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

Private Sub cmdReport_GotFocus()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReport_GotFocus"

8810    With Me
8820      strLastFocus = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
8830      .cmdReport_box01.Visible = True
8840      .cmdReport_box02.Visible = True
8850      .cmdAccount_box01.Visible = False
8860      .cmdAccount_box02.Visible = False
8870      .cmdPost_box01.Visible = False
8880      .cmdPost_box02.Visible = False
8890      .cmdAsset_box01.Visible = False
8900      .cmdAsset_box02.Visible = False
8910      .cmdUtility_box01.Visible = False
8920      .cmdUtility_box02.Visible = False
8930    End With

EXITP:
8940    Exit Sub

ERRH:
8950    Select Case ERR.Number
        Case Else
8960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8970    End Select
8980    Resume EXITP

End Sub

Private Sub cmdReport_Click()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReport_Click"

9010    With Me
9020      If blnFormOpening = False Then

9030        strDocName = "frmMenu_Report"

9040        If .chkLoadTimer.Visible = True Then
9050          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
9060          If gblnTimer = True Then
                'gblnSetFocus = True
9070            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
9080          End If
9090        End If
9100        DoEvents

9110        DoCmd.Hourglass True
9120        DoEvents

9130        .TimerInterval = 0&

9140        blnJustClose = True: blnFormOpening = True
9150        gblnSetFocus = True
9160        gstrReturningForm = vbNullString
9170        DoCmd.Close acForm, THIS_NAME
9180        DoCmd.OpenForm strDocName, , , , , , THIS_NAME

9190      End If
9200    End With

EXITP:
9210    Exit Sub

ERRH:
9220    blnFormOpening = False
9230    DoCmd.Hourglass False
9240    Select Case ERR.Number
        Case Else
9250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9260    End Select
9270    Resume EXITP

End Sub

Private Sub cmdReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReport_MouseMove"

9310    With Me
9320      If .cmdReport_box01.Visible = False Then
9330        .cmdReport_box01.Visible = True
9340        .cmdReport_box02.Visible = True
9350      End If
9360      If .cmdAccount_box01.Visible = True Then
9370        .cmdAccount_box01.Visible = False
9380        .cmdAccount_box02.Visible = False
9390      End If
9400      If .cmdPost_box01.Visible = True Then
9410        .cmdPost_box01.Visible = False
9420        .cmdPost_box02.Visible = False
9430      End If
9440      If .cmdAsset_box01.Visible = True Then
9450        .cmdAsset_box01.Visible = False
9460        .cmdAsset_box02.Visible = False
9470      End If
9480      If .cmdUtility_box01.Visible = True Then
9490        .cmdUtility_box01.Visible = False
9500        .cmdUtility_box02.Visible = False
9510      End If
9520    End With

EXITP:
9530    Exit Sub

ERRH:
9540    Select Case ERR.Number
        Case Else
9550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9560    End Select
9570    Resume EXITP

End Sub

Private Sub cmdReport_LostFocus()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReport_LostFocus"

9610    With Me
9620      .cmdReport_box01.Visible = False
9630      .cmdReport_box02.Visible = False
9640    End With

EXITP:
9650    Exit Sub

ERRH:
9660    Select Case ERR.Number
        Case Else
9670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9680    End Select
9690    Resume EXITP

End Sub

Private Sub cmdAsset_GotFocus()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAsset_GotFocus"

9710    With Me
9720      strLastFocus = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
9730      .cmdAsset_box01.Visible = True
9740      .cmdAsset_box02.Visible = True
9750      .cmdReport_box01.Visible = False
9760      .cmdReport_box02.Visible = False
9770      .cmdAccount_box01.Visible = False
9780      .cmdAccount_box02.Visible = False
9790      .cmdPost_box01.Visible = False
9800      .cmdPost_box02.Visible = False
9810      .cmdUtility_box01.Visible = False
9820      .cmdUtility_box02.Visible = False
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

Private Sub cmdAsset_Click()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAsset_Click"

9910    With Me
9920      If blnFormOpening = False Then

9930        strDocName = "frmMenu_Asset"

9940        If .chkLoadTimer.Visible = True Then
9950          chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
9960          If gblnTimer = True Then
                'gblnSetFocus = True
9970            DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
9980          End If
9990        End If
10000       DoEvents

10010       DoCmd.Hourglass True
10020       DoEvents

10030       .TimerInterval = 0&

10040       blnJustClose = True: blnFormOpening = True
10050       gblnSetFocus = True
10060       gstrReturningForm = vbNullString
10070       DoCmd.Close acForm, THIS_NAME
10080       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
10090       If gblnGoToReport = True Then
10100         blnGoingToReport = False
10110         GTREmblem_Off  ' ** Procedure: Below.
10120       End If

10130     End If
10140   End With

EXITP:
10150   Exit Sub

ERRH:
10160   blnFormOpening = False
10170   DoCmd.Hourglass False
10180   Select Case ERR.Number
        Case Else
10190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10200   End Select
10210   Resume EXITP

End Sub

Private Sub cmdAsset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAsset_MouseMove"

10310   With Me
10320     If .cmdAsset_box01.Visible = False Then
10330       .cmdAsset_box01.Visible = True
10340       .cmdAsset_box02.Visible = True
10350     End If
10360     If .cmdAccount_box01.Visible = True Then
10370       .cmdAccount_box01.Visible = False
10380       .cmdAccount_box02.Visible = False
10390     End If
10400     If .cmdPost_box01.Visible = True Then
10410       .cmdPost_box01.Visible = False
10420       .cmdPost_box02.Visible = False
10430     End If
10440     If .cmdReport_box01.Visible = True Then
10450       .cmdReport_box01.Visible = False
10460       .cmdReport_box02.Visible = False
10470     End If
10480     If .cmdUtility_box01.Visible = True Then
10490       .cmdUtility_box01.Visible = False
10500       .cmdUtility_box02.Visible = False
10510     End If
10520   End With

EXITP:
10530   Exit Sub

ERRH:
10540   Select Case ERR.Number
        Case Else
10550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10560   End Select
10570   Resume EXITP

End Sub

Private Sub cmdAsset_LostFocus()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAsset_LostFocus"

10610   With Me
10620     .cmdAsset_box01.Visible = False
10630     .cmdAsset_box02.Visible = False
10640   End With

EXITP:
10650   Exit Sub

ERRH:
10660   Select Case ERR.Number
        Case Else
10670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10680   End Select
10690   Resume EXITP

End Sub

Private Sub cmdUtility_GotFocus()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUtility_GotFocus"

10710   With Me
10720     strLastFocus = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
10730     .cmdUtility_box01.Visible = True
10740     .cmdUtility_box02.Visible = True
10750     .cmdAsset_box01.Visible = False
10760     .cmdAsset_box02.Visible = False
10770     .cmdReport_box01.Visible = False
10780     .cmdReport_box02.Visible = False
10790     .cmdAccount_box01.Visible = False
10800     .cmdAccount_box02.Visible = False
10810     .cmdPost_box01.Visible = False
10820     .cmdPost_box02.Visible = False
10830   End With

EXITP:
10840   Exit Sub

ERRH:
10850   Select Case ERR.Number
        Case Else
10860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10870   End Select
10880   Resume EXITP

End Sub

Private Sub cmdUtility_Click()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUtility_Click"

10910   With Me
10920     If blnFormOpening = False Then

10930       strDocName = "frmMenu_Utility"

10940       If .chkLoadTimer.Visible = True Then
10950         chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
10960         If gblnTimer = True Then
                'gblnSetFocus = True
10970           DoCmd.OpenForm "frmLoadTimer", , , , , , THIS_NAME & "~" & strDocName & "~True"
10980         End If
10990       End If
11000       DoEvents

11010       DoCmd.Hourglass True
11020       DoEvents

11030       .TimerInterval = 0&

11040       blnJustClose = True: blnFormOpening = True
11050       gblnSetFocus = True
11060       gstrReturningForm = vbNullString
11070       DoCmd.Close acForm, THIS_NAME
11080       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
11090       If gblnGoToReport = True Then
11100         blnGoingToReport = False
11110         GTREmblem_Off  ' ** Procedure: Below.
11120       End If

11130     End If
11140   End With

EXITP:
11150   Exit Sub

ERRH:
11160   blnFormOpening = False
11170   DoCmd.Hourglass False
11180   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
11190   Case Else
11200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11210   End Select
11220   Resume EXITP

End Sub

Private Sub cmdUtility_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUtility_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11320   intShiftDown = (Shift And acShiftMask) > 0
11330   intAltDown = (Shift And acAltMask) > 0
11340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11360     Select Case intRetVal
          Case vbKeyTab
11370       With Me
11380         intRetVal = 0
11390         .cmdSiteMap.SetFocus
11400       End With
11410     End Select
11420   End If

        ' ** Shift keys.
11430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11440     Select Case intRetVal
          Case vbKeyTab
11450       With Me
11460         intRetVal = 0
11470         .cmdAsset.SetFocus
11480       End With
11490     End Select
11500   End If

EXITP:
11510   KeyCode = intRetVal
11520   Exit Sub

ERRH:
11530   intRetVal = 0
11540   Select Case ERR.Number
        Case Else
11550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11560   End Select
11570   Resume EXITP

End Sub

Private Sub cmdUtility_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUtility_MouseMove"

11610   With Me
11620     If .cmdUtility_box01.Visible = False Then
11630       .cmdUtility_box01.Visible = True
11640       .cmdUtility_box02.Visible = True
11650     End If
11660     If .cmdAccount_box01.Visible = True Then
11670       .cmdAccount_box01.Visible = False
11680       .cmdAccount_box02.Visible = False
11690     End If
11700     If .cmdPost_box01.Visible = True Then
11710       .cmdPost_box01.Visible = False
11720       .cmdPost_box02.Visible = False
11730     End If
11740     If .cmdReport_box01.Visible = True Then
11750       .cmdReport_box01.Visible = False
11760       .cmdReport_box02.Visible = False
11770     End If
11780     If .cmdAsset_box01.Visible = True Then
11790       .cmdAsset_box01.Visible = False
11800       .cmdAsset_box02.Visible = False
11810     End If
11820   End With

EXITP:
11830   Exit Sub

ERRH:
11840   Select Case ERR.Number
        Case Else
11850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11860   End Select
11870   Resume EXITP

End Sub

Private Sub cmdUtility_LostFocus()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUtility_LostFocus"

11910   With Me
11920     .cmdUtility_box01.Visible = False
11930     .cmdUtility_box02.Visible = False
11940   End With

EXITP:
11950   Exit Sub

ERRH:
11960   Select Case ERR.Number
        Case Else
11970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11980   End Select
11990   Resume EXITP

End Sub

Private Sub cmdSiteMap_GotFocus()

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSiteMap_GotFocus"

12010   With Me
12020     strLastFocus = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
12030     blnSiteMap_Focus = True
12040     .cmdSiteMap_raised_semifocus_dots_img.Visible = True
12050     .cmdSiteMap_raised_img.Visible = False
12060     .cmdSiteMap_raised_focus_img.Visible = False
12070     .cmdSiteMap_raised_focus_dots_img.Visible = False
12080     .cmdSiteMap_sunken_focus_dots_img.Visible = False
12090     .cmdSiteMap_raised_img_dis.Visible = False
12100     .cmdUtility_box01.Visible = False
12110     .cmdUtility_box02.Visible = False
12120     .cmdAsset_box01.Visible = False
12130     .cmdAsset_box02.Visible = False
12140     .cmdReport_box01.Visible = False
12150     .cmdReport_box02.Visible = False
12160     .cmdAccount_box01.Visible = False
12170     .cmdAccount_box02.Visible = False
12180     .cmdPost_box01.Visible = False
12190     .cmdPost_box02.Visible = False
12200   End With

EXITP:
12210   Exit Sub

ERRH:
12220   Select Case ERR.Number
        Case Else
12230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12240   End Select
12250   Resume EXITP

End Sub

Private Sub cmdSiteMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSiteMap_MouseDown"

12310   With Me
12320     blnSiteMap_MouseDown = True
12330     .cmdSiteMap_sunken_focus_dots_img.Visible = True
12340     .cmdSiteMap_raised_img.Visible = False
12350     .cmdSiteMap_raised_semifocus_dots_img.Visible = False
12360     .cmdSiteMap_raised_focus_img.Visible = False
12370     .cmdSiteMap_raised_focus_dots_img.Visible = False
12380     .cmdSiteMap_raised_img_dis.Visible = False
12390   End With

EXITP:
12400   Exit Sub

ERRH:
12410   Select Case ERR.Number
        Case Else
12420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12430   End Select
12440   Resume EXITP

End Sub

Private Sub cmdSiteMap_Click()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSiteMap_Click"

        Dim strDocName As String

12510   With Me
12520     If blnFormOpening = False Then

12530       .TimerInterval = 0&

12540       gblnSetFocus = True
12550       blnFormOpening = True
12560       gstrReturningForm = vbNullString
12570       strDocName = "frmSiteMap"
12580       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
12590       blnFormOpening = False

12600     End If
12610   End With

EXITP:
12620   Exit Sub

ERRH:
12630   blnFormOpening = False
12640   DoCmd.Hourglass False
12650   Select Case ERR.Number
        Case Else
12660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12670   End Select
12680   Resume EXITP

End Sub

Private Sub cmdSiteMap_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSiteMap_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12720   intShiftDown = (Shift And acShiftMask) > 0
12730   intAltDown = (Shift And acAltMask) > 0
12740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12760     Select Case intRetVal
          Case vbKeyTab
12770       With Me
12780         intRetVal = 0
12790         .cmdAbout.SetFocus
12800       End With
12810     End Select
12820   End If

        ' ** Shift keys.
12830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12840     Select Case intRetVal
          Case vbKeyTab
12850       With Me
12860         intRetVal = 0
12870         .cmdUtility.SetFocus
12880       End With
12890     End Select
12900   End If

EXITP:
12910   KeyCode = intRetVal
12920   Exit Sub

ERRH:
12930   intRetVal = 0
12940   Select Case ERR.Number
        Case Else
12950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12960   End Select
12970   Resume EXITP

End Sub

Private Sub cmdSiteMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSiteMap_MouseMove"

13010   With Me
13020     If blnSiteMap_MouseDown = False Then
13030       Select Case blnSiteMap_Focus
            Case True
13040         .cmdSiteMap_raised_focus_dots_img.Visible = True
13050         .cmdSiteMap_raised_focus_img.Visible = False
13060       Case False
13070         .cmdSiteMap_raised_focus_img.Visible = True
13080         .cmdSiteMap_raised_focus_dots_img.Visible = False
13090       End Select
13100       .cmdSiteMap_raised_img.Visible = False
13110       .cmdSiteMap_raised_semifocus_dots_img.Visible = False
13120       .cmdSiteMap_sunken_focus_dots_img.Visible = False
13130       .cmdSiteMap_raised_img_dis.Visible = False
13140     End If
13150     If .cmdAbout_raised_focus_dots_img.Visible = True Or .cmdAbout_raised_focus_img.Visible = True Then
13160       Select Case blnAbout_Focus
            Case True
13170         .cmdAbout_raised_semifocus_dots_img.Visible = True
13180         .cmdAbout_raised_img.Visible = False
13190       Case False
13200         .cmdAbout_raised_img.Visible = True
13210         .cmdAbout_raised_semifocus_dots_img.Visible = False
13220       End Select
13230       .cmdAbout_raised_focus_img.Visible = False
13240       .cmdAbout_raised_focus_dots_img.Visible = False
13250       .cmdAbout_sunken_focus_dots_img.Visible = False
13260       .cmdAbout_raised_img_dis.Visible = False
13270     End If
13280   End With

EXITP:
13290   Exit Sub

ERRH:
13300   Select Case ERR.Number
        Case Else
13310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13320   End Select
13330   Resume EXITP

End Sub

Private Sub cmdSiteMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSiteMap_MouseUp"

13410   With Me
13420     .cmdSiteMap_raised_focus_dots_img.Visible = True
13430     .cmdSiteMap_raised_img.Visible = False
13440     .cmdSiteMap_raised_semifocus_dots_img.Visible = False
13450     .cmdSiteMap_raised_focus_img.Visible = False
13460     .cmdSiteMap_sunken_focus_dots_img.Visible = False
13470     .cmdSiteMap_raised_img_dis.Visible = False
13480     blnSiteMap_MouseDown = False
13490   End With

EXITP:
13500   Exit Sub

ERRH:
13510   Select Case ERR.Number
        Case Else
13520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13530   End Select
13540   Resume EXITP

End Sub

Private Sub cmdSiteMap_LostFocus()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSiteMap_LostFocus"

13610   With Me
13620     .cmdSiteMap_raised_img.Visible = True
13630     .cmdSiteMap_raised_semifocus_dots_img.Visible = False
13640     .cmdSiteMap_raised_focus_img.Visible = False
13650     .cmdSiteMap_raised_focus_dots_img.Visible = False
13660     .cmdSiteMap_sunken_focus_dots_img.Visible = False
13670     .cmdSiteMap_raised_img_dis.Visible = False
13680     blnSiteMap_Focus = False
13690   End With

EXITP:
13700   Exit Sub

ERRH:
13710   Select Case ERR.Number
        Case Else
13720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13730   End Select
13740   Resume EXITP

End Sub

Private Sub cmdAbout_GotFocus()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAbout_GotFocus"

13810   With Me
13820     strLastFocus = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
13830     blnAbout_Focus = True
13840     .cmdAbout_raised_semifocus_dots_img.Visible = True
13850     .cmdAbout_raised_img.Visible = False
13860     .cmdAbout_raised_focus_img.Visible = False
13870     .cmdAbout_raised_focus_dots_img.Visible = False
13880     .cmdAbout_sunken_focus_dots_img.Visible = False
13890     .cmdAbout_raised_img_dis.Visible = False
13900     .cmdUtility_box01.Visible = False
13910     .cmdUtility_box02.Visible = False
13920     .cmdAsset_box01.Visible = False
13930     .cmdAsset_box02.Visible = False
13940     .cmdReport_box01.Visible = False
13950     .cmdReport_box02.Visible = False
13960     .cmdAccount_box01.Visible = False
13970     .cmdAccount_box02.Visible = False
13980     .cmdPost_box01.Visible = False
13990     .cmdPost_box02.Visible = False
14000   End With

EXITP:
14010   Exit Sub

ERRH:
14020   Select Case ERR.Number
        Case Else
14030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14040   End Select
14050   Resume EXITP

End Sub

Private Sub cmdAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAbout_MouseDown"

14110   With Me
14120     blnAbout_MouseDown = True
14130     .cmdAbout_sunken_focus_dots_img.Visible = True
14140     .cmdAbout_raised_img.Visible = False
14150     .cmdAbout_raised_semifocus_dots_img.Visible = False
14160     .cmdAbout_raised_focus_img.Visible = False
14170     .cmdAbout_raised_focus_dots_img.Visible = False
14180     .cmdAbout_raised_img_dis.Visible = False
14190   End With

EXITP:
14200   Exit Sub

ERRH:
14210   Select Case ERR.Number
        Case Else
14220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14230   End Select
14240   Resume EXITP

End Sub

Private Sub cmdAbout_Click()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAbout_Click"

14310   With Me
14320     If blnFormOpening = False Then

14330       .TimerInterval = 0&

14340       blnFormOpening = True
14350       strDocName = "frmAbout"
14360       gstrReturningForm = strDocName
14370       DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME
14380       blnFormOpening = False
14390       gblnSetFocus = True
14400       .TimerInterval = 100&

14410     End If
14420   End With

EXITP:
14430   Exit Sub

ERRH:
14440   blnFormOpening = False
14450   DoCmd.Hourglass False
14460   Select Case ERR.Number
        Case Else
14470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14480   End Select
14490   Resume EXITP

End Sub

Private Sub cmdAbout_KeyDown(KeyCode As Integer, Shift As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAbout_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14520   intShiftDown = (Shift And acShiftMask) > 0
14530   intAltDown = (Shift And acAltMask) > 0
14540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14560     Select Case intRetVal
          Case vbKeyTab
14570       With Me
14580         intRetVal = 0
14590         If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
14600           .chkLoadTimer.SetFocus
14610         Else
14620           .cmdClose.SetFocus
14630         End If
14640       End With
14650     End Select
14660   End If

        ' ** Shift keys.
14670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14680     Select Case intRetVal
          Case vbKeyTab
14690       With Me
14700         intRetVal = 0
14710         .cmdSiteMap.SetFocus
14720       End With
14730     End Select
14740   End If

EXITP:
14750   KeyCode = intRetVal
14760   Exit Sub

ERRH:
14770   intRetVal = 0
14780   Select Case ERR.Number
        Case Else
14790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14800   End Select
14810   Resume EXITP

End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAbout_MouseMove"

14910   With Me
14920     If blnAbout_MouseDown = False Then
14930       Select Case blnAbout_Focus
            Case True
14940         .cmdAbout_raised_focus_dots_img.Visible = True
14950         .cmdAbout_raised_focus_img.Visible = False
14960       Case False
14970         .cmdAbout_raised_focus_img.Visible = True
14980         .cmdAbout_raised_focus_dots_img.Visible = False
14990       End Select
15000       .cmdAbout_raised_img.Visible = False
15010       .cmdAbout_raised_semifocus_dots_img.Visible = False
15020       .cmdAbout_sunken_focus_dots_img.Visible = False
15030       .cmdAbout_raised_img_dis.Visible = False
15040     End If
15050     If .cmdSiteMap_raised_focus_dots_img.Visible = True Or .cmdSiteMap_raised_focus_img.Visible = True Then
15060       Select Case blnSiteMap_Focus
            Case True
15070         .cmdSiteMap_raised_semifocus_dots_img.Visible = True
15080         .cmdSiteMap_raised_img.Visible = False
15090       Case False
15100         .cmdSiteMap_raised_img.Visible = True
15110         .cmdSiteMap_raised_semifocus_dots_img.Visible = False
15120       End Select
15130       .cmdSiteMap_raised_focus_img.Visible = False
15140       .cmdSiteMap_raised_focus_dots_img.Visible = False
15150       .cmdSiteMap_sunken_focus_dots_img.Visible = False
15160       .cmdSiteMap_raised_img_dis.Visible = False
15170     End If
15180   End With

EXITP:
15190   Exit Sub

ERRH:
15200   Select Case ERR.Number
        Case Else
15210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15220   End Select
15230   Resume EXITP

End Sub

Private Sub cmdAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAbout_MouseUp"

15310   With Me
15320     .cmdAbout_raised_focus_dots_img.Visible = True
15330     .cmdAbout_raised_img.Visible = False
15340     .cmdAbout_raised_semifocus_dots_img.Visible = False
15350     .cmdAbout_raised_focus_img.Visible = False
15360     .cmdAbout_sunken_focus_dots_img.Visible = False
15370     .cmdAbout_raised_img_dis.Visible = False
15380     blnAbout_MouseDown = False
15390   End With

EXITP:
15400   Exit Sub

ERRH:
15410   Select Case ERR.Number
        Case Else
15420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15430   End Select
15440   Resume EXITP

End Sub

Private Sub cmdAbout_LostFocus()

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAbout_LostFocus"

15510   With Me
15520     .cmdAbout_raised_img.Visible = True
15530     .cmdAbout_raised_focus_img.Visible = False
15540     .cmdAbout_raised_semifocus_dots_img.Visible = False
15550     .cmdAbout_raised_focus_dots_img.Visible = False
15560     .cmdAbout_sunken_focus_dots_img.Visible = False
15570     .cmdAbout_raised_img_dis.Visible = False
15580     blnAbout_Focus = False
15590   End With

EXITP:
15600   Exit Sub

ERRH:
15610   Select Case ERR.Number
        Case Else
15620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15630   End Select
15640   Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

15710   With Me
15720     gblnTimer = False
15730     Select Case .chkLoadTimer
          Case True
15740       .chkLoadTimer_lbl.FontBold = True
15750       If .chkLoadTimer.Visible = True Then
15760         gblnTimer = True
15770       End If
15780     Case False
15790       .chkLoadTimer_lbl.FontBold = False
15800     End Select
15810   End With

EXITP:
15820   Exit Sub

ERRH:
15830   Select Case ERR.Number
        Case Else
15840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15850   End Select
15860   Resume EXITP

End Sub

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15920   intShiftDown = (Shift And acShiftMask) > 0
15930   intAltDown = (Shift And acAltMask) > 0
15940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15960     Select Case intRetVal
          Case vbKeyTab
15970       With Me
15980         intRetVal = 0
15990         .cmdClose.SetFocus
16000       End With
16010     End Select
16020   End If

        ' ** Shift keys.
16030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16040     Select Case intRetVal
          Case vbKeyTab
16050       With Me
16060         intRetVal = 0
16070         .cmdAbout.SetFocus
16080       End With
16090     End Select
16100   End If

EXITP:
16110   KeyCode = intRetVal
16120   Exit Sub

ERRH:
16130   intRetVal = 0
16140   Select Case ERR.Number
        Case Else
16150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16160   End Select
16170   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

16210   With Me
16220     blnGTR_Emblem = False
16230     For lngX = 1& To 24&
16240 On Error Resume Next
16250       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
16260       If ERR.Number <> 0 Then
16270 On Error GoTo ERRH
16280         Exit For
16290       Else
16300 On Error GoTo ERRH
16310       End If
16320     Next
16330   End With

EXITP:
16340   Exit Sub

ERRH:
16350   Select Case ERR.Number
        Case Else
16360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16370   End Select
16380   Resume EXITP

End Sub

Public Sub GTRSite_Off()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "GTRSite_Off"

16410   With Me
16420     GTREmblem_Off  ' ** Procedure: Above.
16430     blnGoingToReport = False
16440     .GoToReport_arw_site_img.Visible = False
16450   End With

EXITP:
16460   Exit Sub

ERRH:
16470   Select Case ERR.Number
        Case Else
16480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16490   End Select
16500   Resume EXITP

End Sub
