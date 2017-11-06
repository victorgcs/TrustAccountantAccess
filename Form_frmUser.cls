VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmUser"

'VGC 09/23/2017: CHANGES!

'ADD CAPABILITY TO DELETE EXTRANEOUS, OLD USERS FROM MDW!  ' ** MAY ALREADY BE DONE!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd}
' **   About User:       B {cmdHelp}
' **   Close:            C {cmdClose}
' **   Delete:           D {cmdDelete}
' **   GUID Test:        G {cmdGUIDTest}
' **   Show All Users:   H {cmdShowAll}
' **   Print Report:     P {cmdPrintReport}
' **   User Name:        U {Username on frmUser_Sub}
' **   Change Password:  W {cmdChangePassword}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmUser_Sub}

' **********************************************************************************************
' ** cmdHelp only shows for those logged in as one of the non-changeable default users.
' **********************************************************************************************

' ** VGC 05/16/08: I do not know what the fields s_Generation and s_Lineage are for.
' ** They don't seem to be used.
' ** VGC 11/11/09: They are remnants of data replication. I'm leaving them in
' ** because the s_GUID field is a GUID AutoNumber
' ** GUID: Globally Unique IDentifier

Private Const FILTER_DEF As String = "[Username] <> 'Superuser' And [Username] <> 'Admin' And [Username] <> 'Creator' And [Username] <> 'Engine'"

Private Const strSortOrig As String = "[Username]"

Private CLR_DISABLED_BG As Long

Private blnHelp_Focus As Boolean, blnHelp_MouseDown As Boolean
Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long
Private usr As DAO.User
Private blnAdmin As Boolean, blnUserNameEdit As Boolean, blnPWChange As Boolean
Private lngRecsCur As Long
Private strSortNow As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim grp As DAO.Group
        Dim strFilter As String
        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngScreenWidth As Long, lngScreenHeight As Long
        Dim lngXBorders As Long, lngYBorders As Long
        Dim lngTmpTop As Long, lngTmpHeight As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnAdmin = False: blnUserNameEdit = False: blnPWChange = False
160           blnHelp_Focus = False: blnHelp_MouseDown = False
170           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

180           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

190           CLR_DISABLED_BG = CLR_LTTEAL

200           MouseWheelON  ' ** Module Function: modMouseWheel.

210           .ThisUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
220           .frmUser_Sub.Form.Username_IsLocked.ForeColor = CLR_DISABLED_BG
230           .frmUser_Sub.Form.Employee_Name_IsLocked.ForeColor = CLR_DISABLED_BG
240           .frmUser_Sub.Form.Primary_Group_IsLocked.ForeColor = CLR_DISABLED_BG

              ' ** Moved finer security control from the menu calling this.
250           For Each grp In DBEngine.Workspaces(0).Groups
260             If grp.Name = "Admins" Then
270               For Each usr In grp.Users
280                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
290                   blnAdmin = True
300                   Exit For
310                 End If
320               Next
330             End If
340           Next

              ' ** Variables are fed empty, then populated ByRef.
350           GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

360           lngTmp01 = .Section(0).Height  ' ** Detail, in twips.
370           lngTmp02 = .Section(1).Height  ' ** Header.
380           lngTmp03 = .Section(2).Height  ' ** Footer.

390           lngScreenWidth = GetScreenRes(GSR_HORZRES)  ' ** Module Function: modWindowFunctions.
400           lngScreenHeight = GetScreenRes(GSR_VERTRES)  ' ** Module Function: modWindowFunctions.
410           ConvertPIXELSToTWIPS lngScreenWidth, lngScreenHeight  ' ** Module Function: modWindowFunctions.

420           GetWindowBorders Me, lngXBorders, lngYBorders  ' ** Module Function: modWindowFunctions.

430           lngTmpTop = lngTop  ' ** Current form top.
440           lngTmpHeight = ((lngTmp02 + lngTmp03 + (lngTmp01 * RZ_WINROWS)) + lngYBorders)  ' ** Ideal form height.
450           lngTmp04 = ((lngScreenHeight - lngTmpHeight) / 2)  ' ** Top of centered ideal form.
              ' ** If the (current top + lngTmpHeight) is within lngScreenHeight then...
460           If (lngTmpTop + lngTmpHeight) < lngScreenHeight Then
                ' ** If the bottom gutter is smaller than the top, move it up.
470             If (lngScreenHeight - (lngTmpTop + lngTmpHeight)) < lngTmpTop Then
480               lngTmpTop = lngTmp04
490             End If
500           Else
                ' ** If the ideal height is showable, then...
510             If lngTmpHeight < lngScreenHeight Then
                  ' ** If the (current top + lngTmpHeight) puts the form bottom out of sight, then...
520               If (lngTmpTop + lngTmpHeight) > lngScreenHeight Then
530                 lngTmpTop = ((lngScreenHeight - lngTmpHeight) / 2&)
540               End If
550             Else
560               lngTmpTop = 300&
570               lngTmpHeight = (lngScreenHeight - (lngTmpTop * 2&))
580             End If
590           End If
              'DoCmd.MoveSize lngLeft, lngTmpTop, lngWidth, lngTmpHeight

600           strFilter = FILTER_DEF
610           .frmUser_Sub.Form.Filter = strFilter
620           .frmUser_Sub.Form.FilterOn = True

630           .frmUser_Sub.Form.OrderBy = strSortOrig
640           strSortNow = strSortOrig
650           .frmUser_Sub.Form.OrderByOn = True
660           .frmUser_Sub.Form.Sort_line.Visible = True

              ' ** Disable controls for lower security levels.
670           Select Case blnAdmin
              Case True
680             .IsAdmins = True
690           Case False
700             .frmUser_Sub.Form.Filter = "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
710             .frmUser_Sub.Form.FilterOn = True
720             .cmdClose.SetFocus
730             .cmdAdd.Visible = False
740             .cmdDelete.Visible = False
750             .frmUser_Sub.Form.FocusHolder.SetFocus
760             .frmUser_Sub.Form.Username.Enabled = False
770             .frmUser_Sub.Form.Username.Locked = True
780             .frmUser_Sub.Form.Employee_Name.Enabled = False
790             .frmUser_Sub.Form.Employee_Name.Locked = True
800             .frmUser_Sub.Form.Primary_Group.Enabled = False
810             .frmUser_Sub.Form.Primary_Group.Locked = True
820             .cmdPrintReport.Enabled = False
830             .cmdPrintReport_raised_img_dis.Visible = True
840             .cmdPrintReport_raised_img.Visible = False
850             .cmdPrintReport_raised_semifocus_dots_img.Visible = False
860             .cmdPrintReport_raised_focus_img.Visible = False
870             .cmdPrintReport_raised_focus_dots_img.Visible = False
880             .cmdPrintReport_sunken_focus_dots_img.Visible = False
890           End Select
900           DefaultUserHelp  ' ** Function: Below.

910           If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
                '.cmdShowAll.Visible = True
920           Else
930             .cmdShowAll.Visible = False
940           End If

950         Else
960           Cancel = -1
970         End If
980       Else
990         Cancel = -1
1000      End If

1010    End With

1020    If Cancel = -1 Then
1030      Beep
1040      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1050    End If

EXITP:
1060    Set grp = Nothing
1070    Exit Sub

ERRH:
1080    Select Case ERR.Number
        Case Else
1090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1100    End Select
1110    Resume EXITP

End Sub

Private Sub Form_Load()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1210    With Me
1220      lngRecsCur = .frmUser_Sub.Form.RecCnt  ' ** Form Function: frmUser_Sub.
1230      If lngRecsCur > 0& Then
1240        .frmUser_Sub.Form.Username.SetFocus
1250      Else
1260        .cmdDelete.Enabled = False
1270        .cmdChangePassword.Enabled = False
1280        If .cmdAdd.Enabled = True Then
1290          .cmdAdd.SetFocus
1300        Else
1310          .cmdClose.SetFocus
1320        End If
1330      End If
1340    End With

EXITP:
1350    Exit Sub

ERRH:
1360    Select Case ERR.Number
        Case Else
1370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1380    End Select
1390    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1420    intShiftDown = (Shift And acShiftMask) > 0
1430    intAltDown = (Shift And acAltMask) > 0
1440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyEscape
1470        intRetVal = 0
1480        cmdClose_Click  ' ** Procedure: Below.
1490      Case vbKeyDelete
1500        With Me
1510          intRetVal = 0
1520          If .cmdDelete.Enabled = True Then
1530            cmdDelete_Click  ' ** Procedure: Below.
1540          Else
1550            Beep
1560          End If
1570        End With
1580      End Select
1590    End If

        ' ** Alt keys.
1600    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1610      Select Case intRetVal
          Case vbKeyU
1620        With Me
1630          intRetVal = 0
1640          .frmUser_Sub.SetFocus
1650          .frmUser_Sub.Form.Username.SetFocus
1660        End With
1670      Case vbKeyX
1680        intRetVal = 0
1690        cmdClose_Click  ' ** Procedure: Below.
1700      End Select
1710    End If

        ' ** Ctrl keys.
1720    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1730      Select Case intRetVal
          Case vbKeyS
1740        With Me
1750          .frmUser_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmUser_Sub.
1760        End With
1770      End Select
1780    End If

EXITP:
1790    KeyCode = intRetVal
1800    Exit Sub

ERRH:
1810    intRetVal = 0
1820    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1830    Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub Form_Timer()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1910    With Me
1920      .TimerInterval = 0&
1930      If gblnSetFocus = True Then
1940        gblnSetFocus = False
1950        strFormName = vbNullString
1960  On Error Resume Next
1970        Do Until strFormName = THIS_NAME
1980          DoCmd.SelectObject acForm, THIS_NAME, False
1990          strFormName = Screen.ActiveForm.Name
2000        Loop
2010  On Error GoTo ERRH
2020        .frmUser_Sub.SetFocus
2030      End If  ' ** gblnSetFocus
2040      If gblnGoToReport = True Then
2050        Select Case blnGoingToReport
            Case True
2060          DoCmd.Hourglass False
2070          gblnGoToReport = False  ' ** It should already be off.
2080          blnGoingToReport = False
2090          .GoToReport_arw_pr_img.Visible = False
2100          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
2110          DoEvents
2120        Case False
2130          DoCmd.Hourglass True  ' ** Make sure it's still running.
2140          DoEvents
2150          blnGoingToReport = True
2160          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
2170          .GoToReport_arw_pr_img.Visible = True
2180          .cmdPrintReport.SetFocus
2190          DoEvents
2200          Beep
2210          DoCmd.Hourglass False
2220          DoEvents
2230          .TimerInterval = (GTR_WAIT * 2&)
2240        End Select
2250      Else
2260        DoCmd.Hourglass False
2270      End If
2280    End With

EXITP:
2290    Exit Sub

ERRH:
2300    Select Case ERR.Number
        Case Else
2310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2320    End Select
2330    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim blnRetVal As Boolean

2410    Set usr = Nothing

2420    blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.

2430    Set dbs = CurrentDb
2440    With dbs
          ' ** Delete qrySecurity_Sync_12 (_~xusr, not in Users).
2450      Set qdf = .QueryDefs("qrySecurity_Sync_13")
2460      qdf.Execute
2470      .Close
2480    End With

        ' ** Make sure all reports are closed.
2490    Do While Reports.Count > 0
2500      DoCmd.Close acReport, Reports(0).Name
2510    Loop

        ' ** Make sure all forms closed.
2520    strDocName = "frmUser_Add"
2530    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Functions: modFileUtilities.
2540      DoCmd.Close acForm, strDocName
2550    End If
2560    strDocName = "frmUser_Password"
2570    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Functions: modFileUtilities.
2580      DoCmd.Close acForm, strDocName
2590    End If

2600    gstrReturningForm = THIS_NAME
2610    strDocName = "frmMenu_Utility"
2620    If strCallingForm <> vbNullString Then
2630      Select Case strCallingForm
          Case "frmMenu_Utility"
2640        gblnSetFocus = True
2650        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2660      Case Else
2670        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
2680      End Select
2690    Else
2700      gblnSetFocus = True
2710      DoCmd.OpenForm strDocName, , , , , , "frmMenu_Main"
2720    End If

EXITP:
2730    Set qdf = Nothing
2740    Set dbs = Nothing
2750    Exit Sub

ERRH:
2760    Select Case ERR.Number
        Case Else
2770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2780    End Select
2790    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2810    With Me
2820      If blnAdmin = True And (.ThisUser = "TAAdmin" Or .ThisUser = "TADemo") Then
2830        If ((.cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = True Or .cmdHelp_TAAdmin_R_raised_focus_img.Visible = True) Or _
                (.cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = True Or .cmdHelp_TAAdmin_L_raised_focus_img.Visible = True)) Or _
                ((.cmdHelp_TADemo_R_raised_focus_dots_img.Visible = True Or .cmdHelp_TADemo_R_raised_focus_img.Visible = True) Or _
                (.cmdHelp_TADemo_L_raised_focus_dots_img.Visible = True Or .cmdHelp_TADemo_L_raised_focus_img.Visible = True)) Then
2840          Select Case .ThisUser
              Case "TAAdmin"
2850            Select Case .cmdHelp_lbl1.Visible
                Case True
2860              Select Case blnHelp_Focus
                  Case True
2870                .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = True
2880                .cmdHelp_TADemo_R_raised_img.Visible = False
2890                .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
2900                .cmdHelp_TADemo_L_raised_img.Visible = False
2910                .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
2920                .cmdHelp_TAAdmin_R_raised_img.Visible = False
2930                .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
2940                .cmdHelp_TAAdmin_L_raised_img.Visible = False
2950              Case False
2960                .cmdHelp_TAAdmin_L_raised_img.Visible = True
2970                .cmdHelp_TADemo_R_raised_img.Visible = False
2980                .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
2990                .cmdHelp_TADemo_L_raised_img.Visible = False
3000                .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
3010                .cmdHelp_TAAdmin_R_raised_img.Visible = False
3020                .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
3030                .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
3040              End Select
3050            Case False
3060              Select Case blnHelp_Focus
                  Case True
3070                .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = True
3080                .cmdHelp_TADemo_R_raised_img.Visible = False
3090                .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
3100                .cmdHelp_TADemo_L_raised_img.Visible = False
3110                .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
3120                .cmdHelp_TAAdmin_R_raised_img.Visible = False
3130                .cmdHelp_TAAdmin_L_raised_img.Visible = False
3140                .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
3150              Case False
3160                .cmdHelp_TAAdmin_R_raised_img.Visible = True
3170                .cmdHelp_TADemo_R_raised_img.Visible = False
3180                .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
3190                .cmdHelp_TADemo_L_raised_img.Visible = False
3200                .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
3210                .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
3220                .cmdHelp_TAAdmin_L_raised_img.Visible = False
3230                .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
3240              End Select
3250            End Select
3260          Case "TADemo"
3270            Select Case .cmdHelp_lbl1.Visible
                Case True
3280              Select Case blnHelp_Focus
                  Case True
3290                .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = True
3300                .cmdHelp_TADemo_R_raised_img.Visible = False
3310                .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
3320                .cmdHelp_TADemo_L_raised_img.Visible = False
3330                .cmdHelp_TAAdmin_R_raised_img.Visible = False
3340                .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
3350                .cmdHelp_TAAdmin_L_raised_img.Visible = False
3360                .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
3370              Case False
3380                .cmdHelp_TADemo_L_raised_img.Visible = True
3390                .cmdHelp_TADemo_R_raised_img.Visible = False
3400                .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
3410                .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
3420                .cmdHelp_TAAdmin_R_raised_img.Visible = False
3430                .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
3440                .cmdHelp_TAAdmin_L_raised_img.Visible = False
3450                .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
3460              End Select
3470            Case False
3480              Select Case blnHelp_Focus
                  Case True
3490                .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = True
3500                .cmdHelp_TADemo_R_raised_img.Visible = False
3510                .cmdHelp_TADemo_L_raised_img.Visible = False
3520                .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
3530                .cmdHelp_TAAdmin_R_raised_img.Visible = False
3540                .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
3550                .cmdHelp_TAAdmin_L_raised_img.Visible = False
3560                .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
3570              Case False
3580                .cmdHelp_TADemo_R_raised_img.Visible = True
3590                .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
3600                .cmdHelp_TADemo_L_raised_img.Visible = False
3610                .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
3620                .cmdHelp_TAAdmin_R_raised_img.Visible = False
3630                .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
3640                .cmdHelp_TAAdmin_L_raised_img.Visible = False
3650                .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
3660              End Select
3670            End Select
3680          End Select
3690          .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
3700          .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
3710          .cmdHelp_TAAdmin_R_sunken_focus_dots_img.Visible = False
3720          .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
3730          .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
3740          .cmdHelp_TAAdmin_L_sunken_focus_dots_img.Visible = False
3750          .cmdHelp_TADemo_R_raised_focus_img.Visible = False
3760          .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
3770          .cmdHelp_TADemo_R_sunken_focus_dots_img.Visible = False
3780          .cmdHelp_TADemo_L_raised_focus_img.Visible = False
3790          .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
3800          .cmdHelp_TADemo_L_sunken_focus_dots_img.Visible = False
3810        End If
3820      End If
3830      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3840        Select Case blnPrintReport_Focus
            Case True
3850          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3860          .cmdPrintReport_raised_img.Visible = False
3870        Case False
3880          .cmdPrintReport_raised_img.Visible = True
3890          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3900        End Select
3910        .cmdPrintReport_raised_focus_img.Visible = False
3920        .cmdPrintReport_raised_focus_dots_img.Visible = False
3930        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3940        .cmdPrintReport_raised_img_dis.Visible = False
3950      End If
3960    End With

EXITP:
3970    Exit Sub

ERRH:
3980    Select Case ERR.Number
        Case Else
3990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4000    End Select
4010    Resume EXITP

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

Public Sub cmdClose_Click()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4210    DoCmd.Close acForm, THIS_NAME

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4320    intShiftDown = (Shift And acShiftMask) > 0
4330    intAltDown = (Shift And acAltMask) > 0
4340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4360      Select Case intRetVal
          Case vbKeyTab
4370        With Me
4380          intRetVal = 0
4390          lngRecsCur = .frmUser_Sub.Form.RecCnt  ' ** Form Function: frmUser_Sub.
4400          If lngRecsCur > 0& Then
4410            .frmUser_Sub.SetFocus
4420            .frmUser_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmUser_Sub.
4430            .frmUser_Sub.Form.Username.SetFocus
4440          Else
4450            If .cmdChangePassword.Enabled = True Then
4460              .cmdChangePassword.SetFocus
4470            ElseIf .cmdAdd.Enabled = True Then
4480              .cmdAdd.SetFocus
4490            ElseIf .cmdDelete.Enabled = True Then  ' ** Yes, the question is moot!
4500              .cmdDelete.SetFocus
4510            Else
                  ' ** Nowhere to go!
4520              Beep
4530            End If
4540          End If
4550        End With
4560      End Select
4570    End If

        ' ** Shift keys.
4580    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4590      Select Case intRetVal
          Case vbKeyTab
4600        With Me
4610          intRetVal = 0
4620          If .cmdDelete.Enabled = True Then
4630            .cmdDelete.SetFocus
4640          ElseIf .cmdAdd.Enabled = True Then
4650            .cmdAdd.SetFocus
4660          ElseIf .cmdChangePassword.Enabled = True Then
4670            .cmdChangePassword.SetFocus
4680          Else
4690            lngRecsCur = .frmUser_Sub.Form.RecCnt  ' ** Form Function: frmUser_Sub.
4700            If lngRecsCur > 0& Then
4710              .frmUser_Sub.SetFocus
4720              .frmUser_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmUser_Sub.
4730              If .frmUser_Sub.Form.Primary_Group.Enabled = True Then
4740                .frmUser_Sub.Form.Primary_Group.SetFocus
4750              Else
4760                .frmUser_Sub.Form.Employee_Name.SetFocus
4770              End If
4780            Else
                  ' ** Nowhere to go!
4790              Beep
4800            End If
4810          End If
4820        End With
4830      End Select
4840    End If

EXITP:
4850    KeyCode = intRetVal
4860    Exit Sub

ERRH:
4870    Select Case ERR.Number
        Case Else
4880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4890    End Select
4900    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

5010    With Me
5020      blnPrintReport_Focus = True
5030      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
5040      .cmdPrintReport_raised_img.Visible = False
5050      .cmdPrintReport_raised_focus_img.Visible = False
5060      .cmdPrintReport_raised_focus_dots_img.Visible = False
5070      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5080      .cmdPrintReport_raised_img_dis.Visible = False
5090    End With

EXITP:
5100    Exit Sub

ERRH:
5110    Select Case ERR.Number
        Case Else
5120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5130    End Select
5140    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

5210    With Me
5220      blnPrintReport_MouseDown = True
5230      .cmdPrintReport_sunken_focus_dots_img.Visible = True
5240      .cmdPrintReport_raised_img.Visible = False
5250      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5260      .cmdPrintReport_raised_focus_img.Visible = False
5270      .cmdPrintReport_raised_focus_dots_img.Visible = False
5280      .cmdPrintReport_raised_img_dis.Visible = False
5290    End With

EXITP:
5300    Exit Sub

ERRH:
5310    Select Case ERR.Number
        Case Else
5320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5330    End Select
5340    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

5410    With Me
5420      strDocName = "rptListOfUsers"
5430      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
5440        DoCmd.OpenReport strDocName, acViewPreview
5450        DoCmd.Maximize
5460        DoCmd.RunCommand acCmdFitToWindow
5470        .Visible = False
5480      Else
            '##GTR_Ref: rptListOfUsers
5490        DoCmd.OpenReport strDocName, acViewNormal
5500      End If
5510    End With

EXITP:
5520    Exit Sub

ERRH:
5530    Me.Visible = True
5540    DoCmd.Restore
5550    Select Case ERR.Number
        Case Else
5560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5570    End Select
5580    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5620    intShiftDown = (Shift And acShiftMask) > 0
5630    intAltDown = (Shift And acAltMask) > 0
5640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5660      Select Case intRetVal
          Case vbKeyTab
5670        intRetVal = 0
5680        With Me
5690          If .cmdChangePassword.Enabled = True Then
5700            .cmdChangePassword.SetFocus
5710          ElseIf .cmdAdd.Enabled = True Then
5720            .cmdAdd.SetFocus
5730          ElseIf .cmdDelete.Enabled = True Then
5740            .cmdDelete.SetFocus
5750          Else
5760            .cmdClose.SetFocus
5770          End If
5780        End With
5790      End Select
5800    End If

        ' ** Shift keys.
5810    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5820      Select Case intRetVal
          Case vbKeyTab
5830        intRetVal = 0
5840        With Me
5850          lngRecsCur = .frmUser_Sub.Form.RecCnt  ' ** Form Function: frmUser_Sub.
5860          If lngRecsCur > 0& Then
5870            .frmUser_Sub.SetFocus
5880            .frmUser_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmUser_Sub.
5890            If .frmUser_Sub.Form.Primary_Group.Enabled = True Then
5900              .frmUser_Sub.Form.Primary_Group.SetFocus
5910            Else
5920              .frmUser_Sub.Form.Employee_Name.SetFocus
5930            End If
5940          Else
5950            .cmdClose.SetFocus
5960          End If
5970        End With
5980      End Select
5990    End If

EXITP:
6000    KeyCode = intRetVal
6010    Exit Sub

ERRH:
6020    Select Case ERR.Number
        Case Else
6030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6040    End Select
6050    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

6110    With Me
6120      If blnPrintReport_MouseDown = False Then
6130        Select Case blnPrintReport_Focus
            Case True
6140          .cmdPrintReport_raised_focus_dots_img.Visible = True
6150          .cmdPrintReport_raised_focus_img.Visible = False
6160        Case False
6170          .cmdPrintReport_raised_focus_img.Visible = True
6180          .cmdPrintReport_raised_focus_dots_img.Visible = False
6190        End Select
6200        .cmdPrintReport_raised_img.Visible = False
6210        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6220        .cmdPrintReport_sunken_focus_dots_img.Visible = False
6230        .cmdPrintReport_raised_img_dis.Visible = False
6240      End If
6250    End With

EXITP:
6260    Exit Sub

ERRH:
6270    Select Case ERR.Number
        Case Else
6280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6290    End Select
6300    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

6410    With Me
6420      .cmdPrintReport_raised_focus_dots_img.Visible = True
6430      .cmdPrintReport_raised_img.Visible = False
6440      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6450      .cmdPrintReport_raised_focus_img.Visible = False
6460      .cmdPrintReport_sunken_focus_dots_img.Visible = False
6470      .cmdPrintReport_raised_img_dis.Visible = False
6480      blnPrintReport_MouseDown = False
6490    End With

EXITP:
6500    Exit Sub

ERRH:
6510    Select Case ERR.Number
        Case Else
6520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6530    End Select
6540    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

6610    With Me
6620      .cmdPrintReport_raised_img.Visible = True
6630      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6640      .cmdPrintReport_raised_focus_img.Visible = False
6650      .cmdPrintReport_raised_focus_dots_img.Visible = False
6660      .cmdPrintReport_sunken_focus_dots_img.Visible = False
6670      .cmdPrintReport_raised_img_dis.Visible = False
6680      blnPrintReport_Focus = False
6690    End With

EXITP:
6700    Exit Sub

ERRH:
6710    Select Case ERR.Number
        Case Else
6720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6730    End Select
6740    Resume EXITP

End Sub

Private Sub cmdChangePassword_Click()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChangePassword_Click"

6810    With Me
6820      lngRecsCur = .frmUser_Sub.Form.RecCnt  ' ** Form Function: frmUser_Sub.
6830      If lngRecsCur > 0& Then
6840        If .frmUser_Sub.Form.secusr_defaultx = True Then
6850          MsgBox "The user, " & .frmUser_Sub.Form.Username & ", cannot have its password changed.", vbExclamation + vbOKOnly, "Invalid Action"
6860        Else
6870          DoCmd.OpenForm "frmUser_Password", , , , , , THIS_NAME
6880          blnPWChange = True
6890        End If
6900      Else
6910        Select Case blnPWChange
            Case True
6920          Beep
6930          MsgBox "Only one user password may be changed at a time." & vbCrLf & _
                "Please close the form and reopen it to change another.", vbInformation + vbOKOnly, "One-At-A-Time"
6940        Case False
6950          Beep
6960        End Select
6970      End If
6980    End With

EXITP:
6990    Exit Sub

ERRH:
7000    Select Case ERR.Number
        Case Else
7010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7020    End Select
7030    Resume EXITP

End Sub

Private Sub cmdChangePassword_KeyDown(KeyCode As Integer, Shift As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdChangePassword_KeyDown"

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
          Case vbKeyTab
7170        intRetVal = 0
7180        With Me
7190          If .cmdAdd.Enabled = True Then
7200            .cmdAdd.SetFocus
7210          ElseIf .cmdDelete.Enabled = True Then
7220            .cmdDelete.SetFocus
7230          Else
7240            .cmdClose.SetFocus
7250          End If
7260        End With
7270      End Select
7280    End If

        ' ** Shift keys.
7290    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7300      Select Case intRetVal
          Case vbKeyTab
7310        With Me
7320          intRetVal = 0
7330          If .cmdPrintReport.Enabled = True Then
7340            .cmdPrintReport.SetFocus
7350          Else
7360            lngRecsCur = .frmUser_Sub.Form.RecCnt  ' ** Form Function: frmUser_Sub.
7370            If lngRecsCur > 0& Then
7380              .frmUser_Sub.SetFocus
7390              .frmUser_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmUser_Sub.
7400              If .frmUser_Sub.Form.Primary_Group.Enabled = True Then
7410                .frmUser_Sub.Form.Primary_Group.SetFocus
7420              Else
7430                .frmUser_Sub.Form.Employee_Name.SetFocus
7440              End If
7450            Else
7460              .cmdClose.SetFocus
7470            End If
7480          End If
7490        End With
7500      End Select
7510    End If

EXITP:
7520    KeyCode = intRetVal
7530    Exit Sub

ERRH:
7540    Select Case ERR.Number
        Case Else
7550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7560    End Select
7570    Resume EXITP

End Sub

Private Sub cmdAdd_Click()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

7610    With Me
7620      If blnAdmin = True Then
7630        DoCmd.OpenForm "frmUser_Add", , , , , acDialog, THIS_NAME
7640        .Requery
7650      End If
7660    End With

EXITP:
7670    Exit Sub

ERRH:
7680    Select Case ERR.Number
        Case Else
7690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7700    End Select
7710    Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

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
          Case vbKeyTab
7870        With Me
7880          intRetVal = 0
7890          If .cmdDelete.Enabled = True Then
7900            .cmdDelete.SetFocus
7910          Else
7920            .cmdClose.SetFocus
7930          End If
7940        End With
7950      End Select
7960    End If

        ' ** Shift keys.
7970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7980      Select Case intRetVal
          Case vbKeyTab
7990        With Me
8000          intRetVal = 0
8010          If .cmdChangePassword.Enabled = True Then
8020            .cmdChangePassword.SetFocus
8030          Else
8040            lngRecsCur = .frmUser_Sub.Form.RecCnt  ' ** Form Function: frmUser_Sub.
8050            If lngRecsCur > 0& Then
8060              .frmUser_Sub.SetFocus
8070              .frmUser_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmUser_Sub.
8080              If .frmUser_Sub.Form.Primary_Group.Enabled = True Then
8090                .frmUser_Sub.Form.Primary_Group.SetFocus
8100              Else
8110                .frmUser_Sub.Form.Employee_Name.SetFocus
8120              End If
8130            Else
8140              .cmdClose.SetFocus
8150            End If
8160          End If
8170        End With
8180      End Select
8190    End If

EXITP:
8200    KeyCode = intRetVal
8210    Exit Sub

ERRH:
8220    Select Case ERR.Number
        Case Else
8230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8240    End Select
8250    Resume EXITP

End Sub

Public Sub cmdDelete_Click()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim grp As DAO.Group, usr2 As DAO.User
        Dim strUsers As String, strUsername As String, strUserPIN As String
        Dim strSpecificGroup As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnError As Boolean
        Dim lngRecs As Long, lngXUsrID As Long
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long

8310    With Me
8320      If blnAdmin = True Then
8330        If .frmUser_Sub.Form.secusr_defaultx = True Then
8340          Beep
8350          MsgBox "The user, " & .frmUser_Sub.Form.Username & ", cannot be deleted.", vbExclamation + vbOKOnly, "Invalid Action"
8360        Else
8370          If .frmUser_Sub.Form.Username = .ThisUser Then
8380            Beep
8390            MsgBox "You cannot delete yourself while logged in as yourself!", vbExclamation + vbOKOnly, "Invalid Action"
8400          Else

8410            DoCmd.Hourglass True
8420            blnError = False
8430            strUsers = "Users"

                ' ** Employee Username.
8440            strUsername = .frmUser_Sub.Form.Username
                ' ** Employee Security File PID number.
8450            strUserPIN = .frmUser_Sub.Form.Username
                ' ** Specific Security Group User Belongs to.
8460            strSpecificGroup = .frmUser_Sub.Form.Primary_Group

8470            msgResponse = MsgBox("You are about to delete the user " & strUsername & " from the " & _
                  strSpecificGroup & " group." & vbCrLf & vbCrLf & "Do you wish to continue?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Delete User")
8480            If msgResponse = vbYes Then

                  ' ** Delete the current account where name = strUserName from the Access Security.
8490              For Each grp In DBEngine.Workspaces(0).Groups
8500                For Each usr2 In grp.Users
8510                  If usr2.Name = strUsername Then
8520                    grp.Users.Delete usr2.Name
8530                    grp.Users.Refresh
8540                    Exit For
8550                  End If
8560                Next
8570              Next

                  ' ** First delete the user after we make sure it exists in the MDW file.
8580              For Each usr2 In DBEngine.Workspaces(0).Users
8590                If usr2.Name = strUsername Then
8600                  DBEngine.Workspaces(0).Users.Delete strUsername
8610                  Exit For
8620                End If
8630              Next

8640              Set dbs = CurrentDb

                  ' ** And after finally, delete the user from the _~xusr table.
8650              lngXUsrID = 0&
8660              Set rst = dbs.OpenRecordset("_~xusr", dbOpenDynaset, dbConsistent)
8670              With rst
8680                If .BOF = True And .EOF = True Then
                      ' ** Shouldn't happen!
8690                Else
8700                  .MoveLast
8710                  lngRecs = .RecordCount
8720                  .MoveFirst
8730                  For lngX = 1& To lngRecs
8740                    strTmp01 = FilterGUIDString(StringFromGUID(![s_GUID]))  ' ** Module Utility: modCodeUtilities.
8750                    strTmp02 = FilterGUIDString(Me.frmUser_Sub.Form.s_GUID_string)  ' ** Module Utility: modCodeUtilities.
8760                    If strTmp01 = strTmp02 Then  ' ** Module Utility: modCodeUtilities.
                          ' ** The StringFromGUID function converts a GUID, which is an array of type Byte, to a string.
8770                      lngXUsrID = ![xusr_id]
8780                      Exit For
8790                    End If
8800                    If lngX < lngRecs Then .MoveNext
8810                  Next
8820                End If
8830                .Close
8840              End With

                  ' ** Delete user from _~xusr, by specified [xusr].
8850              If lngXUsrID > 0& Then
8860                Set qdf = dbs.QueryDefs("qryXUsr_01")
8870                With qdf.Parameters
8880                  ![xusr] = lngXUsrID
8890                End With
8900                qdf.Execute dbFailOnError
8910              End If

                  ' ** Then delete the user from the Users table.
8920              Set qdf = dbs.QueryDefs("qryUserMaint_02")
8930              With qdf.Parameters
8940                ![usr] = strUsername
8950              End With
8960              qdf.Execute

                  ' ** Delete from tblSecurity_User by specified [usr].
8970              Set qdf = dbs.QueryDefs("qryUserMaint_03")
8980              With qdf.Parameters
8990                ![usr] = strUsername
9000              End With
9010              qdf.Execute

9020              dbs.Close

9030              .frmUser_Sub.Form.Requery

9040            End If

9050          End If
9060        End If
9070      Else
9080        MsgBox "You do not have sufficient rights to delete users.", vbInformation + vbOKOnly, gstrNoPermission
9090      End If
9100    End With

EXITP:
9110    Set rst = Nothing
9120    Set qdf = Nothing
9130    Set dbs = Nothing
9140    DoCmd.Hourglass False
9150    Me.Requery
9160    If blnError = True Then
9170      MsgBox "User has not been deleted due to an error.", vbCritical + vbOKOnly, "Error"
9180    End If
9190    Exit Sub

ERRH:
9200    Beep
9210    Select Case ERR.Number
        Case 3265 ' ** Item not found in this collection.
9220      Beep
9230      MsgBox "Operation cannot be completed because the user information has been changed." & vbCrLf & vbCrLf & _
            "This may be due to reinstallation of the program or restoring a previous copy of the data.", _
            vbCritical + vbOKOnly, "Delete User"
9240      blnError = False  ' ** Don't show error message for this.
9250    Case Else
9260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9270      blnError = True
9280    End Select
9290    Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9320    intShiftDown = (Shift And acShiftMask) > 0
9330    intAltDown = (Shift And acAltMask) > 0
9340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9360      Select Case intRetVal
          Case vbKeyTab
9370        With Me
9380          intRetVal = 0
9390          .cmdClose.SetFocus
9400        End With
9410      End Select
9420    End If

        ' ** Shift keys.
9430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9440      Select Case intRetVal
          Case vbKeyTab
9450        With Me
9460          intRetVal = 0
9470          If .cmdAdd.Enabled = True Then
9480            .cmdAdd.SetFocus
9490          ElseIf .cmdChangePassword.Enabled = True Then
9500            .cmdChangePassword.SetFocus
9510          Else
9520            lngRecsCur = .frmUser_Sub.Form.RecCnt  ' ** Form Function: frmUser_Sub.
9530            If lngRecsCur > 0& Then
9540              .frmUser_Sub.SetFocus
9550              .frmUser_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmUser_Sub.
9560              If .frmUser_Sub.Form.Primary_Group.Enabled = True Then
9570                .frmUser_Sub.Form.Primary_Group.SetFocus
9580              Else
9590                .frmUser_Sub.Form.Employee_Name.SetFocus
9600              End If
9610            Else
9620              .cmdClose.SetFocus
9630            End If
9640          End If
9650        End With
9660      End Select
9670    End If

EXITP:
9680    KeyCode = intRetVal
9690    Exit Sub

ERRH:
9700    Select Case ERR.Number
        Case Else
9710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9720    End Select
9730    Resume EXITP

End Sub

Private Sub ThisUser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisUser_MouseMove"

9810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9820    Exit Sub

ERRH:
9830    Select Case ERR.Number
        Case Else
9840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9850    End Select
9860    Resume EXITP

End Sub

Private Sub cmdHelp_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_box_MouseMove"

9910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9920    Exit Sub

ERRH:
9930    Select Case ERR.Number
        Case Else
9940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9950    End Select
9960    Resume EXITP

End Sub

Private Sub cmdHelp_lbl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_lbl1_MouseMove"

10010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10020   Exit Sub

ERRH:
10030   Select Case ERR.Number
        Case Else
10040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10050   End Select
10060   Resume EXITP

End Sub

Private Sub cmdHelp_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_lbl2_MouseMove"

10110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10120   Exit Sub

ERRH:
10130   Select Case ERR.Number
        Case Else
10140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10150   End Select
10160   Resume EXITP

End Sub

Private Sub cmdHelp_lbl3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_lbl3_MouseMove"

10210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
10220   Exit Sub

ERRH:
10230   Select Case ERR.Number
        Case Else
10240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10250   End Select
10260   Resume EXITP

End Sub

Private Sub cmdHelp_GotFocus()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_GotFocus"

10310   With Me
10320     If blnAdmin = True And (.ThisUser = "TAAdmin" Or .ThisUser = "TADemo") Then
10330       blnHelp_Focus = True
10340       Select Case .ThisUser
            Case "TAAdmin"
10350         Select Case .cmdHelp_lbl1.Visible
              Case True
10360           .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = True
10370           .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
10380           .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
10390           .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
10400         Case False
10410           .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = True
10420           .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
10430           .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
10440           .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
10450         End Select
10460       Case "TADemo"
10470         Select Case .cmdHelp_lbl1.Visible
              Case True
10480           .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = True
10490           .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
10500           .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
10510           .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
10520         Case False
10530           .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = True
10540           .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
10550           .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
10560           .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
10570         End Select
10580       End Select
10590       .cmdHelp_TAAdmin_R_raised_img.Visible = False
10600       .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
10610       .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
10620       .cmdHelp_TAAdmin_R_sunken_focus_dots_img.Visible = False
10630       .cmdHelp_TAAdmin_L_raised_img.Visible = False
10640       .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
10650       .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
10660       .cmdHelp_TAAdmin_L_sunken_focus_dots_img.Visible = False
10670       .cmdHelp_TADemo_R_raised_img.Visible = False
10680       .cmdHelp_TADemo_R_raised_focus_img.Visible = False
10690       .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
10700       .cmdHelp_TADemo_R_sunken_focus_dots_img.Visible = False
10710       .cmdHelp_TADemo_L_raised_img.Visible = False
10720       .cmdHelp_TADemo_L_raised_focus_img.Visible = False
10730       .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
10740       .cmdHelp_TADemo_L_sunken_focus_dots_img.Visible = False
10750     End If
10760   End With

EXITP:
10770   Exit Sub

ERRH:
10780   Select Case ERR.Number
        Case Else
10790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10800   End Select
10810   Resume EXITP

End Sub

Private Sub cmdHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseDown"

10910   With Me
10920     If blnAdmin = True And (.ThisUser = "TAAdmin" Or .ThisUser = "TADemo") Then
10930       blnHelp_MouseDown = True
10940       Select Case .ThisUser
            Case "TAAdmin"
10950         Select Case .cmdHelp_lbl1.Visible
              Case True
10960           .cmdHelp_TAAdmin_L_sunken_focus_dots_img.Visible = True
10970           .cmdHelp_TAAdmin_R_sunken_focus_dots_img.Visible = False
10980           .cmdHelp_TADemo_R_sunken_focus_dots_img.Visible = False
10990           .cmdHelp_TADemo_L_sunken_focus_dots_img.Visible = False
11000         Case False
11010           .cmdHelp_TAAdmin_R_sunken_focus_dots_img.Visible = True
11020           .cmdHelp_TAAdmin_L_sunken_focus_dots_img.Visible = False
11030           .cmdHelp_TADemo_R_sunken_focus_dots_img.Visible = False
11040           .cmdHelp_TADemo_L_sunken_focus_dots_img.Visible = False
11050         End Select
11060       Case "TADemo"
11070         Select Case .cmdHelp_lbl1.Visible
              Case True
11080           .cmdHelp_TADemo_L_sunken_focus_dots_img.Visible = True
11090           .cmdHelp_TADemo_R_sunken_focus_dots_img.Visible = False
11100           .cmdHelp_TAAdmin_R_sunken_focus_dots_img.Visible = False
11110           .cmdHelp_TAAdmin_L_sunken_focus_dots_img.Visible = False
11120         Case False
11130           .cmdHelp_TADemo_R_sunken_focus_dots_img.Visible = True
11140           .cmdHelp_TADemo_L_sunken_focus_dots_img.Visible = False
11150           .cmdHelp_TAAdmin_R_sunken_focus_dots_img.Visible = False
11160           .cmdHelp_TAAdmin_L_sunken_focus_dots_img.Visible = False
11170         End Select
11180       End Select
11190       .cmdHelp_TAAdmin_R_raised_img.Visible = False
11200       .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
11210       .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
11220       .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
11230       .cmdHelp_TAAdmin_L_raised_img.Visible = False
11240       .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
11250       .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
11260       .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
11270       .cmdHelp_TADemo_R_raised_img.Visible = False
11280       .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
11290       .cmdHelp_TADemo_R_raised_focus_img.Visible = False
11300       .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
11310       .cmdHelp_TADemo_L_raised_img.Visible = False
11320       .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
11330       .cmdHelp_TADemo_L_raised_focus_img.Visible = False
11340       .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
11350     End If
11360   End With

EXITP:
11370   Exit Sub

ERRH:
11380   Select Case ERR.Number
        Case Else
11390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11400   End Select
11410   Resume EXITP

End Sub

Private Sub cmdHelp_Click()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_Click"

11510   With Me
11520     Select Case .cmdHelp_lbl1.Visible
          Case True
            ' ** Hide the Help labels.
11530       .cmdHelp_lbl1.Visible = False
11540       .cmdHelp_lbl2.Visible = False
11550       .cmdHelp_lbl3.Visible = False
11560       .cmdHelp_box.Visible = False
11570     Case False
            ' ** Show the Help labels.
11580       .cmdHelp_box.Visible = True
11590       .cmdHelp_lbl1.Visible = True
11600       .cmdHelp_lbl2.Visible = True
11610       .cmdHelp_lbl3.Visible = True
11620     End Select
11630     .frmUser_Sub.SetFocus
11640   End With

EXITP:
11650   Exit Sub

ERRH:
11660   Select Case ERR.Number
        Case Else
11670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11680   End Select
11690   Resume EXITP

End Sub

Private Sub cmdHelp_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11720   intShiftDown = (Shift And acShiftMask) > 0
11730   intAltDown = (Shift And acAltMask) > 0
11740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab
11770       With Me
11780         intRetVal = 0
11790         lngRecsCur = .frmUser_Sub.Form.RecCnt  ' ** Form Function: frmUser_Sub.
11800         If lngRecsCur > 0& Then
11810           .frmUser_Sub.SetFocus
11820           .frmUser_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmUser_Sub.
11830           .frmUser_Sub.Form.Username.SetFocus
11840         Else
11850           If .cmdChangePassword.Enabled = True Then
11860             .cmdChangePassword.SetFocus
11870           ElseIf .cmdAdd.Enabled = True Then
11880             .cmdAdd.SetFocus
11890           ElseIf .cmdDelete.Enabled = True Then  ' ** Moot!
11900             .cmdDelete.SetFocus
11910           Else
11920             .cmdClose.SetFocus
11930           End If
11940         End If
11950       End With
11960     End Select
11970   End If

        ' ** Shift keys.
11980   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11990     Select Case intRetVal
          Case vbKeyTab
12000       With Me
12010         intRetVal = 0
12020         .cmdClose.SetFocus
12030       End With
12040     End Select
12050   End If

EXITP:
12060   KeyCode = intRetVal
12070   Exit Sub

ERRH:
12080   Select Case ERR.Number
        Case Else
12090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12100   End Select
12110   Resume EXITP

End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseMove"

12210   With Me
12220     If blnAdmin = True And (.ThisUser = "TAAdmin" Or .ThisUser = "TADemo") Then
12230       If blnHelp_MouseDown = False Then
12240         Select Case .ThisUser
              Case "TAAdmin"
12250           Select Case .cmdHelp_lbl1.Visible
                Case True
12260             Select Case blnHelp_Focus
                  Case True
12270               .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = True
12280               .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
12290               .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
12300               .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
12310               .cmdHelp_TADemo_R_raised_focus_img.Visible = False
12320               .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
12330               .cmdHelp_TADemo_L_raised_focus_img.Visible = False
12340               .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
12350             Case False
12360               .cmdHelp_TAAdmin_L_raised_focus_img.Visible = True
12370               .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
12380               .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
12390               .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
12400               .cmdHelp_TADemo_R_raised_focus_img.Visible = False
12410               .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
12420               .cmdHelp_TADemo_L_raised_focus_img.Visible = False
12430               .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
12440             End Select
12450           Case False
12460             Select Case blnHelp_Focus
                  Case True
12470               .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = True
12480               .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
12490               .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
12500               .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
12510               .cmdHelp_TADemo_R_raised_focus_img.Visible = False
12520               .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
12530               .cmdHelp_TADemo_L_raised_focus_img.Visible = False
12540               .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
12550             Case False
12560               .cmdHelp_TAAdmin_R_raised_focus_img.Visible = True
12570               .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
12580               .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
12590               .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
12600               .cmdHelp_TADemo_R_raised_focus_img.Visible = False
12610               .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
12620               .cmdHelp_TADemo_L_raised_focus_img.Visible = False
12630               .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
12640             End Select
12650           End Select
12660         Case "TADemo"
12670           Select Case .cmdHelp_lbl1.Visible
                Case True
12680             Select Case blnHelp_Focus
                  Case True
12690               .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = True
12700               .cmdHelp_TADemo_L_raised_focus_img.Visible = False
12710               .cmdHelp_TADemo_R_raised_focus_img.Visible = False
12720               .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
12730               .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
12740               .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
12750               .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
12760               .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
12770             Case False
12780               .cmdHelp_TADemo_L_raised_focus_img.Visible = True
12790               .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
12800               .cmdHelp_TADemo_R_raised_focus_img.Visible = False
12810               .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
12820               .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
12830               .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
12840               .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
12850               .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
12860             End Select
12870           Case False
12880             Select Case blnHelp_Focus
                  Case True
12890               .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = True
12900               .cmdHelp_TADemo_R_raised_focus_img.Visible = False
12910               .cmdHelp_TADemo_L_raised_focus_img.Visible = False
12920               .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
12930               .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
12940               .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
12950               .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
12960               .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
12970             Case False
12980               .cmdHelp_TADemo_R_raised_focus_img.Visible = True
12990               .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
13000               .cmdHelp_TADemo_L_raised_focus_img.Visible = False
13010               .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
13020               .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
13030               .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
13040               .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
13050               .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
13060             End Select
13070           End Select
13080         End Select
13090         .cmdHelp_TAAdmin_R_raised_img.Visible = False
13100         .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
13110         .cmdHelp_TAAdmin_R_sunken_focus_dots_img.Visible = False
13120         .cmdHelp_TAAdmin_L_raised_img.Visible = False
13130         .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
13140         .cmdHelp_TAAdmin_L_sunken_focus_dots_img.Visible = False
13150         .cmdHelp_TADemo_R_raised_img.Visible = False
13160         .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
13170         .cmdHelp_TADemo_R_sunken_focus_dots_img.Visible = False
13180         .cmdHelp_TADemo_L_raised_img.Visible = False
13190         .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
13200         .cmdHelp_TADemo_L_sunken_focus_dots_img.Visible = False
13210       End If
13220     End If
13230   End With

EXITP:
13240   Exit Sub

ERRH:
13250   Select Case ERR.Number
        Case Else
13260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13270   End Select
13280   Resume EXITP

End Sub

Private Sub cmdHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseUp"

13310   With Me
13320     If blnAdmin = True And (.ThisUser = "TAAdmin" Or .ThisUser = "TADemo") Then
13330       Select Case .ThisUser
            Case "TAAdmin"
13340         Select Case .cmdHelp_lbl1.Visible
              Case True
13350           .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = True
13360           .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
13370           .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
13380           .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
13390         Case False
13400           .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = True
13410           .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
13420           .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
13430           .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
13440         End Select
13450       Case "TADemo"
13460         Select Case .cmdHelp_lbl1.Visible
              Case True
13470           .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = True
13480           .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
13490           .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
13500           .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
13510         Case False
13520           .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = True
13530           .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
13540           .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
13550           .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
13560         End Select
13570       End Select
13580       .cmdHelp_TAAdmin_R_raised_img.Visible = False
13590       .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
13600       .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
13610       .cmdHelp_TAAdmin_R_sunken_focus_dots_img.Visible = False
13620       .cmdHelp_TAAdmin_L_raised_img.Visible = False
13630       .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
13640       .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
13650       .cmdHelp_TAAdmin_L_sunken_focus_dots_img.Visible = False
13660       .cmdHelp_TADemo_R_raised_img.Visible = False
13670       .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
13680       .cmdHelp_TADemo_R_raised_focus_img.Visible = False
13690       .cmdHelp_TADemo_R_sunken_focus_dots_img.Visible = False
13700       .cmdHelp_TADemo_L_raised_img.Visible = False
13710       .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
13720       .cmdHelp_TADemo_L_raised_focus_img.Visible = False
13730       .cmdHelp_TADemo_L_sunken_focus_dots_img.Visible = False
13740       blnHelp_MouseDown = False
13750     End If
13760   End With

EXITP:
13770   Exit Sub

ERRH:
13780   Select Case ERR.Number
        Case Else
13790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13800   End Select
13810   Resume EXITP

End Sub

Private Sub cmdHelp_LostFocus()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_LostFocus"

13910   With Me
13920     If blnAdmin = True And (.ThisUser = "TAAdmin" Or .ThisUser = "TADemo") Then
13930       Select Case .ThisUser
            Case "TAAdmin"
13940         Select Case .cmdHelp_lbl1.Visible
              Case True
13950           .cmdHelp_TAAdmin_L_raised_img.Visible = True
13960           .cmdHelp_TAAdmin_R_raised_img.Visible = False
13970           .cmdHelp_TADemo_L_raised_img.Visible = False
13980           .cmdHelp_TADemo_R_raised_img.Visible = False
13990         Case False
14000           .cmdHelp_TAAdmin_R_raised_img.Visible = True
14010           .cmdHelp_TAAdmin_L_raised_img.Visible = False
14020           .cmdHelp_TADemo_L_raised_img.Visible = False
14030           .cmdHelp_TADemo_R_raised_img.Visible = False
14040         End Select
14050       Case "TADemo"
14060         Select Case .cmdHelp_lbl1.Visible
              Case True
14070           .cmdHelp_TADemo_L_raised_img.Visible = True
14080           .cmdHelp_TADemo_R_raised_img.Visible = False
14090           .cmdHelp_TAAdmin_R_raised_img.Visible = False
14100           .cmdHelp_TAAdmin_L_raised_img.Visible = False
14110         Case False
14120           .cmdHelp_TADemo_R_raised_img.Visible = True
14130           .cmdHelp_TADemo_L_raised_img.Visible = False
14140           .cmdHelp_TAAdmin_R_raised_img.Visible = False
14150           .cmdHelp_TAAdmin_L_raised_img.Visible = False
14160         End Select
14170       End Select
14180       .cmdHelp_TAAdmin_R_raised_semifocus_dots_img.Visible = False
14190       .cmdHelp_TAAdmin_R_raised_focus_img.Visible = False
14200       .cmdHelp_TAAdmin_R_raised_focus_dots_img.Visible = False
14210       .cmdHelp_TAAdmin_R_sunken_focus_dots_img.Visible = False
14220       .cmdHelp_TAAdmin_L_raised_semifocus_dots_img.Visible = False
14230       .cmdHelp_TAAdmin_L_raised_focus_img.Visible = False
14240       .cmdHelp_TAAdmin_L_raised_focus_dots_img.Visible = False
14250       .cmdHelp_TAAdmin_L_sunken_focus_dots_img.Visible = False
14260       .cmdHelp_TADemo_R_raised_semifocus_dots_img.Visible = False
14270       .cmdHelp_TADemo_R_raised_focus_img.Visible = False
14280       .cmdHelp_TADemo_R_raised_focus_dots_img.Visible = False
14290       .cmdHelp_TADemo_R_sunken_focus_dots_img.Visible = False
14300       .cmdHelp_TADemo_L_raised_semifocus_dots_img.Visible = False
14310       .cmdHelp_TADemo_L_raised_focus_img.Visible = False
14320       .cmdHelp_TADemo_L_raised_focus_dots_img.Visible = False
14330       .cmdHelp_TADemo_L_sunken_focus_dots_img.Visible = False
14340       blnHelp_Focus = False
14350     End If
14360   End With

EXITP:
14370   Exit Sub

ERRH:
14380   Select Case ERR.Number
        Case Else
14390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14400   End Select
14410   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

14510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
14520   Exit Sub

ERRH:
14530   Select Case ERR.Number
        Case Else
14540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14550   End Select
14560   Resume EXITP

End Sub

Private Sub cmdShowAll_Click()
' ** This button toggles the filtering on and off.

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShowAll_Click"

        Dim strFilter As String

14610   With Me
14620     If .frmUser_Sub.Form.Filter = vbNullString Then
14630       strFilter = FILTER_DEF
14640       .frmUser_Sub.Form.Filter = strFilter
14650       .frmUser_Sub.Form.FilterOn = True
14660     Else
14670       .frmUser_Sub.Form.Filter = vbNullString
14680     End If
14690   End With

EXITP:
14700   Exit Sub

ERRH:
14710   Select Case ERR.Number
        Case Else
14720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14730   End Select
14740   Resume EXITP

End Sub

Private Function DefaultUserHelp() As Boolean

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "DefaultUserHelp"

        Dim strUser As String

14810   With Me
14820     strUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
14830     .cmdHelp_TAAdmin_R_raised_img.Visible = False
14840     .cmdHelp_TADemo_R_raised_img.Visible = False
14850     .cmdHelp_TAAdmin_L_raised_img.Visible = False
14860     .cmdHelp_TADemo_L_raised_img.Visible = False
14870     Select Case strUser
          Case "TAAdmin", "TADemo"
14880       .cmdHelp.Visible = True
14890       .cmdHelp.Enabled = True
14900       Select Case strUser
            Case "TAAdmin"
14910         .cmdHelp_TAAdmin_R_raised_img.Visible = True
14920         .cmdHelp_lbl1.Caption = "TAAdmin is the default user name of"  ' Trust Accountant licensee."
14930         If .cmdHelp_lbl2.Caption <> "the Trust Accountant licensee. It cannot" Then
14940           .cmdHelp_lbl2.Caption = "the Trust Accountant licensee. It cannot"
14950         End If
14960       Case "TADemo"
14970         .cmdHelp_TADemo_R_raised_img.Visible = True
14980         .cmdHelp_lbl1.Caption = "TADemo is the default user name of"  ' Trust Demo licensee."
14990         If .cmdHelp_lbl2.Caption <> "the Trust Demo licensee. It cannot" Then
15000           .cmdHelp_lbl2.Caption = "the Trust Demo licensee. It cannot"
15010         End If
15020       End Select
15030     Case Else
            ' ** The others aren't restricted.
15040       .cmdHelp.Visible = False
15050       .cmdHelp.Enabled = False
15060     End Select
15070   End With

EXITP:
15080   Exit Function

ERRH:
15090   Select Case ERR.Number
        Case Else
15100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15110   End Select
15120   Resume EXITP

End Function
