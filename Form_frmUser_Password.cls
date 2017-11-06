VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUser_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmUser_Password"

'VGC 03/17/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   cmdHelp_Click()
' #######################################

'xusr_extant     : Current password
'xusr_antecedent : Previous password  extinct
'xusr_origin     : Password date

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Password Help:    H {cmdHelp}
' **   OK:               O {cmdOK}
' **   User Name:        U {Username}

Private blnHelp_Focus As Boolean, blnHelp_MouseDown As Boolean
Private lngHelpBox_Top As Long, lngHelpbox_Height As Long
Private lngHelpLbl_Top As Long, lngHelpLbl_Height As Long, lngFooterHgtLine_Top As Long, lngHelp_OFFSET As Long

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long

Private blnClearedPW As Boolean
Private strHelpCap As String
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long
        Dim intX As Integer

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
160           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

170           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

180           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

190           blnClearedPW = False
200           blnHelp_Focus = False: blnHelp_MouseDown = False

210           strHelpCap = .cmdHelp_lbl.Caption
220           lngHelpBox_Top = .cmdHelp_box.Top
230           lngHelpLbl_Top = .cmdHelp_lbl.Top
240           lngHelpbox_Height = .cmdHelp_box.Height
250           lngHelpLbl_Height = .cmdHelp_lbl.Height
260           lngFooterHgtLine_Top = .form_footer_height_line.Top
270           lngHelp_OFFSET = (.FormFooter.Height - .Footer_vline03.Top)

280           .Username = Forms(strCallingForm).frmUser_Sub.Form.Username
290           .Employee_Name = Forms(strCallingForm).frmUser_Sub.Form.Employee_Name
300           If (.Username = CurrentUser) Then  ' ** Internal Access Function: Trust Accountant login.
310             .OldPassword.Enabled = True  ' ** Enable old PW button for current user.
320             .OldPassword.SetFocus
330           Else
340             .OldPassword.Enabled = False  ' ** Admins may clear password, but will still have to enter a new one.
350             .OldPassword.BorderColor = WIN_CLR_DISR
360             .OldPassword_lbl.BackStyle = acBackStyleTransparent
370           End If
              ' ** The StringFromGUID function converts a GUID, which is an array of type Byte, to a string.
380           .s_GUID = FilterGUIDString(StringFromGUID(Forms(strCallingForm).frmUser_Sub.Form.s_GUID))  ' ** Module Function: modCodeUtilities.

390           lngTmp03 = -2&: strTmp01 = vbNullString: strTmp02 = vbNullString
400           lngTmp03 = Security_PW_Exp_Get("Cycle")  ' ** Module Function: modSecurityFuncs.
410           strTmp01 = .cmdHelp_lbl.Caption
420           Select Case lngTmp03
              Case 28&
430             strTmp02 = "Password must be changed every 4 weeks"
440           Case 30&
450             strTmp02 = "Password must be changed every 30 days"
460           Case 365&
470             strTmp02 = "Password must be changed every 12 months"
480           Case 0&
490             strTmp02 = "Password does not expire"
500           Case -1&  ' ** Not found.
510             strTmp02 = "NOT FOUND!"
520           Case -9&  ' ** Error.
530             strTmp02 = RET_ERR
540           Case Else
550             strTmp02 = "Password must be changed every " & CStr(lngTmp03) & " days"
560           End Select

570           intPos01 = InStr(strTmp01, "Password must be changed")
580           If intPos01 > 0 Then
590             intPos02 = 0
600             For intX = intPos01 To Len(strTmp01)
610               If Mid(strTmp01, intX, 1) = vbCr Or Mid(strTmp01, intX, 1) = vbLf Then
620                 intPos02 = intX
630                 Exit For
640               End If
650             Next
660             If intPos02 > 0 Then
670               .cmdHelp_lbl.Caption = Left(strTmp01, (intPos01 - 1)) & strTmp02 & Mid(strTmp01, intPos02)
680             End If
690           End If

              ' ** Hide the password help.
700           cmdHelp_Click  ' ** Procedure: Below.
710           DoEvents

720           blnIsOpen = False

730         Else
740           Cancel = -1
750         End If
760       Else
770         Cancel = -1
780       End If
790     End With

800     If Cancel = -1 Then
810       Beep
820       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
830     End If

EXITP:
840     Exit Sub

ERRH:
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
870     End Select
880     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyEscape
970         intRetVal = 0
980         cmdCancel_Click  ' ** Procedure: Below.
990       End Select
1000    End If

EXITP:
1010    KeyCode = intRetVal
1020    Exit Sub

ERRH:
1030    intRetVal = 0
1040    Select Case ERR.Number
        Case Else
1050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1060    End Select
1070    Resume EXITP

End Sub

Private Sub Form_Close()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Close"

1110    With Forms("frmUser")
1120      .frmUser_Sub.Form.Requery
1130      gblnSetFocus = True
1140      .TimerInterval = 100&
1150    End With

EXITP:
1160    Exit Sub

ERRH:
1170    Select Case ERR.Number
        Case Else
1180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1190    End Select
1200    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1310    With Me
1320      If ((.cmdHelp_R_raised_focus_dots_img.Visible = True Or .cmdHelp_R_raised_focus_img.Visible = True) Or _
              (.cmdHelp_L_raised_focus_dots_img.Visible = True Or .cmdHelp_L_raised_focus_img.Visible = True)) Then
1330        Select Case .cmdHelp_box.Visible
            Case True
1340          Select Case blnHelp_Focus
              Case True
1350            .cmdHelp_L_raised_semifocus_dots_img.Visible = True
1360            .cmdHelp_L_raised_img.Visible = False
1370            .cmdHelp_R_raised_img.Visible = False
1380            .cmdHelp_R_raised_semifocus_dots_img.Visible = False
1390          Case False
1400            .cmdHelp_L_raised_img.Visible = True
1410            .cmdHelp_L_raised_semifocus_dots_img.Visible = False
1420            .cmdHelp_R_raised_img.Visible = False
1430            .cmdHelp_R_raised_semifocus_dots_img.Visible = False
1440          End Select
1450        Case False
1460          Select Case blnHelp_Focus
              Case True
1470            .cmdHelp_R_raised_semifocus_dots_img.Visible = True
1480            .cmdHelp_R_raised_img.Visible = False
1490            .cmdHelp_L_raised_img.Visible = False
1500            .cmdHelp_L_raised_semifocus_dots_img.Visible = False
1510          Case False
1520            .cmdHelp_R_raised_img.Visible = True
1530            .cmdHelp_R_raised_semifocus_dots_img.Visible = False
1540            .cmdHelp_L_raised_img.Visible = False
1550            .cmdHelp_L_raised_semifocus_dots_img.Visible = False
1560          End Select
1570        End Select
1580        .cmdHelp_R_raised_focus_img.Visible = False
1590        .cmdHelp_R_raised_focus_dots_img.Visible = False
1600        .cmdHelp_R_sunken_focus_dots_img.Visible = False
1610        .cmdHelp_L_raised_focus_img.Visible = False
1620        .cmdHelp_L_raised_focus_dots_img.Visible = False
1630        .cmdHelp_L_sunken_focus_dots_img.Visible = False
1640      End If
1650    End With

EXITP:
1660    Exit Sub

ERRH:
1670    Select Case ERR.Number
        Case Else
1680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1690    End Select
1700    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

1810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, rst As DAO.Recordset
        Dim msgResponse As VbMsgBoxResult
        Dim lngRecs As Long
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnContinue As Boolean

1910    blnContinue = True

1920    With Me

1930      If IIf(IsNull(.NewPassword) = True, vbNullString, .NewPassword) <> _
              IIf(IsNull(.VerifyNewPassword) = True, vbNullString, .VerifyNewPassword) Then
1940        blnContinue = False
1950        MsgBox "The second entry of the password didn't match the first." & vbCrLf & vbCrLf & _
              "Please re-enter the password so that both match.", vbExclamation + vbOKOnly, "No Match"
1960      Else
1970        Set wrk = DBEngine.Workspaces(0)
1980        If IsNull(.NewPassword) = True Or .NewPassword = vbNullString Then
1990          blnContinue = False
2000          msgResponse = MsgBox("Would you like to clear the password?", vbQuestion + vbYesNo + vbDefaultButton1, "Change Password")
2010          If msgResponse = vbYes Then
2020            wrk.Users(.Username).NewPassword vbNullString, vbNullString
2030            Beep
2040            MsgBox "Password change successful.", vbInformation + vbOKOnly, "Password Change Successful"
2050            blnClearedPW = True
2060            MsgBox "Please create a new password." & vbCrLf & vbCrLf & "All users must have a password, maximum 14 characters.", _
                  vbInformation + vbOKOnly, "Password Required"
2070          End If
2080        Else
2090          If Pass_Check(.NewPassword, .Username) = False Then  ' ** Module Function: modCodeUtilities.
                ' ** Failure message shown in Pass_Check() of modCodeUtilities.
2100            blnContinue = False
2110          Else
                ' ** Check the _~xusr table to make sure the new password is different.
2120            Set dbs = wrk.Databases(0)
2130            With dbs
2140              Set rst = dbs.OpenRecordset("_~xusr", dbOpenDynaset, dbConsistent)
2150              With rst
2160                If .BOF = True And .EOF = True Then
                      ' ** Shouldn't happen!
2170                  blnContinue = False
2180                Else
2190                  .MoveLast
2200                  lngRecs = .RecordCount
2210                  .MoveFirst
2220                  blnContinue = False
2230                  varTmp00 = StringFromGUID(Me.s_GUID)
2240                  For lngX = 1& To lngRecs
2250                    If FilterGUIDString(StringFromGUID(![s_GUID])) = varTmp00 Then  ' ** Module Function: modCodeUtilities.
2260                      blnContinue = True
2270                      Exit For
2280                    End If
2290                    If lngX < lngRecs Then .MoveNext
2300                  Next
2310                  If blnContinue = True Then
                        ' ** Update the _~xusr table with the new password and date.
2320                    .Edit
2330                    ![xusr_antecedent] = ![xusr_extant]
2340                    ![xusr_extant] = EncodeString(Me.NewPassword)  ' ** Module Function: modCodeUtilities.
2350                    ![xusr_origin] = EncodeString(Format(Date, "mm/dd/yyyy"))  ' ** Module Function: modCodeUtilities.
2360                    ![xusr_user] = Me.Username  ' ** THIS IS THE USER'S NAME, NOT GetUserName!
2370                    ![xusr_datemodified] = Now()
2380                    .Update
2390                  Else
                        ' ** THIS LEAVES EVERYTHING ELSE SUCCESSFUL EXCEPT THE PASSWORD TABLE!
                        'blnContinue = False
2400                    Beep
2410                    MsgBox "The user's password history was not found." & vbCrLf & vbCrLf & _
                          "Contact Deleta Data, Inc., for assistance.", vbCritical + vbOKOnly, "History Not Found"
2420                  End If
2430                End If
2440                .Close
2450              End With
2460              Set rst = Nothing
2470              .Close
2480            End With
2490            Set dbs = Nothing
2500            If blnContinue = True Then
2510              If .OldPassword.Enabled = True And (IsNull(.OldPassword) = False) Then
2520                wrk.Users(.Username).NewPassword .OldPassword, .NewPassword
2530              Else
2540                wrk.Users(.Username).NewPassword vbNullString, .NewPassword
2550                blnClearedPW = False
2560              End If
2570            End If
2580          End If
2590        End If
2600        wrk.Close
2610        Set wrk = Nothing
2620      End If

2630    End With

2640    If blnContinue = True Then
2650      Forms(strCallingForm).frmUser_Sub.Form.Requery
2660      DoEvents
2670      Beep
2680      MsgBox "Password change successful.", vbInformation + vbOKOnly, "Password Change Successful"
2690      DoCmd.Close acForm, THIS_NAME
2700    End If

EXITP:
2710    Set rst = Nothing
2720    Set dbs = Nothing
2730    Set wrk = Nothing
2740    Exit Sub

ERRH:
2750    Select Case ERR.Number
        Case 3033  ' ** You do not have the necessary permissions to use the '|' object
          ' ** Bad old password.
2760      Beep
2770      MsgBox "The old password entered is not valid.", vbInformation + vbOKOnly, "Password Not Changed"
2780    Case 3265 ' ** Item not found in this collection.
2790      Beep
2800      MsgBox "Operation cannot be completed because the user information has been changed." & vbCrLf & vbCrLf & _
            "This may be due to reinstallation of the program or restoring a previous copy of the data.", _
            vbCritical + vbOKOnly, "Password Maintenance"
2810    Case Else
2820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2830    End Select
2840    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2920    intShiftDown = (Shift And acShiftMask) > 0
2930    intAltDown = (Shift And acAltMask) > 0
2940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2960      Select Case intRetVal
          Case vbKeyTab
2970        With Me
2980          intRetVal = 0
2990          .cmdCancel.SetFocus
3000        End With
3010      End Select
3020    End If

        ' ** Shift keys.
3030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3040      Select Case intRetVal
          Case vbKeyTab
3050        With Me
3060          intRetVal = 0
3070          .VerifyNewPassword.SetFocus
3080        End With
3090      End Select
3100    End If

EXITP:
3110    KeyCode = intRetVal
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

3210    If blnClearedPW = True Then
3220      MsgBox "The old password has been cleared." & vbCrLf & vbCrLf & _
            "All users must have a password, maximum 14 characters.", vbInformation + vbOKOnly, "Password Required"
3230    Else
3240      DoCmd.Close acForm, THIS_NAME
3250    End If

EXITP:
3260    Exit Sub

ERRH:
3270    Select Case ERR.Number
        Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3420    intShiftDown = (Shift And acShiftMask) > 0
3430    intAltDown = (Shift And acAltMask) > 0
3440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3460      Select Case intRetVal
          Case vbKeyTab
3470        With Me
3480          intRetVal = 0
3490          .OldPassword.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab
3550        With Me
3560          intRetVal = 0
3570          .cmdOK.SetFocus
3580        End With
3590      End Select
3600    End If

EXITP:
3610    KeyCode = intRetVal
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub cmdHelp_GotFocus()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_GotFocus"

3710    With Me
3720      blnHelp_Focus = True
3730      Select Case .cmdHelp_box.Visible
          Case True
3740        .cmdHelp_L_raised_semifocus_dots_img.Visible = True
3750        .cmdHelp_R_raised_semifocus_dots_img.Visible = False
3760      Case False
3770        .cmdHelp_R_raised_semifocus_dots_img.Visible = True
3780        .cmdHelp_L_raised_semifocus_dots_img.Visible = False
3790      End Select
3800      .cmdHelp_R_raised_img.Visible = False
3810      .cmdHelp_R_raised_focus_img.Visible = False
3820      .cmdHelp_R_raised_focus_dots_img.Visible = False
3830      .cmdHelp_R_sunken_focus_dots_img.Visible = False
3840      .cmdHelp_L_raised_img.Visible = False
3850      .cmdHelp_L_raised_focus_img.Visible = False
3860      .cmdHelp_L_raised_focus_dots_img.Visible = False
3870      .cmdHelp_L_sunken_focus_dots_img.Visible = False
3880    End With

EXITP:
3890    Exit Sub

ERRH:
3900    Select Case ERR.Number
        Case Else
3910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3920    End Select
3930    Resume EXITP

End Sub

Private Sub cmdHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseDown"

4010    With Me
4020      blnHelp_MouseDown = True
4030      Select Case .cmdHelp_box.Visible
          Case True
4040        .cmdHelp_L_sunken_focus_dots_img.Visible = True
4050        .cmdHelp_R_sunken_focus_dots_img.Visible = False
4060      Case False
4070        .cmdHelp_R_sunken_focus_dots_img.Visible = True
4080        .cmdHelp_L_sunken_focus_dots_img.Visible = False
4090      End Select
4100      .cmdHelp_R_raised_img.Visible = False
4110      .cmdHelp_R_raised_semifocus_dots_img.Visible = False
4120      .cmdHelp_R_raised_focus_img.Visible = False
4130      .cmdHelp_R_raised_focus_dots_img.Visible = False
4140      .cmdHelp_L_raised_img.Visible = False
4150      .cmdHelp_L_raised_semifocus_dots_img.Visible = False
4160      .cmdHelp_L_raised_focus_img.Visible = False
4170      .cmdHelp_L_raised_focus_dots_img.Visible = False
4180    End With

EXITP:
4190    Exit Sub

ERRH:
4200    Select Case ERR.Number
        Case Else
4210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4220    End Select
4230    Resume EXITP

End Sub

Private Sub cmdHelp_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_Click"

        Dim strControl As String
        Dim lngTmp01 As Long

4310    With Me

4320  On Error Resume Next
4330      strControl = Screen.PreviousControl.Name
4340  On Error GoTo ERRH

4350      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
4360      lngMonitorNum = 1&: lngTmp01 = 0&
4370      EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
4380      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

4390      Select Case .cmdHelp_lbl.Visible
          Case True
            ' ** Close the Help box.

            ' ** Shrink and move all the Help stuff out of the way.
4400        .form_footer_height_line.Top = .cmdOK.Top
4410        .cmdHelp_lbl.Visible = False
4420        .cmdHelp_box.Visible = False
4430        .cmdHelp_lbl.Height = 15&
4440        .cmdHelp_box.Height = 15&
4450        .cmdHelp_lbl.Top = .cmdOK.Top
4460        .cmdHelp_box.Top = .cmdOK.Top
4470        .Footer_hline03.Visible = False
4480        .Footer_hline04.Visible = False
4490        .Footer_vline03.Visible = False
4500        .Footer_vline04.Visible = False

            ' ** Get the 2nd set of lines out of the way.
4510        .Footer_hline03.Top = .Footer_hline03.Top - 360&
4520        .Footer_hline04.Top = .Footer_hline04.Top - 360&
4530        .Footer_vline03.Top = .Footer_vline03.Top - 360&
4540        .Footer_vline04.Top = .Footer_vline04.Top - 360&

            ' ** Now close up the footer, and shrink the form.
4550        .FormFooter.Height = .FormFooter.Height - lngHelp_OFFSET
4560        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
4570        DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, (lngFrm_Height - lngHelp_OFFSET)  'lngFrm_Top
4580        If lngMonitorNum > 1& Then
4590          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
4600        End If

4610      Case False
            ' ** Open the Help box.

            ' ** Open up the footer
4620        .FormFooter.Height = .FormFooter.Height + lngHelp_OFFSET

            ' ** Move the 2nd set of lines back to where they belong.
4630        .Footer_hline03.Top = .Footer_hline03.Top + 360&
4640        .Footer_hline04.Top = .Footer_hline04.Top + 360&
4650        .Footer_vline03.Top = .Footer_vline03.Top + 360&
4660        .Footer_vline04.Top = .Footer_vline04.Top + 360&

            ' ** Set up the help stuff.
4670        .form_footer_height_line.Top = lngFooterHgtLine_Top
4680        .cmdHelp_box.Top = lngHelpBox_Top
4690        .cmdHelp_lbl.Top = lngHelpLbl_Top
4700        .cmdHelp_box.Height = lngHelpbox_Height
4710        .cmdHelp_lbl.Height = lngHelpLbl_Height
4720        .cmdHelp_box.Visible = True
4730        .cmdHelp_lbl.Visible = True
4740        .Footer_hline03.Visible = True
4750        .Footer_hline04.Visible = True
4760        .Footer_vline03.Visible = True
4770        .Footer_vline04.Visible = True

            ' ** And lengthen the form.
4780        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
4790        DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, lngFrm_Height  'lngFrm_Top
4800        If lngMonitorNum > 1& Then
4810          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
4820        End If

4830      End Select

4840      If strControl <> vbNullString Then
4850  On Error Resume Next
4860        .Controls(strControl).SetFocus
4870  On Error GoTo ERRH
4880      Else
4890        .cmdOK.SetFocus
4900      End If

4910    End With

EXITP:
4920    Exit Sub

ERRH:
4930    Select Case ERR.Number
        Case Else
4940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4950    End Select
4960    Resume EXITP

End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseMove"

5010    With Me
5020      If blnHelp_MouseDown = False Then
5030        Select Case .cmdHelp_box.Visible
            Case True
5040          Select Case blnHelp_Focus
              Case True
5050            .cmdHelp_L_raised_focus_dots_img.Visible = True
5060            .cmdHelp_L_raised_focus_img.Visible = False
5070            .cmdHelp_R_raised_focus_img.Visible = False
5080            .cmdHelp_R_raised_focus_dots_img.Visible = False
5090          Case False
5100            .cmdHelp_L_raised_focus_img.Visible = True
5110            .cmdHelp_L_raised_focus_dots_img.Visible = False
5120            .cmdHelp_R_raised_focus_img.Visible = False
5130            .cmdHelp_R_raised_focus_dots_img.Visible = False
5140          End Select
5150        Case False
5160          Select Case blnHelp_Focus
              Case True
5170            .cmdHelp_R_raised_focus_dots_img.Visible = True
5180            .cmdHelp_R_raised_focus_img.Visible = False
5190            .cmdHelp_L_raised_focus_img.Visible = False
5200            .cmdHelp_L_raised_focus_dots_img.Visible = False
5210          Case False
5220            .cmdHelp_R_raised_focus_img.Visible = True
5230            .cmdHelp_R_raised_focus_dots_img.Visible = False
5240            .cmdHelp_L_raised_focus_img.Visible = False
5250            .cmdHelp_L_raised_focus_dots_img.Visible = False
5260          End Select
5270        End Select
5280        .cmdHelp_R_raised_img.Visible = False
5290        .cmdHelp_R_raised_semifocus_dots_img.Visible = False
5300        .cmdHelp_R_sunken_focus_dots_img.Visible = False
5310        .cmdHelp_L_raised_img.Visible = False
5320        .cmdHelp_L_raised_semifocus_dots_img.Visible = False
5330        .cmdHelp_L_sunken_focus_dots_img.Visible = False
5340        blnHelp_MouseDown = False
5350      End If
5360    End With

EXITP:
5370    Exit Sub

ERRH:
5380    Select Case ERR.Number
        Case Else
5390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5400    End Select
5410    Resume EXITP

End Sub

Private Sub cmdHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseUp"

5510    With Me
5520      Select Case .cmdHelp_box.Visible
          Case True
5530        .cmdHelp_L_raised_focus_dots_img.Visible = True
5540        .cmdHelp_R_raised_focus_dots_img.Visible = False
5550      Case False
5560        .cmdHelp_R_raised_focus_dots_img.Visible = True
5570        .cmdHelp_L_raised_focus_dots_img.Visible = False
5580      End Select
5590      .cmdHelp_R_raised_img.Visible = False
5600      .cmdHelp_R_raised_semifocus_dots_img.Visible = False
5610      .cmdHelp_R_raised_focus_img.Visible = False
5620      .cmdHelp_R_sunken_focus_dots_img.Visible = False
5630      .cmdHelp_L_raised_img.Visible = False
5640      .cmdHelp_L_raised_semifocus_dots_img.Visible = False
5650      .cmdHelp_L_raised_focus_img.Visible = False
5660      .cmdHelp_L_sunken_focus_dots_img.Visible = False
5670      blnHelp_MouseDown = False
5680    End With

EXITP:
5690    Exit Sub

ERRH:
5700    Select Case ERR.Number
        Case Else
5710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5720    End Select
5730    Resume EXITP

End Sub

Private Sub cmdHelp_LostFocus()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_LostFocus"

5810    With Me
5820      Select Case .cmdHelp_box.Visible
          Case True
5830        .cmdHelp_L_raised_img.Visible = True
5840        .cmdHelp_R_raised_img.Visible = False
5850      Case False
5860        .cmdHelp_R_raised_img.Visible = True
5870        .cmdHelp_L_raised_img.Visible = False
5880      End Select
5890      .cmdHelp_R_raised_semifocus_dots_img.Visible = False
5900      .cmdHelp_R_raised_focus_img.Visible = False
5910      .cmdHelp_R_raised_focus_dots_img.Visible = False
5920      .cmdHelp_R_sunken_focus_dots_img.Visible = False
5930      .cmdHelp_L_raised_semifocus_dots_img.Visible = False
5940      .cmdHelp_L_raised_focus_img.Visible = False
5950      .cmdHelp_L_raised_focus_dots_img.Visible = False
5960      .cmdHelp_L_sunken_focus_dots_img.Visible = False
5970      blnHelp_Focus = False
5980    End With

EXITP:
5990    Exit Sub

ERRH:
6000    Select Case ERR.Number
        Case Else
6010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6020    End Select
6030    Resume EXITP

End Sub

Private Sub cmdHelp_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_box_MouseMove"

6110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6120    Exit Sub

ERRH:
6130    Select Case ERR.Number
        Case Else
6140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6150    End Select
6160    Resume EXITP

End Sub

Private Sub cmdHelp_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_lbl_MouseMove"

6210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6220    Exit Sub

ERRH:
6230    Select Case ERR.Number
        Case Else
6240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6250    End Select
6260    Resume EXITP

End Sub

Private Sub OldPassword_KeyDown(KeyCode As Integer, Shift As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "OldPassword_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6320    intShiftDown = (Shift And acShiftMask) > 0
6330    intAltDown = (Shift And acAltMask) > 0
6340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shift keys.
6350    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6370        With Me
6380          intRetVal = 0
6390          .cmdCancel.SetFocus
6400        End With
6410      End Select
6420    End If

EXITP:
6430    KeyCode = intRetVal
6440    Exit Sub

ERRH:
6450    Select Case ERR.Number
        Case Else
6460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6470    End Select
6480    Resume EXITP

End Sub

Private Sub NewPassword_GotFocus()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "NewPassword_GotFocus"

6510    With Me
6520      .NewPassword.SelStart = 0
6530      .NewPassword.SelLength = Len(Trim(.NewPassword.text))
6540    End With

EXITP:
6550    Exit Sub

ERRH:
6560    Select Case ERR.Number
        Case Else
6570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6580    End Select
6590    Resume EXITP

End Sub

Private Sub VerifyNewPassword_GotFocus()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyNewPassword_GotFocus"

6610    With Me
6620      .VerifyNewPassword.SelStart = 0
6630      .VerifyNewPassword.SelLength = Len(Trim(.VerifyNewPassword.text))
6640    End With

EXITP:
6650    Exit Sub

ERRH:
6660    Select Case ERR.Number
        Case Else
6670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6680    End Select
6690    Resume EXITP

End Sub

Private Sub VerifyNewPassword_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "VerifyNewPassword_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6720    intShiftDown = (Shift And acShiftMask) > 0
6730    intAltDown = (Shift And acAltMask) > 0
6740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6770        With Me
6780          intRetVal = 0
6790          .cmdOK.SetFocus
6800        End With
6810      End Select
6820    End If

EXITP:
6830    KeyCode = intRetVal
6840    Exit Sub

ERRH:
6850    Select Case ERR.Number
        Case Else
6860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6870    End Select
6880    Resume EXITP

End Sub
