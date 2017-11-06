VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUser_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmUser_Add"

'VGC 03/17/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   cmdHelp_Click()
' #######################################

'xusr_extant     : Current password
'xusr_antecedent : Previous password
'xusr_origin     : Password date

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Password Help:    H {cmdHelp}
' **   OK:               O {cmdOK}
' **   User Name:        U {Username on frmUser_Add_Sub}

Private blnHelp_Focus As Boolean, blnHelp_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngHelpBox_Top As Long, lngHelpbox_Height As Long
Private lngHelpLbl_Top As Long, lngHelpLbl_Height As Long, lngFooterHgtLine_Top As Long, lngHelp_OFFSET As Long

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

190           blnHelp_Focus = False: blnHelp_MouseDown = False

200           strHelpCap = .cmdHelp_lbl.Caption
210           lngHelpBox_Top = .cmdHelp_box.Top
220           lngHelpLbl_Top = .cmdHelp_lbl.Top
230           lngHelpbox_Height = .cmdHelp_box.Height
240           lngHelpLbl_Height = .cmdHelp_lbl.Height
250           lngFooterHgtLine_Top = .form_footer_height_line.Top
260           lngHelp_OFFSET = (.FormFooter.Height - .Footer_vline03.Top)

270           lngTmp03 = -2&: strTmp01 = vbNullString: strTmp02 = vbNullString
280           lngTmp03 = Security_PW_Exp_Get("Cycle")  ' ** Module Function: modSecurityFuncs.
290           strTmp01 = .cmdHelp_lbl.Caption
300           Select Case lngTmp03
              Case 28&
310             strTmp02 = "Password must be changed every 4 weeks"
320           Case 30&
330             strTmp02 = "Password must be changed every 30 days"
340           Case 365&
350             strTmp02 = "Password must be changed every 12 months"
360           Case 0&
370             strTmp02 = "Password does not expire"
380           Case -1&  ' ** Not found.
390             strTmp02 = "NOT FOUND!"
400           Case -9&  ' ** Error.
410             strTmp02 = RET_ERR
420           Case Else
430             strTmp02 = "Password must be changed every " & CStr(lngTmp03) & " days"
440           End Select

450           intPos01 = InStr(strTmp01, "Password must be changed")
460           If intPos01 > 0 Then
470             intPos02 = 0
480             For intX = intPos01 To Len(strTmp01)
490               If Mid(strTmp01, intX, 1) = vbCr Or Mid(strTmp01, intX, 1) = vbLf Then
500                 intPos02 = intX
510                 Exit For
520               End If
530             Next
540             If intPos02 > 0 Then
550               .cmdHelp_lbl.Caption = Left(strTmp01, (intPos01 - 1)) & strTmp02 & Mid(strTmp01, intPos02)
560             End If
570           End If

              ' ** Hide the password help.
580           cmdHelp_Click  ' ** Procedure: Below.
590           DoEvents

600           blnIsOpen = False

610         Else
620           Cancel = -1
630         End If
640       Else
650         Cancel = -1
660       End If
670     End With

680     If Cancel = -1 Then
690       Beep
700       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
710     End If

EXITP:
720     Exit Sub

ERRH:
730     Select Case ERR.Number
        Case Else
740       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
750     End Select
760     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         intRetVal = 0
880         cmdCancel_Click  ' ** Procedure: Below.
890       End Select
900     End If

EXITP:
910     KeyCode = intRetVal
920     Exit Sub

ERRH:
930     intRetVal = 0
940     Select Case ERR.Number
        Case Else
950       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
960     End Select
970     Resume EXITP

End Sub

Private Sub Form_Close()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Close"

1010    With Forms("frmUser")
1020      .frmUser_Sub.Form.Requery
1030      gblnSetFocus = True
1040      .TimerInterval = 100&
1050    End With

EXITP:
1060    Exit Sub

ERRH:
1070    Select Case ERR.Number
        Case Else
1080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1090    End Select
1100    Resume EXITP

End Sub

Public Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1210    With Me
1220      If ((.cmdHelp_R_raised_focus_dots_img.Visible = True Or .cmdHelp_R_raised_focus_img.Visible = True) Or _
              (.cmdHelp_L_raised_focus_dots_img.Visible = True Or .cmdHelp_L_raised_focus_img.Visible = True)) Then
1230        Select Case .cmdHelp_box.Visible
            Case True
1240          Select Case blnHelp_Focus
              Case True
1250            .cmdHelp_L_raised_semifocus_dots_img.Visible = True
1260            .cmdHelp_L_raised_img.Visible = False
1270            .cmdHelp_R_raised_img.Visible = False
1280            .cmdHelp_R_raised_semifocus_dots_img.Visible = False
1290          Case False
1300            .cmdHelp_L_raised_img.Visible = True
1310            .cmdHelp_L_raised_semifocus_dots_img.Visible = False
1320            .cmdHelp_R_raised_img.Visible = False
1330            .cmdHelp_R_raised_semifocus_dots_img.Visible = False
1340          End Select
1350        Case False
1360          Select Case blnHelp_Focus
              Case True
1370            .cmdHelp_R_raised_semifocus_dots_img.Visible = True
1380            .cmdHelp_R_raised_img.Visible = False
1390            .cmdHelp_L_raised_img.Visible = False
1400            .cmdHelp_L_raised_semifocus_dots_img.Visible = False
1410          Case False
1420            .cmdHelp_R_raised_img.Visible = True
1430            .cmdHelp_R_raised_semifocus_dots_img.Visible = False
1440            .cmdHelp_L_raised_img.Visible = False
1450            .cmdHelp_L_raised_semifocus_dots_img.Visible = False
1460          End Select
1470        End Select
1480        .cmdHelp_R_raised_focus_img.Visible = False
1490        .cmdHelp_R_raised_focus_dots_img.Visible = False
1500        .cmdHelp_R_sunken_focus_dots_img.Visible = False
1510        .cmdHelp_L_raised_focus_img.Visible = False
1520        .cmdHelp_L_raised_focus_dots_img.Visible = False
1530        .cmdHelp_L_sunken_focus_dots_img.Visible = False
1540      End If
1550    End With

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

1710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOk_Click"

        Dim usr As DAO.User, frm As Access.Form
        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strUsername As String
        Dim strUserPIN As String
        Dim strSelectedGroup As String
        Dim blnUserExists As Boolean, blnErr As Boolean
        Dim lngGroupID As Long, lngGroupID_Users As Long, lngUserID As Long
        Dim lngRecs As Long, lngXUsrID As Long, lngErrNum As Long
        Dim varTmp00 As Variant
        Dim lngX As Long

1810    DoCmd.Hourglass True
1820    lngGroupID = 0&: lngUserID = 0&: lngErrNum = 0&
1830    lngGroupID_Users = DLookup("[secgrp_id]", "tblSecurity_Group", "[secgrp_name] = 'Users'")

1840    With Me

1850      Set frm = .frmUser_Add_Sub.Form

          ' ** Employee Username.
1860      If IsNull(frm.Username) Then
1870        DoCmd.Hourglass False
1880        MsgBox "You must enter a User Name to continue.", vbInformation + vbOKOnly, "Entry Required"
1890      ElseIf Len(frm.Username) > 20 Or Len(frm.Username) < 4 Or InStr(frm.Username, " ") > 0 Then
1900        DoCmd.Hourglass False
1910        MsgBox "The User Name cannot be longer than 20 characters, nor shorter than 4 characters," & vbCrLf & _
              "and cannot contain a space.", vbInformation + vbOKOnly, "Invalid Entry"
1920      Else
1930        strUsername = frm.Username
1940        If IsNull(frm.Employee_Name) Then
1950          DoCmd.Hourglass False
1960          MsgBox "You must enter the employee's name to continue.", vbInformation + vbOKOnly, "Entry Required"
1970        Else
              ' ** Employee Security File PID number.
1980          strUserPIN = frm.Username
              ' ** Specific Security Group User Belongs to.
1990          If IsNull(frm.Primary_Group) Then
2000            DoCmd.Hourglass False
2010            MsgBox "You must select a security level to continue.", vbInformation + vbOKOnly, "Entry Required"
2020          Else
2030            strSelectedGroup = frm.Primary_Group
2040            If strSelectedGroup = vbNullString Then
2050              DoCmd.Hourglass False
2060              MsgBox "You must select a security level to continue.", vbInformation + vbOKOnly, "Entry Required"
2070            Else
2080              lngGroupID = CLng(frm.Primary_Group.Column(1))
                  ' ** User password.
2090              If IsNull(frm.NewPassword) = True Then
2100                DoCmd.Hourglass False
2110                MsgBox "All users must have a password." & vbCrLf & vbCrLf & _
                      "Assign a temporary password, maximum 14 characters." & vbCrLf & "The new user may change it later.", _
                      vbInformation + vbOKOnly, "Password Required"
2120              ElseIf frm.NewPassword = vbNullString Then
2130                DoCmd.Hourglass False
2140                MsgBox "All users must have a password." & vbCrLf & vbCrLf & _
                      "Assign a temporary password, maximum 14 characters." & vbCrLf & "The new user may change it later.", _
                      vbInformation + vbOKOnly, "Password Required"
2150              ElseIf Pass_Check(frm.NewPassword, frm.Username) = False Then  ' ** Module Function: modCodeUtilities.
                    ' ** Failure message shown in Pass_Check() of modCodeUtilities.
2160              Else
2170                If IsNull(frm.VerifyNewPassword) = True Then
2180                  DoCmd.Hourglass False
2190                  MsgBox "You must verify the new password.", vbInformation + vbOKOnly, "Password Verification Required"
2200                  frm.VerifyNewPassword.SetFocus
2210                Else
2220                  If frm.NewPassword <> frm.VerifyNewPassword Then
2230                    DoCmd.Hourglass False
2240                    MsgBox "The second entry of the password didn't match the first." & vbCrLf & vbCrLf & _
                          "Please re-enter the password so that both match.", vbExclamation + vbOKOnly, "No Match"
2250                  Else

2260                    blnUserExists = False: blnErr = False

                        ' ** Create new user account.
2270                    Set usr = DBEngine.Workspaces(0).CreateUser(strUsername, strUserPIN)

                        ' ** Save user account definition by appending it to Users collection, along with its password.
2280  On Error Resume Next
2290                    DBEngine.Workspaces(0).Users.Append usr
2300                    If ERR.Number <> 0 Then
2310                      Select Case ERR.Number
                          Case 3390  ' ** Account name already exists.
                            ' ** Since we're creating it, the error is moot; ignore it.
2320  On Error GoTo ERRH
2330                        blnUserExists = True
2340                        lngErrNum = 1&
2350                      Case Else
2360                        blnErr = True
2370                        zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2380  On Error GoTo ERRH
2390                        lngErrNum = 2&
2400                      End Select
2410                    Else
2420  On Error GoTo ERRH
2430                    End If
2440                    DBEngine.Workspaces(0).Users.Refresh
2450                    DBEngine.Workspaces(0).Users(strUsername).NewPassword vbNullString, frm.NewPassword

2460                    If blnErr = False Then

2470                      Set usr = DBEngine.Workspaces(0).Users(strUsername)

                          ' ** Add user to predefined Users group.
2480                      If blnUserExists = False Then
2490                        usr.Groups.Append usr.CreateGroup("Users")
2500                      End If

2510                      Set dbs = CurrentDb

                          ' ** Add user to tblSecurity_User.
2520                      Set rst = dbs.OpenRecordset("tblSecurity_User", dbOpenDynaset, dbConsistent)
2530                      With rst
2540                        .AddNew
2550                        ![secusr_name] = strUsername
2560                        ![secusr_default] = False
2570                        ![secusr_user] = GetUserName  ' ** Module Function: modFileUtilities.
2580                        ![secusr_datecreated] = Now()
2590                        ![secusr_datemodified] = Now()
2600                      End With
2610  On Error Resume Next
2620                      rst.Update
2630                      If ERR.Number <> 0 Then
2640                        Select Case ERR.Number
                            Case 3022  ' ** The changes you requested to the table were not successful because they would
2650                          rst.CancelUpdate  ' ** create duplicate values in the index, primary key, or relationship.
2660  On Error GoTo ERRH
2670                          lngErrNum = 3&
2680                          rst.Close
                              ' ** Delete from tblSecurity_User by specified [usr].
2690                          Set qdf = dbs.QueryDefs("qryUserMaint_03")
2700                          With qdf.Parameters
2710                            ![usr] = strUsername
2720                          End With
2730                          qdf.Execute
                              ' ** Then re-enter it.
2740                          Set rst = dbs.OpenRecordset("tblSecurity_User", dbOpenDynaset, dbConsistent)
2750                          With rst
2760                            .AddNew
2770                            ![secusr_name] = strUsername
2780                            ![secusr_default] = False
2790                            ![secusr_user] = GetUserName  ' ** Module Function: modFileUtilities.
2800                            ![secusr_datecreated] = Now()
2810                            ![secusr_datemodified] = Now()
2820  On Error Resume Next
2830                            .Update
2840                            If ERR.Number <> 0 Then
                                  ' ** So what do I do here?
2850                              blnErr = True
2860  On Error GoTo ERRH
2870                              lngErrNum = 4&
2880                            Else
2890  On Error GoTo ERRH
2900                            End If
2910                          End With
2920                        Case Else
2930                          blnErr = True
2940                          zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950  On Error GoTo ERRH
2960                          lngErrNum = 5&
2970                          rst.Close
2980                        End Select
2990                      Else
3000  On Error GoTo ERRH
3010                      End If
3020                      If blnErr = False Then
3030                        rst.Bookmark = rst.LastModified
3040                        lngUserID = rst![secusr_id]
3050                        rst.Close
3060                      End If
3070                    End If

3080                    If blnErr = False Then

                          ' ** Add User to tblSecurity_GroupUser for 'Users' group.
3090                      Set rst = dbs.OpenRecordset("tblSecurity_GroupUser", dbOpenDynaset, dbConsistent)
3100                      With rst
3110                        .AddNew
3120                        ![secgrp_id] = lngGroupID_Users
3130                        ![secusr_id] = lngUserID
3140                        ![secgrpusr_default] = False
3150                        ![secgrpusr_user] = GetUserName  ' ** Module Function: modFileUtilities.
3160                        ![secgrpusr_datecreated] = Now()
3170                        ![secgrpusr_datemodified] = Now()
3180  On Error Resume Next
3190                        .Update
3200                        If ERR.Number <> 0 Then
3210                          blnErr = True
3220  On Error GoTo ERRH
3230                          lngErrNum = 6&
3240                        Else
3250  On Error GoTo ERRH
3260                        End If
3270                        .Close
3280                      End With

3290                      Set usr = DBEngine.Workspaces(0).Users(strUsername)

                          ' ** Add user to Selected group.
3300                      usr.Groups.Append usr.CreateGroup(strSelectedGroup)

3310                      DoCmd.RunCommand acCmdSaveRecord

                          ' ** Add User to tblSecurity_GroupUser for selected group.
3320                      Set rst = dbs.OpenRecordset("tblSecurity_GroupUser", dbOpenDynaset, dbConsistent)
3330                      With rst
3340                        .AddNew
3350                        ![secgrp_id] = lngGroupID
3360                        ![secusr_id] = lngUserID
3370                        ![secgrpusr_default] = False
3380                        ![secgrpusr_user] = GetUserName  ' ** Module Function: modFileUtilities.
3390                        ![secgrpusr_datecreated] = Now()
3400                        ![secgrpusr_datemodified] = Now()
3410  On Error Resume Next
3420                        .Update
3430                        If ERR.Number <> 0 Then
3440                          blnErr = True
3450  On Error GoTo ERRH
3460                          lngErrNum = 7&
3470                        Else
3480  On Error GoTo ERRH
3490                        End If
3500                        .Close
3510                      End With

                          ' ** Add user's info to _~xusr.
3520                      Set rst = dbs.OpenRecordset("_~xusr", dbOpenDynaset, dbConsistent)
3530                      With rst
3540                        .AddNew
                            ' ** The StringFromGUID function converts a GUID, which is an array of type Byte, to a string.
3550                        varTmp00 = FilterGUIDString(StringFromGUID(frm.s_GUID))  ' ** Module Function: modCodeUtilities.
3560                        ![s_GUID] = CStr(varTmp00)  ' ** Yes, a GUID field just takes a plain string!
3570                        ![xusr_extant] = EncodeString(frm.NewPassword)  ' ** Module Function: modCodeUtilities.
3580                        ![xusr_antecedent] = Null
3590                        ![xusr_origin] = EncodeString(Format(Date, "mm/dd/yyyy"))  ' ** Module Function: modCodeUtilities.
3600                        ![xusr_user] = frm.Username  ' ** THIS IS THE USER'S NAME, NOT GetUserName!
3610                        ![xusr_datecreated] = Now()
3620                        ![xusr_datemodified] = Now()
3630                      End With
3640  On Error Resume Next
3650                      rst.Update
3660                      If ERR.Number <> 0 Then
3670                        Select Case ERR.Number
                            Case 3022  ' ** The changes you requested to the table were not successful because they would
3680                          rst.CancelUpdate  ' ** create duplicate values in the index, primary key, or relationship.
3690  On Error GoTo ERRH
3700                          lngErrNum = 8&
3710                          rst.Close
3720                          Set rst = dbs.OpenRecordset("_~xusr", dbOpenDynaset, dbConsistent)
3730                          With rst
3740                            .MoveLast
3750                            lngRecs = .RecordCount
3760                            .MoveFirst
3770                            For lngX = 1& To lngRecs
3780                              If FilterGUIDString(StringFromGUID(![s_GUID])) = varTmp00 Then  ' ** Module Function: modCodeUtilities.
3790                                lngXUsrID = ![xusr_id]
3800                                Exit For
3810                              End If
3820                              If lngX < lngRecs Then .MoveNext
3830                            Next
3840                            .Close
3850                          End With
                              ' ** Delete user from _~xusr, by specified [xusr].
3860                          Set qdf = dbs.QueryDefs("qryXUsr_01")
3870                          With qdf.Parameters
3880                            ![xusr] = lngXUsrID
3890                          End With
3900                          qdf.Execute
                              ' ** Re-enter the user.
3910                          Set rst = dbs.OpenRecordset("_~xusr", dbOpenDynaset, dbConsistent)
3920                          With rst
3930                            .AddNew
                                ' ** The StringFromGUID function converts a GUID, which is an array of type Byte, to a string.
3940                            varTmp00 = FilterGUIDString(StringFromGUID(frm.s_GUID))  ' ** Module Function: modCodeUtilities.
3950                            ![s_GUID] = CStr(varTmp00)  ' ** Yes, a GUID field just takes a plain string!
3960                            ![xusr_extant] = EncodeString(frm.NewPassword)  ' ** Module Function: modCodeUtilities.
3970                            ![xusr_antecedent] = Null
3980                            ![xusr_origin] = EncodeString(Format(Date, "mm/dd/yyyy"))  ' ** Module Function: modCodeUtilities.
3990                            ![xusr_user] = frm.Username  ' ** THIS IS THE USER'S NAME, NOT GetUserName!
4000                            ![xusr_datecreated] = Now()
4010                            ![xusr_datemodified] = Now()
4020  On Error Resume Next
4030                            .Update
4040                            If ERR.Number <> 0 Then
4050                              blnErr = True
4060  On Error GoTo ERRH
4070                              lngErrNum = 9&
4080                            Else
4090  On Error GoTo ERRH
4100                            End If
4110                            .Close
4120                          End With
4130                        Case Else
4140                          blnErr = True
4150                          zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4160  On Error GoTo ERRH
4170                          lngErrNum = 10&
4180                        End Select
4190                      Else
4200  On Error GoTo ERRH
4210                        rst.Close
4220                      End If

4230                    End If
4240                    If blnErr = False Then
4250                      DoCmd.Close acForm, THIS_NAME
4260                    End If
4270                  End If
4280                End If
4290              End If
4300            End If
4310          End If
4320        End If
4330      End If

          ' ** lngErrNum:
          ' ** 1.  Error appending to Users, already exists.
          ' ** 2.  Error appending to Users, other error.
          ' ** 3.  Error appending to tblSecurity_User, duplicate key.
          ' ** 4.  Error appending to tblSecurity_User, 2nd try.
          ' ** 5.  Error appending to tblSecurity_User, other error.
          ' ** 6.  Error appending to tblSecurity_GroupUser, for 'Users' group.
          ' ** 7.  Error appending to tblSecurity_GroupUser, for other group.
          ' ** 8.  Error appending to _~xusr, duplicate key.
          ' ** 9.  Error appending to _~xusr, 2nd try.
          ' ** 10. Error appending to _~xusr, other error.

4340    End With

EXITP:
4350    DoCmd.Hourglass False
4360    Set frm = Nothing
4370    Set usr = Nothing
4380    Set rst = Nothing
4390    Set qdf = Nothing
4400    Set dbs = Nothing
4410    Exit Sub

ERRH:
4420    Beep
4430    Select Case ERR.Number
        Case 3390, 3022
          ' ** 3390: The user already exists;
          ' ** 3022: The changes you requested to the table were not successful because
          ' ** they would create duplicate values in the index, primary key, or relationship.
4440      MsgBox "The user already exists." & vbCrLf & "CodeLoc: " & CStr(lngErrNum), vbInformation + vbOKOnly, "User Exists"
4450    Case Else
4460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4470    End Select
4480    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab
4570        With Me
4580          intRetVal = 0
4590          .cmdCancel.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab
4650        With Me
4660          intRetVal = 0
4670          .frmUser_Add_Sub.SetFocus
4680          .frmUser_Add_Sub.Form.VerifyNewPassword.SetFocus
4690        End With
4700      End Select
4710    End If

EXITP:
4720    KeyCode = intRetVal
4730    Exit Sub

ERRH:
4740    Select Case ERR.Number
        Case Else
4750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4760    End Select
4770    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strUser As String
        Dim varTmp00 As Variant

4810    With Me
4820      varTmp00 = .frmUser_Add_Sub.Form.Username
4830      If IsNull(varTmp00) = False Then
4840        strUser = .frmUser_Add_Sub.Form.Username
4850        Set dbs = CurrentDb
4860        With dbs
              ' ** Delete from Users by specified [usr].
4870          Set qdf = .QueryDefs("qryUserMaint_02")
4880          With qdf.Parameters
4890            ![usr] = strUser
4900          End With
4910          qdf.Execute
4920          .Close
4930        End With
4940      End If
4950      DoCmd.Close acForm, THIS_NAME
4960    End With

EXITP:
4970    Set qdf = Nothing
4980    Set dbs = Nothing
4990    Exit Sub

ERRH:
5000    Select Case ERR.Number
        Case Else
5010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5020    End Select
5030    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab
5170        With Me
5180          intRetVal = 0
5190          .frmUser_Add_Sub.SetFocus
5200          .frmUser_Add_Sub.Form.Username.SetFocus
5210        End With
5220      End Select
5230    End If

        ' ** Shift keys.
5240    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5250      Select Case intRetVal
          Case vbKeyTab
5260        With Me
5270          intRetVal = 0
5280          .cmdOK.SetFocus
5290        End With
5300      End Select
5310    End If

EXITP:
5320    KeyCode = intRetVal
5330    Exit Sub

ERRH:
5340    Select Case ERR.Number
        Case Else
5350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5360    End Select
5370    Resume EXITP

End Sub

Private Sub cmdHelp_GotFocus()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_GotFocus"

5410    With Me
5420      blnHelp_Focus = True
5430      Select Case .cmdHelp_box.Visible
          Case True
5440        .cmdHelp_L_raised_semifocus_dots_img.Visible = True
5450        .cmdHelp_R_raised_semifocus_dots_img.Visible = False
5460      Case False
5470        .cmdHelp_R_raised_semifocus_dots_img.Visible = True
5480        .cmdHelp_L_raised_semifocus_dots_img.Visible = False
5490      End Select
5500      .cmdHelp_R_raised_img.Visible = False
5510      .cmdHelp_R_raised_focus_img.Visible = False
5520      .cmdHelp_R_raised_focus_dots_img.Visible = False
5530      .cmdHelp_R_sunken_focus_dots_img.Visible = False
5540      .cmdHelp_L_raised_img.Visible = False
5550      .cmdHelp_L_raised_focus_img.Visible = False
5560      .cmdHelp_L_raised_focus_dots_img.Visible = False
5570      .cmdHelp_L_sunken_focus_dots_img.Visible = False
5580    End With

EXITP:
5590    Exit Sub

ERRH:
5600    Select Case ERR.Number
        Case Else
5610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5620    End Select
5630    Resume EXITP

End Sub

Private Sub cmdHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseDown"

5710    With Me
5720      blnHelp_MouseDown = True
5730      Select Case .cmdHelp_box.Visible
          Case True
5740        .cmdHelp_L_sunken_focus_dots_img.Visible = True
5750        .cmdHelp_R_sunken_focus_dots_img.Visible = False
5760      Case False
5770        .cmdHelp_R_sunken_focus_dots_img.Visible = True
5780        .cmdHelp_L_sunken_focus_dots_img.Visible = False
5790      End Select
5800      .cmdHelp_R_raised_img.Visible = False
5810      .cmdHelp_R_raised_semifocus_dots_img.Visible = False
5820      .cmdHelp_R_raised_focus_img.Visible = False
5830      .cmdHelp_R_raised_focus_dots_img.Visible = False
5840      .cmdHelp_L_raised_img.Visible = False
5850      .cmdHelp_L_raised_semifocus_dots_img.Visible = False
5860      .cmdHelp_L_raised_focus_img.Visible = False
5870      .cmdHelp_L_raised_focus_dots_img.Visible = False
5880    End With

EXITP:
5890    Exit Sub

ERRH:
5900    Select Case ERR.Number
        Case Else
5910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5920    End Select
5930    Resume EXITP

End Sub

Private Sub cmdHelp_Click()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_Click"

        Dim frm As Access.Form
        Dim strUser As String, strControl As String
        Dim varTmp00 As Variant, lngTmp01 As Long

6010    With Me

6020      Set frm = .frmUser_Add_Sub.Form
6030      Select Case IsNull(frm.Username)
          Case True
6040        strUser = vbNullString
6050      Case False
6060        strUser = frm.Username
6070      End Select

6080  On Error Resume Next
6090      strControl = Screen.PreviousControl.Name
6100  On Error GoTo ERRH

6110      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
6120      lngMonitorNum = 1&: lngTmp01 = 0&
6130      EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
6140      If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

6150      Select Case .cmdHelp_lbl.Visible
          Case True
            ' ** Close the Help box.

            ' ** Shrink and move all the Help stuff out of the way.
6160        .form_footer_height_line.Top = .cmdOK.Top
6170        .cmdHelp_lbl.Visible = False
6180        .cmdHelp_box.Visible = False
6190        .cmdHelp_lbl.Height = 15&
6200        .cmdHelp_box.Height = 15&
6210        .cmdHelp_lbl.Top = .cmdOK.Top
6220        .cmdHelp_box.Top = .cmdOK.Top
6230        .Footer_hline03.Visible = False
6240        .Footer_hline04.Visible = False
6250        .Footer_vline03.Visible = False
6260        .Footer_vline04.Visible = False

            ' ** Get the 2nd set of lines out of the way.
6270        .Footer_hline03.Top = .Footer_hline03.Top - 360&
6280        .Footer_hline04.Top = .Footer_hline04.Top - 360&
6290        .Footer_vline03.Top = .Footer_vline03.Top - 360&
6300        .Footer_vline04.Top = .Footer_vline04.Top - 360&

            ' ** Now close up the footer, and shrink the form.
6310        .FormFooter.Height = .FormFooter.Height - lngHelp_OFFSET
6320        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
6330        DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, (lngFrm_Height - lngHelp_OFFSET)  'lngFrm_Top
6340        If lngMonitorNum > 1& Then
6350          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
6360        End If

6370      Case False
            ' ** Open the Help box.

            ' ** Open up the footer
6380        .FormFooter.Height = .FormFooter.Height + lngHelp_OFFSET

            ' ** Move the 2nd set of lines back to where they belong.
6390        .Footer_hline03.Top = .Footer_hline03.Top + 360&
6400        .Footer_hline04.Top = .Footer_hline04.Top + 360&
6410        .Footer_vline03.Top = .Footer_vline03.Top + 360&
6420        .Footer_vline04.Top = .Footer_vline04.Top + 360&

            ' ** Set up the help stuff.
6430        .form_footer_height_line.Top = lngFooterHgtLine_Top
6440        .cmdHelp_box.Top = lngHelpBox_Top
6450        .cmdHelp_lbl.Top = lngHelpLbl_Top
6460        .cmdHelp_box.Height = lngHelpbox_Height
6470        .cmdHelp_lbl.Height = lngHelpLbl_Height
6480        .cmdHelp_box.Visible = True
6490        .cmdHelp_lbl.Visible = True
6500        .Footer_hline03.Visible = True
6510        .Footer_hline04.Visible = True
6520        .Footer_vline03.Visible = True
6530        .Footer_vline04.Visible = True

            ' ** And lengthen the form.
6540        If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
6550        DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, lngFrm_Height  'lngFrm_Top
6560        If lngMonitorNum > 1& Then
6570          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
6580        End If

6590      End Select

6600      If strUser <> vbNullString Then
6610        Set frm = .frmUser_Add_Sub.Form
6620        varTmp00 = frm.Username
6630        Select Case IsNull(varTmp00)
            Case True
6640          frm.MoveRec strUser  ' ** Form Procedure: frmUser_Add_Sub.
6650        Case False
6660          If frm.Username <> strUser Then
6670            frm.MoveRec strUser  ' ** Form Procedure: frmUser_Add_Sub.
6680          End If
6690        End Select
6700      End If

6710      If strControl <> vbNullString Then
6720  On Error Resume Next
6730        .Controls(strControl).SetFocus
6740  On Error GoTo ERRH
6750      Else
6760        .cmdOK.SetFocus
6770      End If

6780    End With

EXITP:
6790    Set frm = Nothing
6800    Exit Sub

ERRH:
6810    Select Case ERR.Number
        Case Else
6820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6830    End Select
6840    Resume EXITP

End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseMove"

6910    With Me
6920      If blnHelp_MouseDown = False Then
6930        Select Case .cmdHelp_box.Visible
            Case True
6940          Select Case blnHelp_Focus
              Case True
6950            .cmdHelp_L_raised_focus_dots_img.Visible = True
6960            .cmdHelp_L_raised_focus_img.Visible = False
6970            .cmdHelp_R_raised_focus_img.Visible = False
6980            .cmdHelp_R_raised_focus_dots_img.Visible = False
6990          Case False
7000            .cmdHelp_L_raised_focus_img.Visible = True
7010            .cmdHelp_L_raised_focus_dots_img.Visible = False
7020            .cmdHelp_R_raised_focus_img.Visible = False
7030            .cmdHelp_R_raised_focus_dots_img.Visible = False
7040          End Select
7050        Case False
7060          Select Case blnHelp_Focus
              Case True
7070            .cmdHelp_R_raised_focus_dots_img.Visible = True
7080            .cmdHelp_R_raised_focus_img.Visible = False
7090            .cmdHelp_L_raised_focus_img.Visible = False
7100            .cmdHelp_L_raised_focus_dots_img.Visible = False
7110          Case False
7120            .cmdHelp_R_raised_focus_img.Visible = True
7130            .cmdHelp_R_raised_focus_dots_img.Visible = False
7140            .cmdHelp_L_raised_focus_img.Visible = False
7150            .cmdHelp_L_raised_focus_dots_img.Visible = False
7160          End Select
7170        End Select
7180        .cmdHelp_R_raised_img.Visible = False
7190        .cmdHelp_R_raised_semifocus_dots_img.Visible = False
7200        .cmdHelp_R_sunken_focus_dots_img.Visible = False
7210        .cmdHelp_L_raised_img.Visible = False
7220        .cmdHelp_L_raised_semifocus_dots_img.Visible = False
7230        .cmdHelp_L_sunken_focus_dots_img.Visible = False
7240        blnHelp_MouseDown = False
7250      End If
7260    End With

EXITP:
7270    Exit Sub

ERRH:
7280    Select Case ERR.Number
        Case Else
7290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7300    End Select
7310    Resume EXITP

End Sub

Private Sub cmdHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_MouseUp"

7410    With Me
7420      Select Case .cmdHelp_box.Visible
          Case True
7430        .cmdHelp_L_raised_focus_dots_img.Visible = True
7440        .cmdHelp_R_raised_focus_dots_img.Visible = False
7450      Case False
7460        .cmdHelp_R_raised_focus_dots_img.Visible = True
7470        .cmdHelp_L_raised_focus_dots_img.Visible = False
7480      End Select
7490      .cmdHelp_R_raised_img.Visible = False
7500      .cmdHelp_R_raised_semifocus_dots_img.Visible = False
7510      .cmdHelp_R_raised_focus_img.Visible = False
7520      .cmdHelp_R_sunken_focus_dots_img.Visible = False
7530      .cmdHelp_L_raised_img.Visible = False
7540      .cmdHelp_L_raised_semifocus_dots_img.Visible = False
7550      .cmdHelp_L_raised_focus_img.Visible = False
7560      .cmdHelp_L_sunken_focus_dots_img.Visible = False
7570      blnHelp_MouseDown = False
7580    End With

EXITP:
7590    Exit Sub

ERRH:
7600    Select Case ERR.Number
        Case Else
7610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7620    End Select
7630    Resume EXITP

End Sub

Private Sub cmdHelp_LostFocus()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_LostFocus"

7710    With Me
7720      Select Case .cmdHelp_box.Visible
          Case True
7730        .cmdHelp_L_raised_img.Visible = True
7740        .cmdHelp_R_raised_img.Visible = False
7750      Case False
7760        .cmdHelp_R_raised_img.Visible = True
7770        .cmdHelp_L_raised_img.Visible = False
7780      End Select
7790      .cmdHelp_R_raised_semifocus_dots_img.Visible = False
7800      .cmdHelp_R_raised_focus_img.Visible = False
7810      .cmdHelp_R_raised_focus_dots_img.Visible = False
7820      .cmdHelp_R_sunken_focus_dots_img.Visible = False
7830      .cmdHelp_L_raised_semifocus_dots_img.Visible = False
7840      .cmdHelp_L_raised_focus_img.Visible = False
7850      .cmdHelp_L_raised_focus_dots_img.Visible = False
7860      .cmdHelp_L_sunken_focus_dots_img.Visible = False
7870      blnHelp_Focus = False
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

Private Sub cmdHelp_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_box_MouseMove"

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

Private Sub cmdHelp_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHelp_lbl_MouseMove"

8110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
8120    Exit Sub

ERRH:
8130    Select Case ERR.Number
        Case Else
8140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8150    End Select
8160    Resume EXITP

End Sub
