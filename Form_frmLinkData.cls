VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLinkData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLinkData"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Unlock Aux:       A {chkUnlock_Aux}
' **   Browse Data:      B {cmdBrowse}
' **   Close:            C {cmdClose}
' **   Unlock Archive:   H {chkUnlock_Arch}
' **   Unlock Main:      M {chkUnlock_Dta}
' **   Compact Now:      N {cmdCompact}
' **   Refresh/Relink:   R {cmdRefresh, cmdRelink}
' **   Set Default:      S {cmdSetDefPath}
' **   Use This Path:    U {cmdUseThisPath}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Form:       X {cmdCloseForm}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const LNK_REFRESH As Integer = 1
Private Const LNK_CHANGE  As Integer = 2

' ** Progress Bar Variables.
Private frmPBar As Access.Form
Private ctlPBox As Access.Rectangle
Private dblPBar_MaxWidth As Double
Private dblPBar_CurWidth As Double
Private dblPBar_Steps As Double
Private dblPBar_Increment As Double
Private dblPBar_ThisStep As Double
Private dblPBar_ThisIncrement As Double

Private blnBrowse_Focus As Boolean, blnBrowse_MouseDown As Boolean
Private blnUseThisPath_Focus As Boolean, blnUseThisPath_MouseDown As Boolean
Private blnSetDefPath_Focus As Boolean, blnSetDefPath_MouseDown As Boolean

Private intLinkMode As Integer, strPathFile_This As String, blnProgBarOff As Boolean
Private strPathFile_Data As String, strPathFile_Archive As String, strPathFile_Auxiliary As String
Private lngDataSize As Long, lngArchSize As Long, lngAuxSize As Long, blnWasKb As Boolean
Private blnJustClose As Boolean, blnCompactQuit As Boolean, blnAuxLoc As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnErrDisplayed As Boolean
        Dim varTmp00 As Variant

110     With Me

120       strCallingForm = vbNullString
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then
150           strCallingForm = Trim(.OpenArgs)
160         End If
170       End If

180       CLR_DISABLED_FG = CLR_DKGRY
190       CLR_DISABLED_BG = CLR_LTTEAL

200       varTmp00 = DLookup("[seclic_auxloc]", "tblSecurity_License")
210       If IsNull(varTmp00) = False Then
220         .chkAuxLoc = CBool(varTmp00)
230       End If

240       strPathFile_This = CurrentDb.Name
250       .PathFile_This = strPathFile_This
260       blnProgBarOff = False

          ' ** I'd like to keep the warning options hidden until they've seen the messages at least once.
270       Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

280       chkFileSizeWarning_AfterUpdate  ' ** Procedure: Below.
290       chkRecordCountWarning_AfterUpdate  ' ** Procedure: Below.
300       chkAuxLoc_AfterUpdate  ' ** Procedure: Below.

310       .Path_Default.DefaultValue = Chr(34) & GetOption("Default Database Directory") & Chr(34)
320       If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
330         .cmdSetDefPath.Visible = True
340         If .Path_Default = Parse_Path(.PathFile_This) Then  ' ** Module Function: modFileUtilities.
350           .cmdSetDefPath.Enabled = False
360           .cmdSetDefPath_raised_img_dis.Visible = True
370           .cmdSetDefPath_raised_img.Visible = False
380         Else
390           .cmdSetDefPath.Enabled = True
400           .cmdSetDefPath_raised_img.Visible = True
410           .cmdSetDefPath_raised_img_dis.Visible = False
420         End If
430         .cmdSetDefPath_raised_semifocus_dots_img.Visible = False
440         .cmdSetDefPath_raised_focus_img.Visible = False
450         .cmdSetDefPath_raised_focus_dots_img.Visible = False
460         .cmdSetDefPath_sunken_focus_dots_img.Visible = False
470       End If
480       DoEvents

490       blnAuxLoc = False
500       Select Case .chkNoLink
          Case True
510         blnErrDisplayed = True
520         Beep
530         MsgBox "A serious error has occurred, and Trust Accountant can" & vbCrLf & "no longer find your data file!" & _
              vbCrLf & vbCrLf & "Browse for the correct location and ReLink.", _
              vbCritical + vbOKOnly, "Trust Accountant Data Not Found"
540       Case False

550         blnErrDisplayed = False: blnJustClose = False
560         blnBrowse_Focus = False: blnBrowse_MouseDown = False
570         blnUseThisPath_Focus = False: blnUseThisPath_MouseDown = False
580         blnSetDefPath_Focus = False: blnSetDefPath_MouseDown = False

590         strPathFile_Data = CurrentBackendPathFile("Ledger")  ' ** Module Function: modFileUtilities.

600   On Error Resume Next
610         lngDataSize = FileLen(strPathFile_Data)
620         If ERR.Number = 53 Or ERR.Number = 76 Then  ' ** 53: File not found; 76: Path not found.
630   On Error GoTo ERRH
640           blnErrDisplayed = True
650           Beep
660           MsgBox "A serious error has occurred, and Trust Accountant can" & vbCrLf & "no longer find your data file!" & _
                vbCrLf & vbCrLf & "Browse for the correct location and ReLink.", _
                vbCritical + vbOKOnly, "Trust Accountant Data Not Found"
670         ElseIf ERR.Number = 75 Then  ' ** Path/File access error.
680   On Error GoTo ERRH
690           blnErrDisplayed = True
700           Beep
710           MsgBox "A serious error has occurred, and Trust Accountant can" & vbCrLf & "no longer access your data file!" & _
                vbCrLf & vbCrLf & "Check your network and folder access rights.", _
                vbCritical + vbOKOnly, "Trust Accountant Data Inaccessible"
720         ElseIf ERR.Number <> 0 Then
730   On Error GoTo ERRH
740           blnErrDisplayed = True
750           Beep
760           MsgBox "A serious error has occurred." & _
                vbCrLf & vbCrLf & ("Error: " & CStr(ERR.Number)) & vbCrLf & "Description: " & ERR.description, _
                vbCritical + vbOKOnly, "Trust Accountant Data Error"
770         Else
780   On Error GoTo ERRH
790         End If

800         strPathFile_Archive = CurrentBackendPathFile("LedgerArchive")  ' ** Module Function: modFileUtilities.

810   On Error Resume Next
820         lngArchSize = FileLen(strPathFile_Archive)
830         If ERR.Number = 53 Or ERR.Number = 76 Then  ' ** 53: File not found; 76: Path not found.
840   On Error GoTo ERRH
850           If blnErrDisplayed = False Then
860             Beep
870             MsgBox "A serious error has occurred, and Trust Accountant can" & vbCrLf & "no longer find your archive file!" & _
                  vbCrLf & vbCrLf & "Browse for the correct location and ReLink.", _
                  vbCritical + vbOKOnly, "Trust Accountant Archive Not Found"
880           End If
890         ElseIf ERR.Number = 75 Then  ' ** Path/File access error.
900   On Error GoTo ERRH
910           If blnErrDisplayed = False Then
920             Beep
930             MsgBox "A serious error has occurred, and Trust Accountant can" & vbCrLf & "no longer access your archive file!" & _
                  vbCrLf & vbCrLf & "Check your network and folder access rights.", _
                  vbCritical + vbOKOnly, "Trust Accountant Archive Inaccessible"
940           End If
950         ElseIf ERR.Number <> 0 Then
960   On Error GoTo ERRH
970           If blnErrDisplayed = False Then
980             Beep
990             MsgBox "A serious error has occurred." & _
                  vbCrLf & vbCrLf & ("Error: " & CStr(ERR.Number)) & vbCrLf & "Description: " & ERR.description, _
                  vbCritical + vbOKOnly, "Trust Accountant Archive Error"
1000          End If
1010        Else
1020  On Error GoTo ERRH
1030        End If

1040        strPathFile_Auxiliary = CurrentBackendPathFile("tblDatabase")  ' ** Module Function: modFileUtilities.

1050  On Error Resume Next
1060        lngAuxSize = FileLen(strPathFile_Auxiliary)
1070        If ERR.Number = 53 Or ERR.Number = 76 Then  ' ** 53: File not found; 76: Path not found.
1080  On Error GoTo ERRH
1090          If blnErrDisplayed = False Then
1100            Beep
1110            MsgBox "A serious error has occurred, and Trust Accountant can" & vbCrLf & "no longer find the auxiliary file!" & _
                  vbCrLf & vbCrLf & "Browse for the correct location and ReLink.", _
                  vbCritical + vbOKOnly, "Trust Accountant Auxiliary not Found"
1120          End If
1130        ElseIf ERR.Number = 75 Then  ' ** Path/File access error.
1140  On Error GoTo ERRH
1150          If blnErrDisplayed = False Then
1160            Beep
1170            MsgBox "A serious error has occurred, and Trust Accountant can" & vbCrLf & "no longer access the auxiliary file!" & _
                  vbCrLf & vbCrLf & "Check your network and folder access rights.", _
                  vbCritical + vbOKOnly, "Trust Accountant Auxiliary Inaccessible"
1180          End If
1190        ElseIf ERR.Number <> 0 Then
1200  On Error GoTo ERRH
1210          If blnErrDisplayed = False Then
1220            Beep
1230            MsgBox "A serious error has occurred." & _
                  vbCrLf & vbCrLf & ("Error: " & CStr(ERR.Number)) & vbCrLf & "Description: " & ERR.description, _
                  vbCritical + vbOKOnly, "Trust Accountant Auxiliary Error"
1240          End If
1250        Else
1260  On Error GoTo ERRH
1270        End If

1280        .PathFile_Dta = strPathFile_Data
1290        If (lngDataSize / 1024&) >= 5120000 Then
              ' ** Over 5Mb, display as Mb.
1300          .DataSize = IIf(Round(((lngDataSize / 1024&) / 1024&), 2) = CLng(Round(((lngDataSize / 1024&) / 1024&), 2)), _
                Format(Round(((lngDataSize / 1024&) / 1024&), 2), "#,##0"), _
                Format(Round(((lngDataSize / 1024&) / 1024&), 2), "#,##0.0#")) & " Mb"
1310        Else
              ' ** Under 5Mb, display as Kb.
1320          .DataSize = IIf(Round((lngDataSize / 1024&), 2) = CLng(Round((lngDataSize / 1024&), 2)), _
                Format(Round((lngDataSize / 1024&), 2), "#,##0"), _
                Format(Round((lngDataSize / 1024&), 2), "#,##0.0#")) & " Kb"
1330        End If

1340        .PathFile_Arch = strPathFile_Archive
1350        If (lngArchSize / 1024&) >= 5120000 Then
              ' ** Over 5Mb, display as Mb.
1360          .ArchSize = IIf(Round(((lngArchSize / 1024&) / 1024&), 2) = CLng(Round(((lngArchSize / 1024&) / 1024&), 2)), _
                Format(Round(((lngArchSize / 1024&) / 1024&), 2), "#,##0"), _
                Format(Round(((lngArchSize / 1024&) / 1024&), 2), "#,##0.0#")) & " Mb"
1370        Else
              ' ** Under 5Mb, display as Kb.
1380          .ArchSize = IIf(Round((lngArchSize / 1024&), 2) = CLng(Round((lngArchSize / 1024&), 2)), _
                Format(Round((lngArchSize / 1024&), 2), "#,##0"), _
                Format(Round((lngArchSize / 1024&), 2), "#,##0.0#")) & " Kb"
1390        End If

1400        .PathFile_Aux = strPathFile_Auxiliary
1410        If (lngAuxSize / 1024&) >= 5120000 Then
              ' ** Over 5Mb, display as Mb.
1420          .AuxSize = IIf(Round(((lngAuxSize / 1024&) / 1024&), 2) = CLng(Round(((lngAuxSize / 1024&) / 1024&), 2)), _
                Format(Round(((lngAuxSize / 1024&) / 1024&), 2), "#,##0"), _
                Format(Round(((lngAuxSize / 1024&) / 1024&), 2), "#,##0.0#")) & " Mb"
1430        Else
              ' ** Under 5Mb, display as Kb.
1440          .AuxSize = IIf(Round((lngAuxSize / 1024&), 2) = CLng(Round((lngAuxSize / 1024&), 2)), _
                Format(Round((lngAuxSize / 1024&), 2), "#,##0"), _
                Format(Round((lngAuxSize / 1024&), 2), "#,##0.0#")) & " Kb"
1450        End If

1460        If .chkWarningSeen = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1470          .Messages_lbl.Visible = True
1480          .chkFileSizeWarning.Visible = True
1490          .chkRecordCountWarning.Visible = True
1500        Else
1510          .Messages_lbl.ForeColor = WIN_CLR_DISF
1520          .Messages_lbl_dim_hi.Visible = True
1530        End If

1540        .opgData = 0
1550        .opgCompany = 0

1560        Select Case Parse_Ext(CurrentAppName)  ' ** Module Functions: modFileUtilities.
            Case gstrExt_AppRun
1570          .opgFileType = .opgFileType_optMDE.OptionValue
1580        Case gstrExt_AppDev
1590          .opgFileType = .opgFileType_optMDB.OptionValue
1600        End Select
1610        opgFileType_AfterUpdate  ' ** Procedure: Below.

1620        If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
1630          .Compact_lbl2.Visible = True
1640          .Compact_lbl3.Visible = True
1650          .Compact_lbl4.Visible = True
1660          .opgCompany.Visible = True
1670          .opgCompany_box.Visible = True
1680          .opgCompany.Enabled = False
1690          .opgData.Visible = True
1700          .opgData_box.Visible = True
1710          .opgFileType.Visible = True
1720          .opgFileType_box.Visible = True
1730          .Path_Default.Visible = True
1740          .cmdUseThisPath.Enabled = True
1750          .cmdUseThisPath.Visible = True
1760          .cmdUseThisPath_raised_img.Visible = True
1770          .cmdSetDefPath.Enabled = True
1780          .cmdSetDefPath.Visible = True
1790          .cmdSetDefPath_raised_img.Visible = True
1800        Else
1810          .opgCompany.Visible = False
1820          .opgCompany_box.Visible = False
1830          .opgData.Visible = False
1840          .opgData_box.Visible = False
1850          .opgFileType.Visible = False
1860          .opgFileType_box.Visible = False
1870          .Path_Default.Visible = False
1880          .cmdUseThisPath.Enabled = False
1890          .cmdUseThisPath.Visible = False
1900          .cmdUseThisPath_raised_img.Visible = False
1910          .cmdSetDefPath.Enabled = False
1920          .cmdSetDefPath.Visible = False
1930          .cmdSetDefPath_raised_img.Visible = False
1940        End If
1950        .cmdUseThisPath_raised_semifocus_dots_img.Visible = False
1960        .cmdUseThisPath_raised_focus_img.Visible = False
1970        .cmdUseThisPath_raised_focus_dots_img.Visible = False
1980        .cmdUseThisPath_sunken_focus_dots_img.Visible = False
1990        .cmdUseThisPath_raised_img_dis.Visible = False
2000        .cmdSetDefPath_raised_semifocus_dots_img.Visible = False
2010        .cmdSetDefPath_raised_focus_img.Visible = False
2020        .cmdSetDefPath_raised_focus_dots_img.Visible = False
2030        .cmdSetDefPath_sunken_focus_dots_img.Visible = False
2040        .cmdSetDefPath_raised_img_dis.Visible = False
2050      End Select

2060      intLinkMode = LNK_REFRESH  ' ** Default to refresh.
2070      .cmdRefresh.Enabled = True
2080      .cmdRelink.Enabled = False

2090      blnCompactQuit = False

2100    End With

2110    If Cancel = -1 Then
          ' ** Just for the record.
2120      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
2130    End If

EXITP:
2140    Exit Sub

ERRH:
2150    Select Case ERR.Number
        Case Else
2160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2170    End Select
2180    Resume EXITP

End Sub

Private Sub Form_Load()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

2210    With Me
2220      .DataSize.ForeColor = CLR_DISABLED_FG
2230      .DataSize.BackColor = CLR_DISABLED_BG
2240      .ArchSize.ForeColor = CLR_DISABLED_FG
2250      .ArchSize.BackColor = CLR_DISABLED_BG
2260      .AuxSize.ForeColor = CLR_DISABLED_FG
2270      .AuxSize.BackColor = CLR_DISABLED_BG
2280      .PathFile_Dta.BackColor = CLR_VLTTEAL
2290      .PathFile_Arch.BackColor = CLR_VLTTEAL
2300      .PathFile_Aux.BackColor = CLR_VLTTEAL
2310      chkUnlock_Dta_AfterUpdate  ' ** Procedure: Below.
2320      chkUnlock_Arch_AfterUpdate  ' ** Procedure: Below.
2330      chkUnlock_Aux_AfterUpdate  ' ** Procedure: Below.
2340      .FocusHolder.SetFocus
2350    End With

EXITP:
2360    Exit Sub

ERRH:
2370    Select Case ERR.Number
        Case Else
2380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2390    End Select
2400    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2520    intShiftDown = (Shift And acShiftMask) > 0
2530    intAltDown = (Shift And acAltMask) > 0
2540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt-Shift keys to other forms and functionality:
        ' **   Refresh/Relink:   R {cmdRefresh, cmdRelink}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Alt keys.
2550    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2560      Select Case intRetVal
          Case vbKeyR
2570        With Me
2580          If .cmdRefresh.Enabled = True Then
2590            .cmdRefresh.SetFocus
2600            cmdRefresh_Click  ' ** Procedure: Below.
2610          ElseIf .cmdRelink.Enabled = True Then
2620            .cmdRelink.SetFocus
2630            cmdRelink_Click  ' ** Procedure: Below.
2640          End If
2650        End With
2660      Case vbKeyX
2670        intRetVal = 0
2680        cmdClose_Click  ' ** Procedure: Below.
2690      End Select
2700    End If

        ' ** Ctrl-Shift keys.
2710    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2720      Select Case intRetVal
          Case vbKeyF
2730        With Me
2740          intRetVal = 0
2750          .FocusHolder.SetFocus
2760        End With
2770      Case vbKeyX
2780        intRetVal = 0
2790        blnJustClose = True
2800        SetOption_Dev  ' ** Module Function: modUtilities.
2810        DoCmd.Close acForm, THIS_NAME
2820      End Select
2830    End If

EXITP:
2840    KeyCode = intRetVal
2850    Exit Sub

ERRH:
2860    intRetVal = 0
2870    Select Case ERR.Number
        Case Else
2880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2890    End Select
2900    Resume EXITP

End Sub

Private Sub Form_Timer()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

3010    With Me
3020      .TimerInterval = 0
3030      If blnProgBarOff = True Then
3040        ProgBar_Width_Link False, 1  ' ** Procedure: Below.
3050        blnProgBarOff = False  ' ** This variable is used in ProgBar_Width_Link().
3060      End If
3070    End With

EXITP:
3080    Exit Sub

ERRH:
3090    Select Case ERR.Number
        Case Else
3100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3110    End Select
3120    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3210    With Me
3220      If .chkNoLink = False Then
3230        Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
3240      End If
3250    End With

3260    If blnJustClose = False Then
3270      If gblnMessage = False Then
3280        Select Case strCallingForm
            Case "frmMenu_Utility"
3290          gblnSetFocus = True
3300          If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
3310            DoCmd.SelectObject acForm, strCallingForm, False
3320            Forms(strCallingForm).TimerInterval = 250&
3330          Else
3340            DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
3350          End If
3360        Case vbNullString
3370          SetOption_Dev  ' ** Module Function: modUtilities.
3380        Case Else
3390          If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
3400            DoCmd.SelectObject acForm, strCallingForm, False
3410          Else
3420            DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
3430            DoCmd.SelectObject acForm, strCallingForm, False
3440          End If
3450        End Select
3460      Else
            ' ** This means a Compact has finished and QuitNow() has been called.
3470        gblnMessage = False
3480      End If
3490    End If

EXITP:
3500    Exit Sub

ERRH:
3510    Select Case ERR.Number
        Case 2486  ' ** You can't carry out this action at the present time.
          ' ** Ignore.
3520    Case Else
3530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3540    End Select
3550    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3610    With Me
3620      If .cmdBrowse_raised_focus_dots_img.Visible = True Or .cmdBrowse_raised_focus_img.Visible = True Then
3630        Select Case blnBrowse_Focus
            Case True
3640          .cmdBrowse_raised_semifocus_dots_img.Visible = True
3650          .cmdBrowse_raised_img.Visible = False
3660        Case False
3670          .cmdBrowse_raised_img.Visible = True
3680          .cmdBrowse_raised_semifocus_dots_img.Visible = False
3690        End Select
3700        .cmdBrowse_raised_focus_img.Visible = False
3710        .cmdBrowse_raised_focus_dots_img.Visible = False
3720        .cmdBrowse_sunken_focus_dots_img.Visible = False
3730        .cmdBrowse_raised_img_dis.Visible = False
3740      End If
3750      If cmdUseThisPath_raised_focus_dots_img.Visible = True Or .cmdUseThisPath_raised_focus_img.Visible = True Then
3760        Select Case blnUseThisPath_Focus
            Case True
3770          .cmdUseThisPath_raised_semifocus_dots_img.Visible = True
3780          .cmdUseThisPath_raised_img.Visible = False
3790        Case False
3800          .cmdUseThisPath_raised_img.Visible = True
3810          .cmdUseThisPath_raised_semifocus_dots_img.Visible = False
3820        End Select
3830        .cmdUseThisPath_raised_focus_img.Visible = False
3840        .cmdUseThisPath_raised_focus_dots_img.Visible = False
3850        .cmdUseThisPath_sunken_focus_dots_img.Visible = False
3860        .cmdUseThisPath_raised_img_dis.Visible = False
3870      End If
3880      If .cmdSetDefPath_raised_focus_dots_img.Visible = True Or .cmdSetDefPath_raised_focus_img.Visible = True Then
3890        Select Case blnSetDefPath_Focus
            Case True
3900          .cmdSetDefPath_raised_semifocus_dots_img.Visible = True
3910          .cmdSetDefPath_raised_img.Visible = False
3920        Case False
3930          .cmdSetDefPath_raised_img.Visible = True
3940          .cmdSetDefPath_raised_semifocus_dots_img.Visible = False
3950        End Select
3960        .cmdSetDefPath_raised_focus_img.Visible = False
3970        .cmdSetDefPath_raised_focus_dots_img.Visible = False
3980        .cmdSetDefPath_sunken_focus_dots_img.Visible = False
3990        .cmdSetDefPath_raised_img_dis.Visible = False
4000      End If
4010      .DataSize_lbl.BackStyle = acBackStyleTransparent
4020      .DataSize_lbl.FontBold = False
4030      .ArchSize_lbl.BackStyle = acBackStyleTransparent
4040      .ArchSize_lbl.FontBold = False
4050      .AuxSize_lbl.BackStyle = acBackStyleTransparent
4060      .AuxSize_lbl.FontBold = False
4070    End With

EXITP:
4080    Exit Sub

ERRH:
4090    Select Case ERR.Number
        Case Else
4100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4110    End Select
4120    Resume EXITP

End Sub

Private Sub cmdClose_Click()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim msgResponse As VbMsgBoxResult

4210    With Me
4220      If .PathFile_Dta <> strPathFile_Data Or .PathFile_Arch <> strPathFile_Archive Or .PathFile_Aux <> strPathFile_Auxiliary Then
4230        msgResponse = MsgBox("You have chosen a new path for the data," & vbCrLf & _
              "but have not clicked Relink to make that change." & vbCrLf & vbCrLf & _
              "Are you sure you want to abandon the data path change?", vbQuestion + vbYesNo, "Data Not Relinked")
4240      Else
4250        msgResponse = vbYes
4260      End If
4270      If msgResponse = vbYes Then
4280        DoCmd.Close acForm, THIS_NAME
4290      Else
4300        If .cmdRelink.Enabled = True Then
4310          .cmdRelink.SetFocus
4320        ElseIf .cmdBrowse.Enabled = True Then
4330          .cmdBrowse.SetFocus
4340        Else
              ' ** Stay put.
4350        End If
4360      End If
4370    End With

EXITP:
4380    Exit Sub

ERRH:
4390    Select Case ERR.Number
        Case Else
4400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4410    End Select
4420    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
4590          .cmdCompact.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab
4650        With Me
4660          intRetVal = 0
4670          If .cmdRelink.Enabled = True Then
4680            .cmdRelink.SetFocus
4690          ElseIf .cmdRefresh.Enabled = True Then
4700            .cmdRefresh.SetFocus
4710          ElseIf .cmdBrowse.Enabled = True Then
4720            .cmdBrowse.SetFocus
4730          Else
4740            .cmdCompact.SetFocus
4750          End If
4760        End With
4770      End Select
4780    End If

EXITP:
4790    KeyCode = intRetVal
4800    Exit Sub

ERRH:
4810    intRetVal = 0
4820    Select Case ERR.Number
        Case Else
4830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4840    End Select
4850    Resume EXITP

End Sub

Private Sub cmdCompact_Click()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCompact_Click"

        Dim blnContinue As Boolean

4910    DoCmd.Hourglass True
4920    DoEvents
4930    TableEmpty "tblErrorLog"  ' ** Module Function: modFileUtilities.
4940    TableEmpty "tblXAdmin_Load"  ' ** Module Function: modFileUtilities.
4950    gblnMessage = False
4960    blnContinue = CurrentBackendCompact(False)  ' ** Module Function: modFileUtilities.
4970    If blnContinue = True Then
4980      If strCallingForm = vbNullString And IsNull(Me.OpenArgs) = False Then
4990        strCallingForm = Trim(Me.OpenArgs)
5000      End If
5010      Select Case gblnMessage
          Case True
5020        gblnMessage = False
5030      Case False
5040        cmdClose_Click  ' ** Procedure: Above.
5050      End Select
5060    End If

EXITP:
5070    Exit Sub

ERRH:
5080    DoCmd.Hourglass False
5090    Select Case ERR.Number
        Case Else
5100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5110    End Select
5120    Resume EXITP

End Sub

Private Sub cmdCompact_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCompact_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab
5270        With Me
5280          intRetVal = 0
5290          If .chkFileSizeWarning.Visible = True And .chkFileSizeWarning.Enabled = True Then
5300            .chkFileSizeWarning.SetFocus
5310          ElseIf .chkRecordCountWarning.Visible = True And .chkRecordCountWarning.Enabled = True Then
5320            .chkRecordCountWarning.SetFocus
5330          ElseIf .cmdUseThisPath.Visible = True And .cmdUseThisPath.Enabled = True Then
5340            .cmdUseThisPath.SetFocus
5350          ElseIf .cmdSetDefPath.Visible = True And .cmdSetDefPath.Enabled = True Then
5360            .cmdSetDefPath.SetFocus
5370          Else
5380            .cmdBrowse.SetFocus
5390          End If
5400        End With
5410      End Select
5420    End If

        ' ** Shift keys.
5430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5440      Select Case intRetVal
          Case vbKeyTab
5450        With Me
5460          intRetVal = 0
5470          .cmdClose.SetFocus
5480        End With
5490      End Select
5500    End If

EXITP:
5510    KeyCode = intRetVal
5520    Exit Sub

ERRH:
5530    intRetVal = 0
5540    Select Case ERR.Number
        Case Else
5550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5560    End Select
5570    Resume EXITP

End Sub

Private Sub cmdUseThisPath_GotFocus()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUseThisPath_GotFocus"

5610    With Me
5620      blnUseThisPath_Focus = True
5630      .cmdUseThisPath_raised_semifocus_dots_img.Visible = True
5640      .cmdUseThisPath_raised_img.Visible = False
5650      .cmdUseThisPath_raised_focus_img.Visible = False
5660      .cmdUseThisPath_raised_focus_dots_img.Visible = False
5670      .cmdUseThisPath_sunken_focus_dots_img.Visible = False
5680      .cmdUseThisPath_raised_img_dis.Visible = False
5690    End With

EXITP:
5700    Exit Sub

ERRH:
5710    Select Case ERR.Number
        Case Else
5720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5730    End Select
5740    Resume EXITP

End Sub

Private Sub cmdUseThisPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUseThisPath_MouseDown"

5810    With Me
5820      blnUseThisPath_MouseDown = True
5830      .cmdUseThisPath_sunken_focus_dots_img.Visible = True
5840      .cmdUseThisPath_raised_img.Visible = False
5850      .cmdUseThisPath_raised_semifocus_dots_img.Visible = False
5860      .cmdUseThisPath_raised_focus_img.Visible = False
5870      .cmdUseThisPath_raised_focus_dots_img.Visible = False
5880      .cmdUseThisPath_raised_img_dis.Visible = False
5890    End With

EXITP:
5900    Exit Sub

ERRH:
5910    Select Case ERR.Number
        Case Else
5920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5930    End Select
5940    Resume EXITP

End Sub

Private Sub cmdUseThisPath_Click()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUseThisPath_Click"

        Dim strThisExt As String
        Dim blnContinue As Boolean
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, blnTmp04 As Boolean, blnTmp05 As Boolean, blnTmp06 As Boolean

6010    With Me

6020      DoCmd.Hourglass True
6030      DoEvents

6040      blnContinue = True

6050      strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
6060      blnTmp04 = False: blnTmp05 = False: blnTmp06 = False

          ' ** I know these are all the same, but maybe...
6070      strThisExt = Parse_Ext(CurrentAppName)  ' ** Module Functions: modFileUtilities.
6080      strTmp01 = Parse_Path(.PathFile_This)  ' ** Module Function: modFileUtilities.
6090      Select Case strTmp01
          Case gstrDir_Def
6100        strTmp02 = strTmp01 & LNK_SEP & "Database"
6110        If DirExists(strTmp02) = True Then  ' ** Module Function: modFileUtilities.
6120          strTmp03 = strTmp02 & LNK_SEP & gstrFile_DataName
6130          If .PathFile_Dta <> strTmp03 Then
6140            .PathFile_Dta = strTmp03
6150            .PathFile_Dta.BackColor = CLR_LTRED
6160          Else
6170            blnTmp04 = True
6180          End If
6190          strTmp03 = strTmp02 & LNK_SEP & gstrFile_ArchDataName
6200          If .PathFile_Arch <> strTmp03 Then
6210            .PathFile_Arch = strTmp03
6220            .PathFile_Arch.BackColor = CLR_LTRED
6230          Else
6240            blnTmp05 = True
6250          End If
6260          strTmp03 = strTmp01 & LNK_SEP & gstrFile_AuxDataName
6270          If .PathFile_Aux <> strTmp03 Then
6280            .PathFile_Aux = strTmp03
6290            .PathFile_Aux.BackColor = CLR_LTRED
6300          Else
6310            blnTmp06 = True
6320          End If
6330          If blnTmp04 = False Or blnTmp05 = False Or blnTmp06 = False Then
6340            intLinkMode = LNK_CHANGE
6350          End If
6360        Else
6370          blnContinue = False
6380          DoCmd.Hourglass False
6390          Beep
6400          MsgBox "The Path," & vbCrLf & "  " & strTmp02 & vbCrLf & "does not exist.", vbInformation + vbOKOnly, "Path Not Found"
6410        End If
6420      Case gstrDir_Def64
6430        strTmp02 = strTmp01 & LNK_SEP & "Database"
6440        If DirExists(strTmp02) = True Then  ' ** Module Function: modFileUtilities.
6450          strTmp03 = strTmp02 & LNK_SEP & gstrFile_DataName
6460          If .PathFile_Dta <> strTmp03 Then
6470            .PathFile_Dta = strTmp03
6480            .PathFile_Dta.BackColor = CLR_LTRED
6490          Else
6500            blnTmp04 = True
6510          End If
6520          strTmp03 = strTmp02 & LNK_SEP & gstrFile_ArchDataName
6530          If .PathFile_Arch <> strTmp03 Then
6540            .PathFile_Arch = strTmp03
6550            .PathFile_Arch.BackColor = CLR_LTRED
6560          Else
6570            blnTmp05 = True
6580          End If
6590          strTmp03 = strTmp01 & LNK_SEP & gstrFile_AuxDataName
6600          If .PathFile_Aux <> strTmp03 Then
6610            .PathFile_Aux = strTmp03
6620            .PathFile_Aux.BackColor = CLR_LTRED
6630          Else
6640            blnTmp06 = True
6650          End If
6660          If blnTmp04 = False Or blnTmp05 = False Or blnTmp06 = False Then
6670            intLinkMode = LNK_CHANGE
6680          End If
6690        Else
6700          blnContinue = False
6710          DoCmd.Hourglass False
6720          Beep
6730          MsgBox "The Path," & vbCrLf & "  " & strTmp02 & vbCrLf & "does not exist.", vbInformation + vbOKOnly, "Path Not Found"
6740        End If
6750      Case Else
6760        strTmp02 = strTmp01 & LNK_SEP & "Database"
6770        If DirExists(strTmp02) = True Then  ' ** Module Function: modFileUtilities.
6780          strTmp03 = strTmp02 & LNK_SEP & gstrFile_DataName
6790          If .PathFile_Dta <> strTmp03 Then
6800            .PathFile_Dta = strTmp03
6810            .PathFile_Dta.BackColor = CLR_LTRED
6820          Else
6830            blnTmp04 = True
6840          End If
6850          strTmp03 = strTmp02 & LNK_SEP & gstrFile_ArchDataName
6860          If .PathFile_Arch <> strTmp03 Then
6870            .PathFile_Arch = strTmp03
6880            .PathFile_Arch.BackColor = CLR_LTRED
6890          Else
6900            blnTmp05 = True
6910          End If
6920          strTmp03 = strTmp01 & LNK_SEP & gstrFile_AuxDataName
6930          If .PathFile_Aux <> strTmp03 Then
6940            .PathFile_Aux = strTmp03
6950            .PathFile_Aux.BackColor = CLR_LTRED
6960          Else
6970            blnTmp06 = True
6980          End If
6990          If blnTmp04 = False Or blnTmp05 = False Or blnTmp06 = False Then
7000            intLinkMode = LNK_CHANGE
7010          End If
7020        Else
7030          blnContinue = False
7040          DoCmd.Hourglass False
7050          Beep
7060          MsgBox "The Path," & vbCrLf & "  " & strTmp02 & vbCrLf & "does not exist.", vbInformation + vbOKOnly, "Path Not Found"
7070        End If
7080      End Select  ' ** strTmp01.
7090      If blnTmp04 = True Or blnTmp05 = True Or blnTmp06 = True Then
7100        blnContinue = False
7110        DoCmd.Hourglass False
7120        MsgBox "All paths match. No change required.", vbInformation + vbOKOnly, "Nothing To Do"
7130      Else
7140        If intLinkMode = LNK_CHANGE Then
7150          .cmdRelink.Enabled = True
7160          .cmdRefresh.Enabled = False
7170        End If
7180      End If

7190      DoCmd.Hourglass False

7200    End With

EXITP:
7210    Exit Sub

ERRH:
7220    blnContinue = False
7230    DoCmd.Hourglass False
7240    Select Case ERR.Number
        Case Else
7250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7260    End Select
7270    Resume EXITP

End Sub

Private Sub cmdUseThisPath_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUseThisPath_KeyDown"

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
7390          If .cmdSetDefPath.Enabled = True Then
7400            .cmdSetDefPath.SetFocus
7410          Else
7420            .cmdBrowse.SetFocus
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
7510          If .chkRecordCountWarning.Visible = True And .chkRecordCountWarning.Enabled = True Then
7520            .chkRecordCountWarning.SetFocus
7530          ElseIf .chkFileSizeWarning.Visible = True And .chkFileSizeWarning.Enabled = True Then
7540            .chkFileSizeWarning.SetFocus
7550          Else
7560            .cmdCompact.SetFocus
7570          End If
7580        End With
7590      End Select
7600    End If

EXITP:
7610    KeyCode = intRetVal
7620    Exit Sub

ERRH:
7630    intRetVal = 0
7640    Select Case ERR.Number
        Case Else
7650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7660    End Select
7670    Resume EXITP

End Sub

Private Sub cmdUseThisPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUseThisPath_MouseMove"

7710    With Me
7720      If blnUseThisPath_MouseDown = False Then
7730        Select Case blnUseThisPath_Focus
            Case True
7740          .cmdUseThisPath_raised_focus_dots_img.Visible = True
7750          .cmdUseThisPath_raised_focus_img.Visible = False
7760        Case False
7770          .cmdUseThisPath_raised_focus_img.Visible = True
7780          .cmdUseThisPath_raised_focus_dots_img.Visible = False
7790        End Select
7800        .cmdUseThisPath_raised_img.Visible = False
7810        .cmdUseThisPath_raised_semifocus_dots_img.Visible = False
7820        .cmdUseThisPath_sunken_focus_dots_img.Visible = False
7830        .cmdUseThisPath_raised_img_dis.Visible = False
7840      End If
7850    End With

EXITP:
7860    Exit Sub

ERRH:
7870    Select Case ERR.Number
        Case Else
7880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7890    End Select
7900    Resume EXITP

End Sub

Private Sub cmdUseThisPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUseThisPath_MouseUp"

8010    With Me
8020      .cmdUseThisPath_raised_focus_dots_img.Visible = True
8030      .cmdUseThisPath_raised_img.Visible = False
8040      .cmdUseThisPath_raised_semifocus_dots_img.Visible = False
8050      .cmdUseThisPath_raised_focus_img.Visible = False
8060      .cmdUseThisPath_sunken_focus_dots_img.Visible = False
8070      .cmdUseThisPath_raised_img_dis.Visible = False
8080      blnUseThisPath_MouseDown = False
8090    End With

EXITP:
8100    Exit Sub

ERRH:
8110    Select Case ERR.Number
        Case Else
8120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8130    End Select
8140    Resume EXITP

End Sub

Private Sub cmdUseThisPath_LostFocus()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUseThisPath_LostFocus"

8210    With Me
8220      .cmdUseThisPath_raised_img.Visible = True
8230      .cmdUseThisPath_raised_semifocus_dots_img.Visible = False
8240      .cmdUseThisPath_raised_focus_img.Visible = False
8250      .cmdUseThisPath_raised_focus_dots_img.Visible = False
8260      .cmdUseThisPath_sunken_focus_dots_img.Visible = False
8270      .cmdUseThisPath_raised_img_dis.Visible = False
8280      blnUseThisPath_Focus = False
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

Private Sub cmdSetDefPath_GotFocus()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSetDefPath_GotFocus"

8410    With Me
8420      blnSetDefPath_Focus = True
8430      .cmdSetDefPath_raised_semifocus_dots_img.Visible = True
8440      .cmdSetDefPath_raised_img.Visible = False
8450      .cmdSetDefPath_raised_focus_img.Visible = False
8460      .cmdSetDefPath_raised_focus_dots_img.Visible = False
8470      .cmdSetDefPath_sunken_focus_dots_img.Visible = False
8480      .cmdSetDefPath_raised_img_dis.Visible = False
8490    End With

EXITP:
8500    Exit Sub

ERRH:
8510    Select Case ERR.Number
        Case Else
8520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8530    End Select
8540    Resume EXITP

End Sub

Private Sub cmdSetDefPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSetDefPath_MouseDown"

8610    With Me
8620      blnSetDefPath_MouseDown = True
8630      .cmdSetDefPath_sunken_focus_dots_img.Visible = True
8640      .cmdSetDefPath_raised_img.Visible = False
8650      .cmdSetDefPath_raised_semifocus_dots_img.Visible = False
8660      .cmdSetDefPath_raised_focus_img.Visible = False
8670      .cmdSetDefPath_raised_focus_dots_img.Visible = False
8680      .cmdSetDefPath_raised_img_dis.Visible = False
8690    End With

EXITP:
8700    Exit Sub

ERRH:
8710    Select Case ERR.Number
        Case Else
8720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8730    End Select
8740    Resume EXITP

End Sub

Private Sub cmdSetDefPath_Click()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSetDefPath_Click"

8810    With Me
8820      Application.SetOption "Default Database Directory", Parse_Path(.PathFile_This)  ' ** Module Function: modFileUtilities.
8830      DoEvents
8840      .Path_Default = Application.GetOption("Default Database Directory")
8850      DoEvents
8860      If .Path_Default = Parse_Path(.PathFile_This) Then  ' ** Module Function: modFileUtilities.
8870        .FocusHolder.SetFocus
8880        .cmdSetDefPath.Enabled = False
8890        .cmdSetDefPath_raised_img_dis.Visible = True
8900        .cmdSetDefPath_raised_img.Visible = False
8910        .cmdSetDefPath_raised_semifocus_dots_img.Visible = False
8920        .cmdSetDefPath_raised_focus_img.Visible = False
8930        .cmdSetDefPath_raised_focus_dots_img.Visible = False
8940        .cmdSetDefPath_sunken_focus_dots_img.Visible = False
8950      End If
8960    End With

EXITP:
8970    Exit Sub

ERRH:
8980    Select Case ERR.Number
        Case Else
8990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9000    End Select
9010    Resume EXITP

End Sub

Private Sub cmdSetDefPath_KeyDown(KeyCode As Integer, Shift As Integer)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSetDefPath_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9120    intShiftDown = (Shift And acShiftMask) > 0
9130    intAltDown = (Shift And acAltMask) > 0
9140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9170        With Me
9180          intRetVal = 0
9190          .cmdBrowse.SetFocus
9200        End With
9210      End Select
9220    End If

        ' ** Shift keys.
9230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9250        With Me
9260          intRetVal = 0
9270          If .cmdUseThisPath.Visible = True And .cmdUseThisPath.Enabled = True Then
9280            .cmdUseThisPath.SetFocus
9290          ElseIf .chkRecordCountWarning.Visible = True And .chkRecordCountWarning.Enabled = True Then
9300            .chkRecordCountWarning.SetFocus
9310          ElseIf .chkFileSizeWarning.Visible = True And .chkFileSizeWarning.Enabled = True Then
9320            .chkFileSizeWarning.SetFocus
9330          Else
9340            .cmdCompact.SetFocus
9350          End If
9360        End With
9370      End Select
9380    End If

EXITP:
9390    KeyCode = intRetVal
9400    Exit Sub

ERRH:
9410    intRetVal = 0
9420    Select Case ERR.Number
        Case Else
9430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9440    End Select
9450    Resume EXITP

End Sub

Private Sub cmdSetDefPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSetDefPath_MouseMove"

9510    With Me
9520      If blnSetDefPath_MouseDown = False Then
9530        Select Case blnSetDefPath_Focus
            Case True
9540          .cmdSetDefPath_raised_focus_dots_img.Visible = True
9550          .cmdSetDefPath_raised_focus_img.Visible = False
9560        Case False
9570          .cmdSetDefPath_raised_focus_img.Visible = True
9580          .cmdSetDefPath_raised_focus_dots_img.Visible = False
9590        End Select
9600        .cmdSetDefPath_raised_img.Visible = False
9610        .cmdSetDefPath_raised_semifocus_dots_img.Visible = False
9620        .cmdSetDefPath_sunken_focus_dots_img.Visible = False
9630        .cmdSetDefPath_raised_img_dis.Visible = False
9640      End If
9650    End With

EXITP:
9660    Exit Sub

ERRH:
9670    Select Case ERR.Number
        Case Else
9680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9690    End Select
9700    Resume EXITP

End Sub

Private Sub cmdSetDefPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSetDefPath_MouseUp"

9810    With Me
9820      .cmdSetDefPath_raised_focus_dots_img.Visible = True
9830      .cmdSetDefPath_raised_img.Visible = False
9840      .cmdSetDefPath_raised_semifocus_dots_img.Visible = False
9850      .cmdSetDefPath_raised_focus_img.Visible = False
9860      .cmdSetDefPath_sunken_focus_dots_img.Visible = False
9870      .cmdSetDefPath_raised_img_dis.Visible = False
9880      blnSetDefPath_MouseDown = False
9890    End With

EXITP:
9900    Exit Sub

ERRH:
9910    Select Case ERR.Number
        Case Else
9920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9930    End Select
9940    Resume EXITP

End Sub

Private Sub cmdSetDefPath_LostFocus()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSetDefPath_LostFocus"

10010   With Me
10020     .cmdSetDefPath_raised_img.Visible = True
10030     .cmdSetDefPath_raised_semifocus_dots_img.Visible = False
10040     .cmdSetDefPath_raised_focus_img.Visible = False
10050     .cmdSetDefPath_raised_focus_dots_img.Visible = False
10060     .cmdSetDefPath_sunken_focus_dots_img.Visible = False
10070     .cmdSetDefPath_raised_img_dis.Visible = False
10080     blnSetDefPath_Focus = False
10090   End With

EXITP:
10100   Exit Sub

ERRH:
10110   Select Case ERR.Number
        Case Else
10120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10130   End Select
10140   Resume EXITP

End Sub

Private Sub cmdRefresh_Click()

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRefresh_Click"

        Dim dbs As DAO.Database, tdf As DAO.TableDef
        Dim blnRetVal As Boolean
        Dim varTmp00 As Variant

10210   With Me
10220     If intLinkMode = LNK_REFRESH Then

10230       blnRetVal = True

10240       DoCmd.Hourglass True
10250       DoEvents

10260       varTmp00 = DLookup("[seclic_auxloc]", "tblSecurity_License")
10270       Select Case IsNull(varTmp00)
            Case True
10280         blnAuxLoc = False
10290       Case False
10300         Select Case varTmp00
              Case True
10310           blnAuxLoc = True
10320         Case False
10330           blnAuxLoc = False
10340         End Select
10350       End Select

10360       DoCmd.Hourglass True
10370       .ProgBar_box.Visible = True
10380       .ProgBar_box2.Visible = True
10390       ProgBar_Width_Link 0#, 2  ' ** Procedure: Below.
10400       ProgBar_Width_Link True, 1  ' ** Procedure: Below.
10410       DoEvents

10420       Set dbs = CurrentDb
10430       With dbs
              ' ****************************************
              ' ** Progress Bar setup.
              ' ****************************************
10440         Set frmPBar = Me
10450         With frmPBar
10460           Set ctlPBox = .ProgBar_box
10470           dblPBar_MaxWidth = ctlPBox.Width
10480           If ctlPBox.Visible = False Then ctlPBox.Visible = True
10490           ProgBar_Width_Link True, 1  ' ** Procedure: Below.
10500         End With
10510         DoEvents
10520         dblPBar_Steps = .TableDefs.Count
10530         dblPBar_Increment = (dblPBar_MaxWidth / dblPBar_Steps)
10540         dblPBar_CurWidth = 0#
10550         dblPBar_ThisIncrement = 0#
              ' ****************************************
10560         For Each tdf In .TableDefs
10570           With tdf
10580             If .Connect <> vbNullString Then
10590 On Error Resume Next
10600               .RefreshLink
10610               If ERR.Number <> 0 Then
10620                 blnRetVal = False
10630                 DoCmd.Hourglass False
10640                 Beep
10650                 MsgBox "Error: " & CStr(ERR.Number) & vbCrLf & "Description: " & ERR.description & vbCrLf & vbCrLf & _
                        "The refresh has failed.", vbCritical + vbOKOnly, "Refresh Failed"
10660                 Exit For
10670               Else
10680 On Error GoTo ERRH
10690               End If
10700             End If
10710           End With
                ' ******************************
10720           dblPBar_ThisStep = dblPBar_ThisStep + 1#
10730           dblPBar_ThisIncrement = dblPBar_Increment
10740           dblPBar_CurWidth = (dblPBar_CurWidth + dblPBar_ThisIncrement)
10750           ProgBar_Width_Link dblPBar_CurWidth, 2  ' ** Procedure: Below.
10760           DoEvents
                ' ******************************
10770         Next
10780         .Close
10790       End With

10800       DoCmd.Hourglass False
10810       If blnRetVal = True Then
10820 On Error Resume Next
              ' ** If the form's already closed, then it'll error.
10830         strPathFile_Data = .PathFile_Dta
10840         If ERR.Number = 0 Then
10850           strPathFile_Archive = .PathFile_Arch
10860           strPathFile_Auxiliary = .PathFile_Aux
10870           MsgBox "Refresh successful.", vbInformation + vbOKOnly, ("Finished" & Space(40))
10880           .chkNoLink = False  ' ** Somehow this errored once, as if it's running code while the form isn't open!
10890 On Error GoTo ERRH
10900         Else
10910 On Error GoTo ERRH
10920         End If
10930         .cmdClose.SetFocus
10940       End If  ' ** blnRetVal.

10950       If strCallingForm <> vbNullString Then
              ' ** This seems silly! How could it be here with a Forms.Count of Zero?
10960         DoCmd.Hourglass True
10970         DoEvents
10980         If Forms.Count > 0 Then
10990           If Forms(0).Name <> THIS_NAME And Forms(0).Name <> "frmMenu_Utility" Then
11000             DoCmd.Close acForm, Forms(0).Name
11010           End If
11020           If Forms.Count = 0 Then
11030             gblnSetFocus = True
11040             gstrReturningForm = THIS_NAME
11050             DoCmd.OpenForm "frmMenu_Utility"
11060             DoCmd.OpenForm THIS_NAME
11070             DoCmd.SelectObject acForm, THIS_NAME, False
11080           End If
11090         Else
11100           gblnSetFocus = True
11110           gstrReturningForm = THIS_NAME
11120           DoCmd.OpenForm "frmMenu_Utility"
11130           DoCmd.OpenForm THIS_NAME
11140           DoCmd.SelectObject acForm, THIS_NAME, False
11150         End If
11160         DoCmd.Hourglass False
11170       End If

11180       blnProgBarOff = True
11190       .TimerInterval = 2000&

11200     End If  ' ** intLinkMode.
11210   End With

EXITP:
11220   Set frmPBar = Nothing
11230   Set ctlPBox = Nothing
11240   Set tdf = Nothing
11250   Set dbs = Nothing
11260   Exit Sub

ERRH:
11270   DoCmd.Hourglass False
11280   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Just ignore.
11290   Case Else
11300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11310   End Select
11320   Resume EXITP

End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRefresh_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11420   intShiftDown = (Shift And acShiftMask) > 0
11430   intAltDown = (Shift And acAltMask) > 0
11440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11460     Select Case intRetVal
          Case vbKeyTab
11470       With Me
11480         intRetVal = 0
11490         .cmdClose.SetFocus
11500       End With
11510     End Select
11520   End If

        ' ** Shift keys.
11530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11540     Select Case intRetVal
          Case vbKeyTab
11550       With Me
11560         intRetVal = 0
11570         .cmdBrowse.SetFocus
11580       End With
11590     End Select
11600   End If

        ' ** Ctrl-Shift keys.
11610   If intCtrlDown And (Not intAltDown) And intShiftDown Then
11620     Select Case intRetVal
          Case vbKeyTab
11630       With Me
11640         intRetVal = 0
11650         .chkUnlock_Aux.SetFocus
11660       End With
11670     End Select
11680   End If

EXITP:
11690   KeyCode = intRetVal
11700   Exit Sub

ERRH:
11710   intRetVal = 0
11720   Select Case ERR.Number
        Case Else
11730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11740   End Select
11750   Resume EXITP

End Sub

Private Sub cmdRelink_Click()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRelink_Click"

        Dim dbs As DAO.Database, tdf As DAO.TableDef
        Dim strPath As String, strAuxPath As String
        Dim strPathFile_Dta As String, strPathFile_Arc As String, strPathFile_Aux As String
        Dim blnRetVal As Boolean
        Dim varTmp00 As Variant

11810   With Me
11820     If intLinkMode = LNK_CHANGE Then

11830       blnRetVal = True

11840       varTmp00 = DLookup("[seclic_auxloc]", "tblSecurity_License")
11850       Select Case IsNull(varTmp00)
            Case True
11860         blnAuxLoc = False
11870       Case False
11880         Select Case varTmp00
              Case True
11890           blnAuxLoc = True
11900         Case False
11910           blnAuxLoc = False
11920         End Select
11930       End Select

            ' ** I don't see how opgData ever gets a value other than Zero!
            'If .opgData > 0 Then
            '  intLinkMode = LNK_CHANGE
            'Else
            '  If intLinkMode = 0 Then  ' ** intLinkMode only changes if they Browse!
            '    intLinkMode = LNK_REFRESH
            '  End If
            'End If

11940       DoCmd.Hourglass True
11950       .ProgBar_box.Visible = True
11960       .ProgBar_box2.Visible = True
11970       ProgBar_Width_Link 0#, 2  ' ** Procedure: Below.
11980       ProgBar_Width_Link True, 1  ' ** Procedure: Below.
11990       DoEvents

12000       If .opgData.Visible = True And .opgCompany.Visible = True And .opgCompany.Enabled = True Then
12010         If .opgData = .opgData_optDemo.OptionValue And .opgCompany > 0 Then
12020           strPathFile_Dta = .PathFile_Dta
12030           strPathFile_Arc = .PathFile_Arch
12040           strPathFile_Aux = .PathFile_Aux
12050           strPath = Parse_Path(strPathFile_Dta)  ' ** Module Function: modFileUtilities.
12060           strAuxPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
12070           If blnAuxLoc = False Then strAuxPath = strPath
12080           If FileExists(strPath & LNK_SEP & gstrFile_DataLockfile) = True Or _
                    FileExists(strPath & LNK_SEP & gstrFile_ArchDataLockfile) = True Or _
                    FileExists(strAuxPath & LNK_SEP & gstrFile_AuxDataLockfile) = True Then  ' ** Module Function: modFileUtilities.
12090             blnRetVal = False
12100             DoCmd.Hourglass False
12110             Beep
12120             MsgBox "One of the data files is locked." & vbCrLf & vbCrLf & _
                    "Exit Trust Accountant, then log in and try again.", vbCritical + vbOKOnly, "Data File Locked"
12130           Else
12140             Select Case .opgCompany
                  Case .opgCompany_optNorthFork.OptionValue
12150               If FileExists(strPath & LNK_SEP & gstrFile_DevNorthForkDta) = True Then
12160                 Kill strPathFile_Dta
12170                 FileCopy strPath & LNK_SEP & gstrFile_DevNorthForkDta, strPathFile_Dta
12180                 Kill strPathFile_Arc
12190                 FileCopy strPath & LNK_SEP & gstrFile_DevNorthForkArch, strPathFile_Arc
12200                 DoEvents
12210               Else
12220                 blnRetVal = False
12230                 DoCmd.Hourglass False
12240                 Beep
12250                 MsgBox "The chosen Demo database doesn't exist", vbCritical + vbOKOnly, "File Not Found"
12260               End If
12270             Case .opgCompany_optHintonCPA.OptionValue
12280               If FileExists(strPath & LNK_SEP & gstrFile_DevHintonDta) = True Then
12290                 Kill strPathFile_Dta
12300                 FileCopy strPath & LNK_SEP & gstrFile_DevHintonDta, strPathFile_Dta
12310                 Kill strPathFile_Arc
12320                 FileCopy strPath & LNK_SEP & gstrFile_DevHintonArch, strPathFile_Arc
12330                 DoEvents
12340               Else
12350                 blnRetVal = False
12360                 DoCmd.Hourglass False
12370                 Beep
12380                 MsgBox "The chosen Demo database doesn't exist", vbCritical + vbOKOnly, "File Not Found"
12390               End If
12400             End Select
12410           End If
12420         End If
12430       End If

12440       If blnRetVal = True Then
12450         gstrTrustDataLocation = Parse_Path(.PathFile_Dta) & LNK_SEP  ' ** Module Function: modFileUtilities.
12460         gstrTrustAuxLocation = Parse_Path(.PathFile_Aux) & LNK_SEP   ' ** Module Function: modFileUtilities.
12470         blnRetVal = IniFile_Set("Files", "Location", gstrTrustDataLocation, CurrentAppPath & LNK_SEP & gstrFile_INI)  ' ** Module Function: modStartupFuncs, modFileUtilities.
12480         If blnRetVal = True Then
12490           IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
12500           If gstrTrustDataLocation <> Parse_Path(.PathFile_Dta) & LNK_SEP Then
12510             blnRetVal = False
12520             DoCmd.Hourglass False
12530             MsgBox "Relinking failed.", vbExclamation + vbOKOnly, "Relink Failed"
12540           Else
                  ' ** InitializeTables() will handle progress bar.
12550             blnRetVal = InitializeTables  ' ** Module Function: modStartupFuncs.
12560             If blnRetVal = False Then
12570               DoCmd.Hourglass False
12580               MsgBox "Relinking failed.", vbExclamation + vbOKOnly, "Relink Failed"
12590             End If
12600           End If
12610         Else
12620           DoCmd.Hourglass False
12630           MsgBox "Relinking failed.", vbExclamation + vbOKOnly, "Relink Failed"
12640         End If
12650       End If

12660       DoCmd.Hourglass False
12670       If blnRetVal = True Then
12680 On Error Resume Next
              ' ** If the form's already closed, then it'll error.
12690         strPathFile_Data = .PathFile_Dta
12700         If ERR.Number = 0 Then
12710           strPathFile_Archive = .PathFile_Arch
12720           strPathFile_Auxiliary = .PathFile_Aux
12730           MsgBox "Relink successful.", vbInformation + vbOKOnly, ("Finished" & Space(40))
12740           .chkNoLink = False  ' ** Somehow this errored once, as if it's running code while the form isn't open!
12750 On Error GoTo ERRH
12760           .cmdClose.SetFocus
12770           DoEvents
12780           .cmdRefresh.Enabled = True
12790           .cmdRelink.Enabled = False
12800         Else
12810 On Error GoTo ERRH
12820           .cmdClose.SetFocus
12830         End If
12840       End If  ' **  blnRetVal.

12850       If strCallingForm <> vbNullString Then
              ' ** This seems silly! How could it be here with a Forms.Count of Zero?
12860         DoCmd.Hourglass True
12870         DoEvents
12880         If Forms.Count > 0 Then
12890           If Forms(0).Name <> THIS_NAME And Forms(0).Name <> "frmMenu_Utility" Then
12900             DoCmd.Close acForm, Forms(0).Name
12910           End If
12920           If Forms.Count = 0 Then
12930             gblnSetFocus = True
12940             gstrReturningForm = THIS_NAME
12950             DoCmd.OpenForm "frmMenu_Utility"
12960             DoCmd.OpenForm THIS_NAME
12970             DoCmd.SelectObject acForm, THIS_NAME, False
12980           End If
12990         Else
13000           gblnSetFocus = True
13010           gstrReturningForm = THIS_NAME
13020           DoCmd.OpenForm "frmMenu_Utility"
13030           DoCmd.OpenForm THIS_NAME
13040           DoCmd.SelectObject acForm, THIS_NAME, False
13050         End If
13060         DoCmd.Hourglass False
13070       End If

13080       blnProgBarOff = True
13090       .TimerInterval = 2000&

13100     End If  ' ** intLinkMode.
13110   End With

EXITP:
13120   Set frmPBar = Nothing
13130   Set ctlPBox = Nothing
13140   Set tdf = Nothing
13150   Set dbs = Nothing
13160   Exit Sub

ERRH:
13170   DoCmd.Hourglass False
13180   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Just ignore.
13190   Case Else
13200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13210   End Select
13220   Resume EXITP

End Sub

Private Sub cmdRelink_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRelink_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13320   intShiftDown = (Shift And acShiftMask) > 0
13330   intAltDown = (Shift And acAltMask) > 0
13340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13360     Select Case intRetVal
          Case vbKeyTab
13370       With Me
13380         intRetVal = 0
13390         .cmdClose.SetFocus
13400       End With
13410     End Select
13420   End If

        ' ** Shift keys.
13430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13440     Select Case intRetVal
          Case vbKeyTab
13450       With Me
13460         intRetVal = 0
13470         .cmdBrowse.SetFocus
13480       End With
13490     End Select
13500   End If

        ' ** Ctrl-Shift keys.
13510   If intCtrlDown And (Not intAltDown) And intShiftDown Then
13520     Select Case intRetVal
          Case vbKeyTab
13530       With Me
13540         intRetVal = 0
13550         .chkUnlock_Aux.SetFocus
13560       End With
13570     End Select
13580   End If

EXITP:
13590   KeyCode = intRetVal
13600   Exit Sub

ERRH:
13610   intRetVal = 0
13620   Select Case ERR.Number
        Case Else
13630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13640   End Select
13650   Resume EXITP

End Sub

Private Sub chkFileSizeWarning_AfterUpdate()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkFileSizeWarning_AfterUpdate"

13710   With Me
13720     Select Case .chkFileSizeWarning
          Case True
13730       .chkFileSizeWarning_lbl.FontBold = True
13740     Case False
13750       .chkFileSizeWarning_lbl.FontBold = False
13760     End Select
13770   End With

EXITP:
13780   Exit Sub

ERRH:
13790   Select Case ERR.Number
        Case Else
13800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13810   End Select
13820   Resume EXITP

End Sub

Private Sub chkFileSizeWarning_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkFileSizeWarning_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13920   intShiftDown = (Shift And acShiftMask) > 0
13930   intAltDown = (Shift And acAltMask) > 0
13940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13970       With Me
13980         intRetVal = 0
13990         If .chkRecordCountWarning.Visible = True And .chkRecordCountWarning.Enabled = True Then
14000           .chkRecordCountWarning.SetFocus
14010         ElseIf .cmdUseThisPath.Visible = True And .cmdUseThisPath.Enabled = True Then
14020           .cmdUseThisPath.SetFocus
14030         ElseIf .cmdSetDefPath.Visible = True And .cmdSetDefPath.Enabled = True Then
14040           .cmdSetDefPath.SetFocus
14050         Else
14060           .cmdBrowse.SetFocus
14070         End If
14080       End With
14090     End Select
14100   End If

        ' ** Shift keys.
14110   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14130       With Me
14140         intRetVal = 0
14150         .cmdCompact.SetFocus
14160       End With
14170     End Select
14180   End If

EXITP:
14190   KeyCode = intRetVal
14200   Exit Sub

ERRH:
14210   intRetVal = 0
14220   Select Case ERR.Number
        Case Else
14230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14240   End Select
14250   Resume EXITP

End Sub

Private Sub chkRecordCountWarning_AfterUpdate()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRecordCountWarning_AfterUpdate"

14310   With Me
14320     Select Case .chkRecordCountWarning
          Case True
14330       .chkRecordCountWarning_lbl.FontBold = True
14340     Case False
14350       .chkRecordCountWarning_lbl.FontBold = False
14360     End Select
14370   End With

EXITP:
14380   Exit Sub

ERRH:
14390   Select Case ERR.Number
        Case Else
14400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14410   End Select
14420   Resume EXITP

End Sub

Private Sub chkRecordCountWarning_KeyDown(KeyCode As Integer, Shift As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRecordCountWarning_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
14570       With Me
14580         intRetVal = 0
14590         If .cmdUseThisPath.Visible = True And .cmdUseThisPath.Enabled = True Then
14600           .cmdUseThisPath.SetFocus
14610         ElseIf .cmdSetDefPath.Visible = True And .cmdSetDefPath.Enabled = True Then
14620           .cmdSetDefPath.SetFocus
14630         Else
14640           .cmdBrowse.SetFocus
14650         End If
14660       End With
14670     End Select
14680   End If

        ' ** Shift keys.
14690   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14700     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14710       With Me
14720         intRetVal = 0
14730         If .chkFileSizeWarning.Visible = True And .chkFileSizeWarning.Enabled = True Then
14740           .chkFileSizeWarning.SetFocus
14750         Else
14760           .cmdCompact.SetFocus
14770         End If
14780       End With
14790     End Select
14800   End If

EXITP:
14810   KeyCode = intRetVal
14820   Exit Sub

ERRH:
14830   intRetVal = 0
14840   Select Case ERR.Number
        Case Else
14850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14860   End Select
14870   Resume EXITP

End Sub

Private Sub chkAuxLoc_AfterUpdate()

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAuxLoc_AfterUpdate"

14910   With Me
14920     Select Case .chkAuxLoc
          Case True
14930       .chkAuxLoc_lbl.FontBold = True
14940     Case False
14950       .chkAuxLoc_lbl.FontBold = False
14960     End Select
14970   End With

EXITP:
14980   Exit Sub

ERRH:
14990   Select Case ERR.Number
        Case Else
15000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15010   End Select
15020   Resume EXITP

End Sub

Private Sub PathFile_This_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_This_box_MouseMove"

15110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15120   Exit Sub

ERRH:
15130   Select Case ERR.Number
        Case Else
15140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15150   End Select
15160   Resume EXITP

End Sub

Private Sub PathFile_This_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_This_MouseMove"

15210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15220   Exit Sub

ERRH:
15230   Select Case ERR.Number
        Case Else
15240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15250   End Select
15260   Resume EXITP

End Sub

Private Sub opgFileType_AfterUpdate()

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFileType_AfterUpdate"

15310   With Me
15320     Select Case .opgFileType
          Case .opgFileType_optMDE.OptionValue
15330       .opgFileType_optMDE_lbl.FontBold = True
15340       .opgFileType_optMDB_lbl.FontBold = False
15350     Case .opgFileType_optMDB.OptionValue
15360       .opgFileType_optMDE_lbl.FontBold = False
15370       .opgFileType_optMDB_lbl.FontBold = True
15380     End Select
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

Private Sub opgFileType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFileType_MouseMove"

15510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15520   Exit Sub

ERRH:
15530   Select Case ERR.Number
        Case Else
15540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15550   End Select
15560   Resume EXITP

End Sub

Private Sub opgFileType_optMDE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFileType_optMDE_MouseMove"

15610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15620   Exit Sub

ERRH:
15630   Select Case ERR.Number
        Case Else
15640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15650   End Select
15660   Resume EXITP

End Sub

Private Sub opgFileType_optMDB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgFileType_optMDB_MouseMove"

15710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15720   Exit Sub

ERRH:
15730   Select Case ERR.Number
        Case Else
15740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15750   End Select
15760   Resume EXITP

End Sub

Private Sub Path_Default_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "Path_Default_MouseMove"

15810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
15820   Exit Sub

ERRH:
15830   Select Case ERR.Number
        Case Else
15840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15850   End Select
15860   Resume EXITP

End Sub

Private Sub PathFile_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_lbl_MouseMove"

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

Private Sub PathFile_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_box_MouseMove"

16010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16020   Exit Sub

ERRH:
16030   Select Case ERR.Number
        Case Else
16040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16050   End Select
16060   Resume EXITP

End Sub

Private Sub opgData_AfterUpdate()

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgData_AfterUpdate"

        Dim strPath As String, strAuxPath As String, strPathFile_Dta As String, strPathFile_Arc As String, strPathFile_Aux As String
        Dim varTmp00 As Variant

16110   With Me

16120     strPath = vbNullString: strAuxPath = vbNullString
16130     strPathFile_Dta = vbNullString: strPathFile_Arc = vbNullString: strPathFile_Aux = vbNullString

16140     varTmp00 = DLookup("[seclic_auxloc]", "tblSecurity_License")
16150     Select Case IsNull(varTmp00)
          Case True
16160       blnAuxLoc = False
16170     Case False
16180       Select Case varTmp00
            Case True
16190         blnAuxLoc = True
16200       Case False
16210         blnAuxLoc = False
16220       End Select
16230     End Select

16240     Select Case .opgData
          Case .opgData_optEmpty.OptionValue
16250       .opgData_optEmpty_lbl.FontBold = True
16260       .opgData_optDemo_lbl.FontBold = False
16270       If .opgCompany.Enabled = True Then
16280         .opgCompany.Enabled = False
16290       End If
16300       strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
16310       strAuxPath = strPath
16320       If DirExists(strPath & LNK_SEP & gstrDir_DevEmpty) = True Then  ' ** Module Function: modFileUtilities.
16330         strPath = strPath & LNK_SEP & gstrDir_DevEmpty
16340       Else
16350         If DirExists(strPath & LNK_SEP & "Database") = True Then  ' ** Module Function: modFileUtilities.
16360           If DirExists(strPath & LNK_SEP & "Database" & LNK_SEP & gstrDir_DevEmpty) = True Then  ' ** Module Function: modFileUtilities.
16370             strPath = strPath & LNK_SEP & "Database" & LNK_SEP & gstrDir_DevEmpty
16380           Else
16390             strPath = vbNullString
16400             Beep
16410             MsgBox "The EmptyDatabase path does not exist.", vbExclamation + vbOKOnly, "Path Not Found"
16420           End If
16430         Else
16440           strPath = vbNullString
16450           Beep
16460           MsgBox "The EmptyDatabase path does not exist.", vbExclamation + vbOKOnly, "Path Not Found"
16470         End If
16480       End If
16490       If blnAuxLoc = False Then strAuxPath = strPath
16500     Case .opgData_optDemo.OptionValue
16510       .opgData_optEmpty_lbl.FontBold = False
16520       .opgData_optDemo_lbl.FontBold = True
16530       If .opgCompany.Enabled = False Then
16540         .opgCompany.Enabled = True
16550       End If
16560       strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
16570       strAuxPath = strPath
16580       If DirExists(strPath & LNK_SEP & gstrDir_DevDemo) = True Then  ' ** Module Function: modFileUtilities.
16590         strPath = strPath & LNK_SEP & gstrDir_DevDemo
16600       Else
16610         If DirExists(strPath & LNK_SEP & "Database") = True Then  ' ** Module Function: modFileUtilities.
16620           If DirExists(strPath & LNK_SEP & "Database" & LNK_SEP & gstrDir_DevDemo) = True Then  ' ** Module Function: modFileUtilities.
16630             strPath = strPath & LNK_SEP & "Database" & LNK_SEP & gstrDir_DevDemo
16640           Else
16650             If FileExists(strPath & LNK_SEP & "Database" & LNK_SEP & gstrFile_DevNorthForkDta) = True And _
                      FileExists(strPath & LNK_SEP & "Database" & LNK_SEP & gstrFile_DevHintonDta) = True Then  ' ** Module Function: modFileUtilities.
16660               strPath = strPath & LNK_SEP & "Database"
16670             Else
16680               strPath = vbNullString
16690               Beep
16700               MsgBox "The DemoDatabase path does not exist.", vbExclamation + vbOKOnly, "Path Not Found"
16710             End If
16720           End If
16730         Else
16740           strPath = vbNullString
16750           Beep
16760           MsgBox "The DemoDatabase path does not exist.", vbExclamation + vbOKOnly, "Path Not Found"
16770         End If
16780       End If
16790       If blnAuxLoc = False Then strAuxPath = strPath
16800     End Select

          ' ** I HAVE TO SAY, THIS PROCEDURE GOT WAY OUT OF HAND!!
16810     If strPath <> vbNullString Then
16820       strPathFile_Dta = strPath & LNK_SEP & gstrFile_DataName
16830       .PathFile_Dta = strPathFile_Dta
16840       strPathFile_Arc = strPath & LNK_SEP & gstrFile_ArchDataName
16850       .PathFile_Arch = strPathFile_Arc
16860       strPathFile_Aux = strAuxPath & LNK_SEP & gstrFile_AuxDataName
16870       .PathFile_Aux = strPathFile_Aux
16880       DoEvents
16890       Select Case .opgData
            Case .opgData_optEmpty.OptionValue
              ' ** Nothing more to do.
16900       Case .opgData_optDemo.OptionValue
16910         Select Case .opgCompany
              Case 0
                ' ** Nothing more to do.
16920         Case .opgCompany_optNorthFork.OptionValue
16930           If FileExists(strPath & LNK_SEP & gstrFile_DevNorthForkDta) = True Then  ' ** Module Function: modFileUtilities.
16940             If FileExists(strPath & LNK_SEP & gstrFile_DataLockfile) = True Or _
                      FileExists(strPath & LNK_SEP & gstrFile_ArchDataLockfile) = True Or _
                      FileExists(strAuxPath & LNK_SEP & gstrFile_AuxDataLockfile) = True Then  ' ** Module Function: modFileUtilities.
16950               .opgCompany = 0
16960               .opgCompany_optNorthFork_lbl.FontBold = False
16970               .opgCompany_optHintonCPA_lbl.FontBold = False
16980             Else
                    ' ** Do nothing until Relink is clicked!
16990             End If
17000           Else
17010             .opgCompany = 0
17020             .opgCompany_optNorthFork_lbl.FontBold = False
17030             .opgCompany_optHintonCPA_lbl.FontBold = False
17040           End If
17050         Case .opgCompany_optHintonCPA.OptionValue
17060           If FileExists(strPath & LNK_SEP & gstrFile_DevHintonDta) = True Then  ' ** Module Function: modFileUtilities.
17070             If FileExists(strPath & LNK_SEP & gstrFile_DataLockfile) = True Or _
                      FileExists(strPath & LNK_SEP & gstrFile_ArchDataLockfile) = True Or _
                      FileExists(strAuxPath & LNK_SEP & gstrFile_AuxDataLockfile) = True Then  ' ** Module Function: modFileUtilities.
17080               .opgCompany = 0
17090               .opgCompany_optNorthFork_lbl.FontBold = False
17100               .opgCompany_optHintonCPA_lbl.FontBold = False
17110             Else
                    ' ** Do nothing until Relink is clicked!
17120             End If
17130           Else
17140             .opgCompany = 0
17150             .opgCompany_optNorthFork_lbl.FontBold = False
17160             .opgCompany_optHintonCPA_lbl.FontBold = False
17170           End If
17180         End Select
17190       End Select
17200     Else
17210       .opgData = 0
17220       .opgData_optEmpty_lbl.FontBold = False
17230       .opgData_optDemo_lbl.FontBold = False
17240     End If

17250   End With

EXITP:
17260   Exit Sub

ERRH:
17270   Select Case ERR.Number
        Case Else
17280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17290   End Select
17300   Resume EXITP

End Sub

Private Sub opgData_optEmpty_KeyDown(KeyCode As Integer, Shift As Integer)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgData_optEmpty_KeyDown"

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
17490         If .opgCompany.Enabled = True Then
17500           .opgCompany.SetFocus
17510         Else
17520           .cmdBrowse.SetFocus
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
17610         If .cmdSetDefPath.Visible = True And .cmdSetDefPath.Enabled = True Then
17620           .cmdSetDefPath.SetFocus
17630         ElseIf .cmdUseThisPath.Visible = True And .cmdUseThisPath.Enabled = True Then
17640           .cmdUseThisPath.SetFocus
17650         ElseIf .chkRecordCountWarning.Visible = True And .chkRecordCountWarning.Enabled = True Then
17660           .chkRecordCountWarning.SetFocus
17670         ElseIf .chkFileSizeWarning.Visible = True And .chkFileSizeWarning.Enabled = True Then
17680           .chkFileSizeWarning.SetFocus
17690         Else
17700           .cmdCompact.SetFocus
17710         End If
17720       End With
17730     End Select
17740   End If

EXITP:
17750   KeyCode = intRetVal
17760   Exit Sub

ERRH:
17770   intRetVal = 0
17780   Select Case ERR.Number
        Case Else
17790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17800   End Select
17810   Resume EXITP

End Sub

Private Sub opgData_optDemo_KeyDown(KeyCode As Integer, Shift As Integer)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgData_optDemo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17920   intShiftDown = (Shift And acShiftMask) > 0
17930   intAltDown = (Shift And acAltMask) > 0
17940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17970       With Me
17980         intRetVal = 0
17990         If .opgCompany.Enabled = True Then
18000           .opgCompany.SetFocus
18010         Else
18020           .cmdBrowse.SetFocus
18030         End If
18040       End With
18050     End Select
18060   End If

        ' ** Shift keys.
18070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18090       With Me
18100         intRetVal = 0
18110         If .cmdSetDefPath.Visible = True And .cmdSetDefPath.Enabled = True Then
18120           .cmdSetDefPath.SetFocus
18130         ElseIf .cmdUseThisPath.Visible = True And .cmdUseThisPath.Enabled = True Then
18140           .cmdUseThisPath.SetFocus
18150         ElseIf .chkRecordCountWarning.Visible = True And .chkRecordCountWarning.Enabled = True Then
18160           .chkRecordCountWarning.SetFocus
18170         ElseIf .chkFileSizeWarning.Visible = True And .chkFileSizeWarning.Enabled = True Then
18180           .chkFileSizeWarning.SetFocus
18190         Else
18200           .cmdCompact.SetFocus
18210         End If
18220       End With
18230     End Select
18240   End If

EXITP:
18250   KeyCode = intRetVal
18260   Exit Sub

ERRH:
18270   intRetVal = 0
18280   Select Case ERR.Number
        Case Else
18290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18300   End Select
18310   Resume EXITP

End Sub

Private Sub opgCompany_AfterUpdate()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCompany_AfterUpdate"

18410   With Me
18420     Select Case .opgCompany
          Case .opgCompany_optNorthFork.OptionValue
18430       .opgCompany_optNorthFork_lbl.FontBold = True
18440       .opgCompany_optHintonCPA_lbl.FontBold = False
18450     Case .opgCompany_optHintonCPA.OptionValue
18460       .opgCompany_optNorthFork_lbl.FontBold = False
18470       .opgCompany_optHintonCPA_lbl.FontBold = True
18480     End Select
18490   End With

EXITP:
18500   Exit Sub

ERRH:
18510   Select Case ERR.Number
        Case Else
18520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18530   End Select
18540   Resume EXITP

End Sub

Private Sub opgCompany_optNorthFork_KeyDown(KeyCode As Integer, Shift As Integer)

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCompany_optNorthFork_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18620   intShiftDown = (Shift And acShiftMask) > 0
18630   intAltDown = (Shift And acAltMask) > 0
18640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18670       With Me
18680         intRetVal = 0
18690         .cmdBrowse.SetFocus
18700       End With
18710     End Select
18720   End If

        ' ** Shift keys.
18730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18750       With Me
18760         intRetVal = 0
18770         If .opgData.Visible = True And .opgData.Enabled = True Then
18780           .opgData.SetFocus
18790         ElseIf .cmdSetDefPath.Visible = True And .cmdSetDefPath.Enabled = True Then
18800           .cmdSetDefPath.SetFocus
18810         ElseIf .cmdUseThisPath.Visible = True And .cmdUseThisPath.Enabled = True Then
18820           .cmdUseThisPath.SetFocus
18830         ElseIf .chkRecordCountWarning.Visible = True And .chkRecordCountWarning.Enabled = True Then
18840           .chkRecordCountWarning.SetFocus
18850         ElseIf .chkFileSizeWarning.Visible = True And .chkFileSizeWarning.Enabled = True Then
18860           .chkFileSizeWarning.SetFocus
18870         Else
18880           .cmdCompact.SetFocus
18890         End If
18900       End With
18910     End Select
18920   End If

EXITP:
18930   KeyCode = intRetVal
18940   Exit Sub

ERRH:
18950   intRetVal = 0
18960   Select Case ERR.Number
        Case Else
18970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18980   End Select
18990   Resume EXITP

End Sub

Private Sub opgCompany_optHintonCPA_KeyDown(KeyCode As Integer, Shift As Integer)

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgCompany_optHintonCPA_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19020   intShiftDown = (Shift And acShiftMask) > 0
19030   intAltDown = (Shift And acAltMask) > 0
19040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19070       With Me
19080         intRetVal = 0
19090         .cmdBrowse.SetFocus
19100       End With
19110     End Select
19120   End If

        ' ** Shift keys.
19130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19150       With Me
19160         intRetVal = 0
19170         If .opgData.Visible = True And .opgData.Enabled = True Then
19180           .opgData.SetFocus
19190         ElseIf .cmdSetDefPath.Visible = True And .cmdSetDefPath.Enabled = True Then
19200           .cmdSetDefPath.SetFocus
19210         ElseIf .cmdUseThisPath.Visible = True And .cmdUseThisPath.Enabled = True Then
19220           .cmdUseThisPath.SetFocus
19230         ElseIf .chkRecordCountWarning.Visible = True And .chkRecordCountWarning.Enabled = True Then
19240           .chkRecordCountWarning.SetFocus
19250         ElseIf .chkFileSizeWarning.Visible = True And .chkFileSizeWarning.Enabled = True Then
19260           .chkFileSizeWarning.SetFocus
19270         Else
19280           .cmdCompact.SetFocus
19290         End If
19300       End With
19310     End Select
19320   End If

EXITP:
19330   KeyCode = intRetVal
19340   Exit Sub

ERRH:
19350   intRetVal = 0
19360   Select Case ERR.Number
        Case Else
19370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19380   End Select
19390   Resume EXITP

End Sub

Private Sub PathFile_Dta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_Dta_MouseMove"

19410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
19420   Exit Sub

ERRH:
19430   Select Case ERR.Number
        Case Else
19440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19450   End Select
19460   Resume EXITP

End Sub

Private Sub PathFile_Dta_GotFocus()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_Dta_GotFocus"

19510   With Me
19520     .PathFile_Dta.SelLength = 0
19530     .PathFile_Dta.SelStart = 99
19540   End With

EXITP:
19550   Exit Sub

ERRH:
19560   Select Case ERR.Number
        Case Else
19570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19580   End Select
19590   Resume EXITP

End Sub

Private Sub cmdDataSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDataSize_MouseDown"

19610   With Me
19620     Select Case Right(.DataSize, 2)
          Case "Kb"
19630       blnWasKb = True
19640       .DataSize = IIf(Round(((lngDataSize / 1024&) / 1024&), 2) = CLng(Round(((lngDataSize / 1024&) / 1024&), 2)), _
              Format(Round(((lngDataSize / 1024&) / 1024&), 2), "#,##0"), _
              Format(Round(((lngDataSize / 1024&) / 1024&), 2), "#,##0.0#")) & " Mb"
19650     Case "Mb"
19660       blnWasKb = False
19670       .DataSize = IIf(Round((lngDataSize / 1024&), 2) = CLng(Round((lngDataSize / 1024&), 2)), _
              Format(Round((lngDataSize / 1024&), 2), "#,##0"), _
              Format(Round((lngDataSize / 1024&), 2), "#,##0.0#")) & " Kb"
19680     End Select
19690   End With

EXITP:
19700   Exit Sub

ERRH:
19710   Select Case ERR.Number
        Case Else
19720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19730   End Select
19740   Resume EXITP

End Sub

Private Sub cmdDataSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDataSize_MouseMove"

19810   With Me
19820     .DataSize_lbl.BackStyle = acBackStyleNormal
19830     .DataSize_lbl.FontBold = True
19840   End With

EXITP:
19850   Exit Sub

ERRH:
19860   Select Case ERR.Number
        Case Else
19870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19880   End Select
19890   Resume EXITP

End Sub

Private Sub cmdDataSize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDataSize_MouseUp"

19910   With Me
19920     Select Case blnWasKb
          Case True
19930       .DataSize = IIf(Round((lngDataSize / 1024&), 2) = CLng(Round((lngDataSize / 1024&), 2)), _
              Format(Round((lngDataSize / 1024&), 2), "#,##0"), _
              Format(Round((lngDataSize / 1024&), 2), "#,##0.0#")) & " Kb"
19940     Case False
19950       .DataSize = IIf(Round(((lngDataSize / 1024&) / 1024&), 2) = CLng(Round(((lngDataSize / 1024&) / 1024&), 2)), _
              Format(Round(((lngDataSize / 1024&) / 1024&), 2), "#,##0"), _
              Format(Round(((lngDataSize / 1024&) / 1024&), 2), "#,##0.0#")) & " Mb"
19960     End Select
19970   End With

EXITP:
19980   Exit Sub

ERRH:
19990   Select Case ERR.Number
        Case Else
20000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20010   End Select
20020   Resume EXITP

End Sub

Private Sub DataSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "DataSize_MouseMove"

20110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
20120   Exit Sub

ERRH:
20130   Select Case ERR.Number
        Case Else
20140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20150   End Select
20160   Resume EXITP

End Sub

Private Sub chkUnlock_Dta_AfterUpdate()

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkUnlock_Dta_AfterUpdate"

20210   With Me
20220     Select Case .chkUnlock_Dta
          Case True
20230       .chkUnlock_Dta_lbl.FontBold = True
20240       .PathFile_Dta.Locked = False
20250       .PathFile_Dta.BackColor = CLR_WHT
20260       .PathFile_Dta.SetFocus
20270     Case False
20280       .chkUnlock_Dta_lbl.FontBold = False
20290       .PathFile_Dta.Locked = True
20300       .PathFile_Dta.BackColor = CLR_VLTTEAL
20310     End Select
20320   End With

EXITP:
20330   Exit Sub

ERRH:
20340   Select Case ERR.Number
        Case Else
20350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20360   End Select
20370   Resume EXITP

End Sub

Private Sub chkUnlock_Dta_KeyDown(KeyCode As Integer, Shift As Integer)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkUnlock_Dta_KeyDown"

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
20490         .chkUnlock_Arch.SetFocus
20500       End With
20510     End Select
20520   End If

        ' ** Shift keys.
20530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
20550       With Me
20560         intRetVal = 0
20570         .cmdBrowse.SetFocus
20580       End With
20590     End Select
20600   End If

EXITP:
20610   KeyCode = intRetVal
20620   Exit Sub

ERRH:
20630   intRetVal = 0
20640   Select Case ERR.Number
        Case Else
20650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20660   End Select
20670   Resume EXITP

End Sub

Private Sub cmdBrowse_GotFocus()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_GotFocus"

20710   With Me
20720     blnBrowse_Focus = True
20730     .cmdBrowse_raised_semifocus_dots_img.Visible = True
20740     .cmdBrowse_raised_img.Visible = False
20750     .cmdBrowse_raised_focus_img.Visible = False
20760     .cmdBrowse_raised_focus_dots_img.Visible = False
20770     .cmdBrowse_sunken_focus_dots_img.Visible = False
20780     .cmdBrowse_raised_img_dis.Visible = False
20790   End With

EXITP:
20800   Exit Sub

ERRH:
20810   Select Case ERR.Number
        Case Else
20820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20830   End Select
20840   Resume EXITP

End Sub

Private Sub cmdBrowse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_MouseDown"

20910   With Me
20920     blnBrowse_MouseDown = True
20930     .cmdBrowse_sunken_focus_dots_img.Visible = True
20940     .cmdBrowse_raised_img.Visible = False
20950     .cmdBrowse_raised_semifocus_dots_img.Visible = False
20960     .cmdBrowse_raised_focus_img.Visible = False
20970     .cmdBrowse_raised_focus_dots_img.Visible = False
20980     .cmdBrowse_raised_img_dis.Visible = False
20990   End With

EXITP:
21000   Exit Sub

ERRH:
21010   Select Case ERR.Number
        Case Else
21020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21030   End Select
21040   Resume EXITP

End Sub

Private Sub cmdBrowse_Click()
' ** Use Windows API for returning a path but no file name.

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_Click"

        Dim strTitle1 As String
        Dim strPath As String, strAuxPath As String
        Dim blnDtaFound As Boolean, blnArchFound As Boolean, blnAuxFound As Boolean
        Dim varTmp00 As Variant

21110   With Me
21120     varTmp00 = DLookup("[seclic_auxloc]", "tblSecurity_License")
21130     Select Case IsNull(varTmp00)
          Case True
21140       blnAuxLoc = False
21150     Case False
21160       Select Case varTmp00
            Case True
21170         blnAuxLoc = True
21180       Case False
21190         blnAuxLoc = False
21200       End Select
21210     End Select
21220     strTitle1 = "Select A Folder"
21230     strPath = GetFolderPathSIS(strTitle1)  ' ** Module Function: modBrowseFilesAndFolders.
21240     strAuxPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
21250     If blnAuxLoc = False Then strAuxPath = strPath
21260     If strPath <> vbNullString Then
21270       blnDtaFound = FileExists(strPath & LNK_SEP & gstrFile_DataName)  ' ** Module Function: modFileUtilities.
21280       blnArchFound = FileExists(strPath & LNK_SEP & gstrFile_ArchDataName)  ' ** Module Function: modFileUtilities.
21290       blnAuxFound = FileExists(strAuxPath & LNK_SEP & gstrFile_AuxDataName)  ' ** Module Function: modFileUtilities.
21300       If blnDtaFound = True And blnArchFound = True And blnAuxFound = True Then
21310         If ((strPathFile_Data = vbNullString) Or (Parse_Path(strPathFile_Data) <> strPath)) Or _
                  ((strPathFile_Archive = vbNullString) Or (Parse_Path(strPathFile_Archive) <> strPath)) Or _
                  ((strPathFile_Auxiliary = vbNullString) Or (Parse_Path(strPathFile_Auxiliary) <> strAuxPath)) Then  ' ** Module Function: modFileUtilities.
21320           Beep
21330           If MsgBox("Are you sure you want to choose a different location for your Trust Accountant data?", _
                    vbQuestion + vbYesNo, "Change Data Location") = vbYes Then
                  ' ** Relink to new location.
21340             .PathFile_Dta = strPath & LNK_SEP & gstrFile_DataName
21350             .PathFile_Dta.BackColor = CLR_LTRED
21360             .PathFile_Arch = strPath & LNK_SEP & gstrFile_ArchDataName
21370             .PathFile_Arch.BackColor = CLR_LTRED
21380             .PathFile_Aux = strAuxPath & LNK_SEP & gstrFile_AuxDataName
21390             .PathFile_Aux.BackColor = CLR_LTRED
21400             DoEvents
21410             intLinkMode = LNK_CHANGE
21420           Else
21430             intLinkMode = LNK_REFRESH
21440           End If
21450         Else
                ' ** Refresh existing location.
21460           intLinkMode = LNK_REFRESH
21470         End If
21480       Else
21490         If blnDtaFound = False Then
21500           MsgBox "The file " & gstrFile_DataName & " could not be found in the location chosen.", vbCritical + vbOKOnly, "File Not Found"
21510         ElseIf blnArchFound = False Then
21520           MsgBox "The file " & gstrFile_ArchDataName & " could not be found in the location chosen.", vbCritical + vbOKOnly, "File Not Found"
21530         ElseIf blnAuxFound = False Then
21540           MsgBox "The file " & gstrFile_AuxDataName & " could not be found in the location chosen.", vbCritical + vbOKOnly, "File Not Found"
21550         End If
21560       End If
21570     End If
21580     Select Case intLinkMode
          Case LNK_REFRESH
21590       .cmdRefresh.Enabled = True
21600       .cmdRelink.Enabled = False
21610     Case LNK_CHANGE
21620       .cmdRelink.Enabled = True
21630       .cmdRefresh.Enabled = False
21640     End Select
21650   End With

EXITP:
21660   Exit Sub

ERRH:
21670   Select Case ERR.Number
        Case Else
21680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21690   End Select
21700   Resume EXITP

End Sub

Private Sub cmdBrowse_KeyDown(KeyCode As Integer, Shift As Integer)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_KeyDown"

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
          Case vbKeyTab
21870       With Me
21880         intRetVal = 0
21890         If .cmdRefresh.Enabled = True Then
21900           .cmdRefresh.SetFocus
21910         ElseIf .cmdRelink.Enabled = True Then
21920           .cmdRelink.SetFocus
21930         Else
21940           .cmdClose.SetFocus
21950         End If
21960       End With
21970     End Select
21980   End If

        ' ** Shift keys.
21990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22000     Select Case intRetVal
          Case vbKeyTab
22010       With Me
22020         intRetVal = 0
22030         If .opgCompany.Visible = True And .opgCompany.Enabled = True Then
22040           .opgCompany.SetFocus
22050         ElseIf .opgData.Visible = True And .opgData.Enabled = True Then
22060           .opgData.SetFocus
22070         ElseIf .cmdSetDefPath.Visible = True And .cmdSetDefPath.Enabled = True Then
22080           .cmdSetDefPath.SetFocus
22090         ElseIf .cmdUseThisPath.Visible = True And .cmdUseThisPath.Enabled = True Then
22100           .cmdUseThisPath.SetFocus
22110         ElseIf .chkRecordCountWarning.Visible = True And .chkRecordCountWarning.Enabled = True Then
22120           .chkRecordCountWarning.SetFocus
22130         ElseIf .chkFileSizeWarning.Visible = True And .chkFileSizeWarning.Enabled = True Then
22140           .chkFileSizeWarning.SetFocus
22150         Else
22160           .cmdCompact.SetFocus
22170         End If
22180       End With
22190     End Select
22200   End If

        ' ** Ctrl keys.
22210   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
22220     Select Case intRetVal
          Case vbKeyTab
22230       With Me
22240         intRetVal = 0
22250         .chkUnlock_Dta.SetFocus
22260       End With
22270     End Select
22280   End If

EXITP:
22290   KeyCode = intRetVal
22300   Exit Sub

ERRH:
22310   intRetVal = 0
22320   Select Case ERR.Number
        Case Else
22330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22340   End Select
22350   Resume EXITP

End Sub

Private Sub cmdBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_MouseMove"

22410   With Me
22420     If blnBrowse_MouseDown = False Then
22430       Select Case blnBrowse_Focus
            Case True
22440         .cmdBrowse_raised_focus_dots_img.Visible = True
22450         .cmdBrowse_raised_focus_img.Visible = False
22460       Case False
22470         .cmdBrowse_raised_focus_img.Visible = True
22480         .cmdBrowse_raised_focus_dots_img.Visible = False
22490       End Select
22500       .cmdBrowse_raised_img.Visible = False
22510       .cmdBrowse_raised_semifocus_dots_img.Visible = False
22520       .cmdBrowse_sunken_focus_dots_img.Visible = False
22530       .cmdBrowse_raised_img_dis.Visible = False
22540     End If
22550   End With

EXITP:
22560   Exit Sub

ERRH:
22570   Select Case ERR.Number
        Case Else
22580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22590   End Select
22600   Resume EXITP

End Sub

Private Sub cmdBrowse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_MouseUp"

22710   With Me
22720     .cmdBrowse_raised_focus_dots_img.Visible = True
22730     .cmdBrowse_raised_img.Visible = False
22740     .cmdBrowse_raised_semifocus_dots_img.Visible = False
22750     .cmdBrowse_raised_focus_img.Visible = False
22760     .cmdBrowse_sunken_focus_dots_img.Visible = False
22770     .cmdBrowse_raised_img_dis.Visible = False
22780     blnBrowse_MouseDown = False
22790   End With

EXITP:
22800   Exit Sub

ERRH:
22810   Select Case ERR.Number
        Case Else
22820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22830   End Select
22840   Resume EXITP

End Sub

Private Sub cmdBrowse_LostFocus()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_LostFocus"

22910   With Me
22920     .cmdBrowse_raised_img.Visible = True
22930     .cmdBrowse_raised_semifocus_dots_img.Visible = False
22940     .cmdBrowse_raised_focus_img.Visible = False
22950     .cmdBrowse_raised_focus_dots_img.Visible = False
22960     .cmdBrowse_sunken_focus_dots_img.Visible = False
22970     .cmdBrowse_raised_img_dis.Visible = False
22980     blnBrowse_Focus = False
22990   End With

EXITP:
23000   Exit Sub

ERRH:
23010   Select Case ERR.Number
        Case Else
23020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23030   End Select
23040   Resume EXITP

End Sub

Private Sub PathFile_Arch_GotFocus()

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_Arch_GotFocus"

23110   With Me
23120     .PathFile_Arch.SelLength = 0
23130     .PathFile_Arch.SelStart = 99
23140   End With

EXITP:
23150   Exit Sub

ERRH:
23160   Select Case ERR.Number
        Case Else
23170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23180   End Select
23190   Resume EXITP

End Sub

Private Sub cmdArchSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchSize_MouseDown"

23210   With Me
23220     Select Case Right(.ArchSize, 2)
          Case "Kb"
23230       blnWasKb = True
23240       .ArchSize = IIf(Round(((lngArchSize / 1024&) / 1024&), 2) = CLng(Round(((lngArchSize / 1024&) / 1024&), 2)), _
              Format(Round(((lngArchSize / 1024&) / 1024&), 2), "#,##0"), _
              Format(Round(((lngArchSize / 1024&) / 1024&), 2), "#,##0.0#")) & " Mb"
23250     Case "Mb"
23260       blnWasKb = False
23270       .ArchSize = IIf(Round((lngArchSize / 1024&), 2) = CLng(Round((lngArchSize / 1024&), 2)), _
              Format(Round((lngArchSize / 1024&), 2), "#,##0"), _
              Format(Round((lngArchSize / 1024&), 2), "#,##0.0#")) & " Kb"
23280     End Select
23290   End With

EXITP:
23300   Exit Sub

ERRH:
23310   Select Case ERR.Number
        Case Else
23320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23330   End Select
23340   Resume EXITP

End Sub

Private Sub cmdArchSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchSize_MouseMove"

23410   With Me
23420     .ArchSize_lbl.BackStyle = acBackStyleNormal
23430     .ArchSize_lbl.FontBold = True
23440   End With

EXITP:
23450   Exit Sub

ERRH:
23460   Select Case ERR.Number
        Case Else
23470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23480   End Select
23490   Resume EXITP

End Sub

Private Sub cmdArchSize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdArchSize_MouseUp"

23510   With Me
23520     Select Case blnWasKb
          Case True
23530       .ArchSize = IIf(Round((lngArchSize / 1024&), 2) = CLng(Round((lngArchSize / 1024&), 2)), _
              Format(Round((lngArchSize / 1024&), 2), "#,##0"), _
              Format(Round((lngArchSize / 1024&), 2), "#,##0.0#")) & " Kb"
23540     Case False
23550       .ArchSize = IIf(Round(((lngArchSize / 1024&) / 1024&), 2) = CLng(Round(((lngArchSize / 1024&) / 1024&), 2)), _
              Format(Round(((lngArchSize / 1024&) / 1024&), 2), "#,##0"), _
              Format(Round(((lngArchSize / 1024&) / 1024&), 2), "#,##0.0#")) & " Mb"
23560     End Select
23570   End With

EXITP:
23580   Exit Sub

ERRH:
23590   Select Case ERR.Number
        Case Else
23600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23610   End Select
23620   Resume EXITP

End Sub

Private Sub ArchSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "ArchSize_MouseMove"

23710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23720   Exit Sub

ERRH:
23730   Select Case ERR.Number
        Case Else
23740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23750   End Select
23760   Resume EXITP

End Sub

Private Sub chkUnlock_Arch_AfterUpdate()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkUnlock_Arch_AfterUpdate"

23810   With Me
23820     Select Case .chkUnlock_Arch
          Case True
23830       .chkUnlock_Arch_lbl.FontBold = True
23840       .PathFile_Arch.Locked = False
23850       .PathFile_Arch.BackColor = CLR_WHT
23860       .PathFile_Arch.SetFocus
23870     Case False
23880       .chkUnlock_Arch_lbl.FontBold = False
23890       .PathFile_Arch.Locked = True
23900       .PathFile_Arch.BackColor = CLR_VLTTEAL
23910     End Select
23920   End With

EXITP:
23930   Exit Sub

ERRH:
23940   Select Case ERR.Number
        Case Else
23950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23960   End Select
23970   Resume EXITP

End Sub

Private Sub chkUnlock_Arch_KeyDown(KeyCode As Integer, Shift As Integer)

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkUnlock_Arch_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24020   intShiftDown = (Shift And acShiftMask) > 0
24030   intAltDown = (Shift And acAltMask) > 0
24040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24070       With Me
24080         intRetVal = 0
24090         .chkUnlock_Aux.SetFocus
24100       End With
24110     End Select
24120   End If

        ' ** Shift keys.
24130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
24150       With Me
24160         intRetVal = 0
24170         .chkUnlock_Dta.SetFocus
24180       End With
24190     End Select
24200   End If

EXITP:
24210   KeyCode = intRetVal
24220   Exit Sub

ERRH:
24230   intRetVal = 0
24240   Select Case ERR.Number
        Case Else
24250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24260   End Select
24270   Resume EXITP

End Sub

Private Sub PathFile_Aux_GotFocus()

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "PathFile_Aux_GotFocus"

24310   With Me
24320     .PathFile_Aux.SelLength = 0
24330     .PathFile_Aux.SelStart = 99
24340   End With

EXITP:
24350   Exit Sub

ERRH:
24360   Select Case ERR.Number
        Case Else
24370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24380   End Select
24390   Resume EXITP

End Sub

Private Sub cmdAuxSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAuxSize_MouseDown"

24410   With Me
24420     Select Case Right(.AuxSize, 2)
          Case "Kb"
24430       blnWasKb = True
24440       .AuxSize = IIf(Round(((lngAuxSize / 1024&) / 1024&), 2) = CLng(Round(((lngAuxSize / 1024&) / 1024&), 2)), _
              Format(Round(((lngAuxSize / 1024&) / 1024&), 2), "#,##0"), _
              Format(Round(((lngAuxSize / 1024&) / 1024&), 2), "#,##0.0#")) & " Mb"
24450     Case "Mb"
24460       blnWasKb = False
24470       .AuxSize = IIf(Round((lngAuxSize / 1024&), 2) = CLng(Round((lngAuxSize / 1024&), 2)), _
              Format(Round((lngAuxSize / 1024&), 2), "#,##0"), _
              Format(Round((lngAuxSize / 1024&), 2), "#,##0.0#")) & " Kb"
24480     End Select
24490   End With

EXITP:
24500   Exit Sub

ERRH:
24510   Select Case ERR.Number
        Case Else
24520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24530   End Select
24540   Resume EXITP

End Sub

Private Sub cmdAuxSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAuxSize_MouseMove"

24610   With Me
24620     .AuxSize_lbl.BackStyle = acBackStyleNormal
24630     .AuxSize_lbl.FontBold = True
24640   End With

EXITP:
24650   Exit Sub

ERRH:
24660   Select Case ERR.Number
        Case Else
24670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24680   End Select
24690   Resume EXITP

End Sub

Private Sub cmdAuxSize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAuxSize_MouseUp"

24710   With Me
24720     Select Case blnWasKb
          Case True
24730       .AuxSize = IIf(Round((lngAuxSize / 1024&), 2) = CLng(Round((lngAuxSize / 1024&), 2)), _
              Format(Round((lngAuxSize / 1024&), 2), "#,##0"), _
              Format(Round((lngAuxSize / 1024&), 2), "#,##0.0#")) & " Kb"
24740     Case False
24750       .AuxSize = IIf(Round(((lngAuxSize / 1024&) / 1024&), 2) = CLng(Round(((lngAuxSize / 1024&) / 1024&), 2)), _
              Format(Round(((lngAuxSize / 1024&) / 1024&), 2), "#,##0"), _
              Format(Round(((lngAuxSize / 1024&) / 1024&), 2), "#,##0.0#")) & " Mb"
24760     End Select
24770   End With

EXITP:
24780   Exit Sub

ERRH:
24790   Select Case ERR.Number
        Case Else
24800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24810   End Select
24820   Resume EXITP

End Sub

Private Sub AuxSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "AuxSize_MouseMove"

24910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
24920   Exit Sub

ERRH:
24930   Select Case ERR.Number
        Case Else
24940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24950   End Select
24960   Resume EXITP

End Sub

Private Sub chkUnlock_Aux_AfterUpdate()

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkUnlock_Aux_AfterUpdate"

25010   With Me
25020     Select Case .chkUnlock_Aux
          Case True
25030       .chkUnlock_Aux_lbl.FontBold = True
25040       .PathFile_Aux.Locked = False
25050       .PathFile_Aux.BackColor = CLR_WHT
25060       .PathFile_Aux.SetFocus
25070     Case False
25080       .chkUnlock_Aux_lbl.FontBold = False
25090       .PathFile_Aux.Locked = True
25100       .PathFile_Aux.BackColor = CLR_VLTTEAL
25110     End Select
25120   End With

EXITP:
25130   Exit Sub

ERRH:
25140   Select Case ERR.Number
        Case Else
25150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25160   End Select
25170   Resume EXITP

End Sub

Private Sub chkUnlock_Aux_KeyDown(KeyCode As Integer, Shift As Integer)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkUnlock_Aux_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25220   intShiftDown = (Shift And acShiftMask) > 0
25230   intAltDown = (Shift And acAltMask) > 0
25240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25270       With Me
25280         intRetVal = 0
25290         If .cmdRefresh.Enabled = True Then
25300           .cmdRefresh.SetFocus
25310         ElseIf .cmdRelink.Enabled = True Then
25320           .cmdRelink.SetFocus
25330         End If
25340       End With
25350     End Select
25360   End If

        ' ** Shift keys.
25370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
25390       With Me
25400         intRetVal = 0
25410         .chkUnlock_Arch.SetFocus
25420       End With
25430     End Select
25440   End If

EXITP:
25450   KeyCode = intRetVal
25460   Exit Sub

ERRH:
25470   intRetVal = 0
25480   Select Case ERR.Number
        Case Else
25490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25500   End Select
25510   Resume EXITP

End Sub

Public Sub ProgBar_Width_Link(dblWidth As Double, intMode As Integer)

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "ProgBar_Width_Link"

        Dim strCtlName As String, blnVis As Boolean
        Dim lngX As Long

25610   With Me
25620     Select Case intMode
          Case 1
25630       blnVis = CBool(dblWidth)
25640       For lngX = 1& To 6&
25650         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
25660         .Controls(strCtlName).Visible = blnVis
25670       Next
25680       If blnProgBarOff = True Then
25690         .ProgBar_box.Visible = False
25700         .ProgBar_box2.Visible = False
25710       End If
25720     Case 2
25730       For lngX = 1& To 6&
25740         strCtlName = "ProgBar_bar" & Right("00" & CStr(lngX), 2)
25750         .Controls(strCtlName).Width = dblWidth
25760       Next
25770     End Select
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
