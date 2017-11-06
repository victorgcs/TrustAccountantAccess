VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBackupRestore_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmBackupRestore_File"

'VGC 07/14/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Browse:           B {cmdBrowse}
' **   Cancel:           C {cmdCancel}
' **   File:             F {TargetFile}
' **   Backup Now:       N {cmdBackup}
' **   Path:             P {TargetPath}

' ** Called by:
' **   frmBackupOptions.cmdBackup_Click()

Private strTargetFile As String
Private strModeCmd As String, lngTpp As Long, lngTargetPath_Height As Long
Private blnIsQuit As Boolean, blnDone As Boolean
Private blnBrowse_Focus As Boolean, blnBrowse_MouseDown As Boolean
Private strCallingForm As String, strCallersCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String

110     blnDone = False

120     With Me
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then
150           strTmp01 = Trim(.OpenArgs)
160           intPos01 = InStr(strTmp01, "~")
170           If intPos01 > 0 Then

180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             strModeCmd = Left(strTmp01, (intPos01 - 1))
220             strCallersCallingForm = Mid(strTmp01, (intPos01 + 1))  ' ** Calling form's calling form.

230             blnIsQuit = False
240             blnBrowse_Focus = False: blnBrowse_MouseDown = False

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
250             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
260             lngTargetPath_Height = .TargetPath.Height

270             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
280             DoEvents

290             TargetPath_AfterUpdate  ' ** Procedure: Below.

300             strTargetFile = "TrstDta_" & year(Date) & Format(month(Date), "00") & Format(day(Date), "00") & ".md_"
310             strTmp01 = strTargetFile
320             If Right(strTmp01, 4) = ".md_" Then
330               strTmp01 = Left(strTmp01, (Len(strTmp01) - 4))
340             End If
350             .TargetFile = strTmp01

360           Else
370             Cancel = -1
380           End If
390         Else
400           Cancel = -1
410         End If
420       Else
430         Cancel = -1
440       End If
450     End With

460     If Cancel = -1 Then
470       Beep
480       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
490     End If

EXITP:
500     Exit Sub

ERRH:
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyEscape
670         intRetVal = 0
680         cmdCancel_Click  ' ** Procedure: Below.
690       End Select
700     End If

        ' ** Ctrl-Shift keys.
710     If intCtrlDown And (Not intAltDown) And intShiftDown Then
720       Select Case intRetVal
          Case vbKeyF
730         With Me
740           intRetVal = 0
750           .FocusHolder.SetFocus
760         End With
770       End Select
780     End If

EXITP:
790     KeyCode = intRetVal
800     Exit Sub

ERRH:
810     intRetVal = 0
820     Select Case ERR.Number
        Case Else
830       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
840     End Select
850     Resume EXITP

End Sub

Private Sub Form_Timer()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

910     With Me
920       .TimerInterval = 0&
930       If gblnSetFocus = True Then
940         gblnSetFocus = False
950         strFormName = vbNullString
960   On Error Resume Next
970         Do Until strFormName = THIS_NAME
980           DoCmd.SelectObject acForm, THIS_NAME, False
990           strFormName = Screen.ActiveForm.Name
1000        Loop
1010  On Error GoTo ERRH
1020      End If
1030    End With

EXITP:
1040    Exit Sub

ERRH:
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1110    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1120    Select Case blnIsQuit
        Case True
          ' ** Do nothing else, just let it close.
1130    Case False
1140      gstrReturningForm = THIS_NAME
1150      gblnSetFocus = True
1160      Select Case strCallingForm
          Case "frmBackupRestore"
1170        Select Case strCallersCallingForm
            Case "frmMenu_Utility"
1180          DoCmd.OpenForm strCallersCallingForm, , , , , , "frmMenu_Main"
1190        Case "frmMenu_Title"
1200          Select Case blnDone
              Case True
1210            QuitNow  ' ** Module Procedure: modStartupFuncs.
1220          Case False
                ' ** I don't think it'll get here.
1230          End Select
1240        Case Else
              ' ** Nothing at the moment.
1250        End Select
1260      Case Else
            ' ** Nothing at the moment.
1270      End Select
1280    End Select

EXITP:
1290    Exit Sub

ERRH:
1300    Select Case ERR.Number
        Case Else
1310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1320    End Select
1330    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1410    With Me
1420      If .cmdBrowse_raised_focus_dots_img.Visible = True Or .cmdBrowse_raised_focus_img.Visible = True Then
1430        Select Case blnBrowse_Focus
            Case True
1440          .cmdBrowse_raised_semifocus_dots_img.Visible = True
1450          .cmdBrowse_raised_img.Visible = False
1460        Case False
1470          .cmdBrowse_raised_img.Visible = True
1480          .cmdBrowse_raised_semifocus_dots_img.Visible = False
1490        End Select
1500        .cmdBrowse_raised_focus_img.Visible = False
1510        .cmdBrowse_raised_focus_dots_img.Visible = False
1520        .cmdBrowse_sunken_focus_dots_img.Visible = False
1530        .cmdBrowse_raised_img_dis.Visible = False
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

Private Sub cmdCancel_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

1710    gblnBeenToBackup = False  ' ** Reset marker if you close out of this form.
1720    blnDone = True

1730    DoCmd.Close acForm, THIS_NAME

EXITP:
1740    Exit Sub

ERRH:
1750    Select Case ERR.Number
        Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1820    intShiftDown = (Shift And acShiftMask) > 0
1830    intAltDown = (Shift And acAltMask) > 0
1840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1860      Select Case intRetVal
          Case vbKeyTab
1870        With Me
1880          intRetVal = 0
1890          .TargetPath.SetFocus
1900        End With
1910      End Select
1920    End If

        ' ** Shift keys.
1930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1940      Select Case intRetVal
          Case vbKeyTab
1950        With Me
1960          intRetVal = 0
1970          If .cmdBackup.Enabled = True Then
1980            .cmdBackup.SetFocus
1990          Else
2000            .TargetFile.SetFocus
2010          End If
2020        End With
2030      End Select
2040    End If

EXITP:
2050    KeyCode = intRetVal
2060    Exit Sub

ERRH:
2070    intRetVal = 0
2080    Select Case ERR.Number
        Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Sub

Private Sub Target_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Target_box_MouseMove"

2210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub TargetPath_AfterUpdate()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "TargetPath_AfterUpdate"

        Dim strPath As String, lngDiff As Long

2310    With Me
2320      If IsNull(.TargetPath) = False Then
2330        strPath = Trim(.TargetPath)
2340        If strPath <> vbNullString Then
2350          If Len(strPath) > 50 And .TargetPath.Height = lngTargetPath_Height Then
2360            lngDiff = ((34& * lngTpp) - lngTargetPath_Height)  ' ** 1-Line: 285 Twips; 2-Line: 510 Twips.
2370            .TargetPath.Height = (.TargetPath.Height + lngDiff)
2380            .cmdBrowse.Top = (.cmdBrowse.Top + lngDiff)
2390            .cmdBrowse_raised_focus_dots_img.Top = .cmdBrowse.Top
2400            .cmdBrowse_raised_focus_img.Top = .cmdBrowse.Top
2410            .cmdBrowse_raised_img.Top = .cmdBrowse.Top
2420            .cmdBrowse_raised_img_dis.Top = .cmdBrowse.Top
2430            .cmdBrowse_raised_semifocus_dots_img.Top = .cmdBrowse.Top
2440            .cmdBrowse_sunken_focus_dots_img.Top = .cmdBrowse.Top
2450          ElseIf Len(strPath) <= 50 And .TargetPath.Height > lngTargetPath_Height Then
2460            lngDiff = .TargetPath.Height - lngTargetPath_Height
2470            .TargetPath.Height = lngTargetPath_Height
2480            .cmdBrowse.Top = (.cmdBrowse.Top - lngDiff)
2490            .cmdBrowse_raised_focus_dots_img.Top = .cmdBrowse.Top
2500            .cmdBrowse_raised_focus_img.Top = .cmdBrowse.Top
2510            .cmdBrowse_raised_img.Top = .cmdBrowse.Top
2520            .cmdBrowse_raised_img_dis.Top = .cmdBrowse.Top
2530            .cmdBrowse_raised_semifocus_dots_img.Top = .cmdBrowse.Top
2540            .cmdBrowse_sunken_focus_dots_img.Top = .cmdBrowse.Top
2550          End If
2560          .TargetPath.SetFocus
2570          .TargetPath.SelLength = 0
2580          .TargetPath.SelStart = 99
2590        End If
2600      End If
2610    End With

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub TargetPath_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "TargetPath_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2770        With Me
2780          intRetVal = 0
2790          .TargetFile.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2850        With Me
2860          intRetVal = 0
2870          .cmdCancel.SetFocus
2880        End With
2890      End Select
2900    End If

EXITP:
2910    KeyCode = intRetVal
2920    Exit Sub

ERRH:
2930    intRetVal = 0
2940    Select Case ERR.Number
        Case Else
2950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2960    End Select
2970    Resume EXITP

End Sub

Private Sub TargetPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "TargetPath_MouseMove"

3010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3020    Exit Sub

ERRH:
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub cmdBrowse_GotFocus()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_GotFocus"

3110    With Me
3120      blnBrowse_Focus = True
3130      .cmdBrowse_raised_semifocus_dots_img.Visible = True
3140      .cmdBrowse_raised_img.Visible = False
3150      .cmdBrowse_raised_focus_img.Visible = False
3160      .cmdBrowse_raised_focus_dots_img.Visible = False
3170      .cmdBrowse_sunken_focus_dots_img.Visible = False
3180      .cmdBrowse_raised_img_dis.Visible = False
3190    End With

EXITP:
3200    Exit Sub

ERRH:
3210    Select Case ERR.Number
        Case Else
3220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3230    End Select
3240    Resume EXITP

End Sub

Private Sub cmdBrowse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_MouseDown"

3310    With Me
3320      blnBrowse_MouseDown = True
3330      .cmdBrowse_sunken_focus_dots_img.Visible = True
3340      .cmdBrowse_raised_img.Visible = False
3350      .cmdBrowse_raised_semifocus_dots_img.Visible = False
3360      .cmdBrowse_raised_focus_img.Visible = False
3370      .cmdBrowse_raised_focus_dots_img.Visible = False
3380      .cmdBrowse_raised_img_dis.Visible = False
3390    End With

EXITP:
3400    Exit Sub

ERRH:
3410    Select Case ERR.Number
        Case Else
3420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3430    End Select
3440    Resume EXITP

End Sub

Private Sub cmdBrowse_Click()
' ** Use Windows API for returning a path but no file name.

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_Click"

        Dim strTitle1 As String
        Dim strPath As String

3510    With Me
3520      strTitle1 = "Select a folder"
3530      strPath = GetFolderPathSIS(strTitle1)  ' ** Module Function: modBrowseFilesAndFolders.
3540      If strPath <> vbNullString Then
3550        .TargetPath = strPath
3560        DoEvents
3570        TargetPath_AfterUpdate  ' ** Procedure: Below.
3580      End If
3590      gblnSetFocus = True
3600      .TimerInterval = 100&
3610    End With

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub cmdBrowse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_MouseMove"

3710    With Me
3720      If blnBrowse_MouseDown = False Then
3730        Select Case blnBrowse_Focus
            Case True
3740          .cmdBrowse_raised_focus_dots_img.Visible = True
3750          .cmdBrowse_raised_focus_img.Visible = False
3760        Case False
3770          .cmdBrowse_raised_focus_img.Visible = True
3780          .cmdBrowse_raised_focus_dots_img.Visible = False
3790        End Select
3800        .cmdBrowse_raised_img.Visible = False
3810        .cmdBrowse_raised_semifocus_dots_img.Visible = False
3820        .cmdBrowse_sunken_focus_dots_img.Visible = False
3830        .cmdBrowse_raised_img_dis.Visible = False
3840      End If
3850    End With

EXITP:
3860    Exit Sub

ERRH:
3870    Select Case ERR.Number
        Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Sub

Private Sub cmdBrowse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_MouseUp"

4010    With Me
4020      .cmdBrowse_raised_focus_dots_img.Visible = True
4030      .cmdBrowse_raised_img.Visible = False
4040      .cmdBrowse_raised_semifocus_dots_img.Visible = False
4050      .cmdBrowse_raised_focus_img.Visible = False
4060      .cmdBrowse_sunken_focus_dots_img.Visible = False
4070      .cmdBrowse_raised_img_dis.Visible = False
4080      blnBrowse_MouseDown = False
4090    End With

EXITP:
4100    Exit Sub

ERRH:
4110    Select Case ERR.Number
        Case Else
4120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4130    End Select
4140    Resume EXITP

End Sub

Private Sub cmdBrowse_LostFocus()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBrowse_LostFocus"

4210    With Me
4220      .cmdBrowse_raised_img.Visible = True
4230      .cmdBrowse_raised_semifocus_dots_img.Visible = False
4240      .cmdBrowse_raised_focus_img.Visible = False
4250      .cmdBrowse_raised_focus_dots_img.Visible = False
4260      .cmdBrowse_sunken_focus_dots_img.Visible = False
4270      .cmdBrowse_raised_img_dis.Visible = False
4280      blnBrowse_Focus = False
4290    End With

EXITP:
4300    Exit Sub

ERRH:
4310    Select Case ERR.Number
        Case Else
4320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4330    End Select
4340    Resume EXITP

End Sub

Private Sub TargetFile_AfterUpdate()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "TargetFile_AfterUpdate"

        Dim intPos01 As Integer

4410    With Me
4420      If IsNull(.TargetFile) = False Then
4430        strTargetFile = Trim(.TargetFile)
4440        If strTargetFile <> vbNullString Then
4450          intPos01 = InStr(strTargetFile, ".")
4460          If intPos01 > 0 Then
4470            strTargetFile = Left(strTargetFile, (intPos01 - 1))
4480            .TargetFile = strTargetFile  ' ** No extension in field.
4490            strTargetFile = strTargetFile & ".md_"
4500          Else
4510            strTargetFile = strTargetFile & ".md_"
4520          End If
4530        End If
4540      Else
4550        strTargetFile = vbNullString
4560      End If
4570    End With

EXITP:
4580    Exit Sub

ERRH:
4590    Select Case ERR.Number
        Case Else
4600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4610    End Select
4620    Resume EXITP

End Sub

Private Sub TargetFile_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "TargetFile_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4720    intShiftDown = (Shift And acShiftMask) > 0
4730    intAltDown = (Shift And acAltMask) > 0
4740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4770        With Me
4780          intRetVal = 0
4790          If .cmdBackup.Enabled = True Then
4800            .cmdBackup.SetFocus
4810          Else
4820            .cmdCancel.SetFocus
4830          End If
4840        End With
4850      End Select
4860    End If

        ' ** Shift keys.
4870    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4880      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4890        With Me
4900          intRetVal = 0
4910          .TargetPath.SetFocus
4920        End With
4930      End Select
4940    End If

EXITP:
4950    KeyCode = intRetVal
4960    Exit Sub

ERRH:
4970    intRetVal = 0
4980    Select Case ERR.Number
        Case Else
4990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5000    End Select
5010    Resume EXITP

End Sub

Private Sub cmdBackup_Click()

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBackup_Click"

        Dim strTargetPath As String, strTargetPathFile As String
        Dim intRetVal_Backup As Integer
        Dim strTmp01 As String
        Dim blnContinue As Boolean, blnIsDir As Boolean, blnIsVol As Boolean

5110    blnContinue = True
5120    intRetVal_Backup = 0  ' ** All's well.

5130    With Me

          ' ** Target path must be filled.
5140      If IsNull(.TargetPath) Then
5150        MsgBox "Please specify a target path.", vbInformation + vbOKOnly, "Entry Required"
5160        .TargetPath.SetFocus
5170      Else
            ' ** Target file must be filled.
5180        If IsNull(.TargetFile) Then
5190          MsgBox "Please specify a target file name.", vbInformation + vbOKOnly, "Entry Required"
5200          .TargetFile.SetFocus
5210        Else
              ' ** strTargetFile should already be filled.

5220          DoCmd.Hourglass True
5230          DoEvents

              ' ** Remove backslashes at end of folder paths.
5240          If Right(.TargetPath, 1) = LNK_SEP Then
5250            .TargetPath = Left(.TargetPath, Len(.TargetPath) - 1)
5260          End If

5270          strTargetPath = .TargetPath

5280          strTargetPathFile = strTargetPath & LNK_SEP & strTargetFile

              ' ** A COMPLETELY EMPTY FLASH DRIVE IS NOT A DIRECTORY!
              ' ** Therefore, the Dir(strTargetPath, vbDirectory) will come up empty, and it'll
              ' ** try to make that directory, when in fact it's a volume, not a directory.
              ' ** Tri-County Bank errored with 75, 'Path/File access error'.
              ' ** When we created a directory on the thumb drive, e.g., TrustBackup, it had
              ' ** no problem putting it there, and did not then try to create a new directory.

5290          blnIsDir = False: blnIsVol = False

5300  On Error Resume Next
5310          strTmp01 = Dir(strTargetPath, vbDirectory)
5320          If ERR.Number <> 0 Then
5330            Select Case ERR.Number
                Case 75  ' ** Path/File access error.
5340  On Error GoTo ERRH
5350  On Error Resume Next
5360              strTmp01 = Dir(strTargetPath, vbVolume)
5370              If ERR.Number <> 0 Then
5380                DoCmd.Hourglass False
5390                MsgBox "There is a problem with the chosen path." & vbCrLf & vbCrLf & _
                      "Error: " & CStr(ERR.Number) & vbCrLf & _
                      "Description: " & ERR.description, vbInformation + vbOKOnly, "Path Not Found"
5400  On Error GoTo ERRH
5410                blnContinue = False
5420              Else
5430  On Error GoTo ERRH
                    ' ** No error as volume.
5440                blnIsVol = True
5450                strTmp01 = strTargetPath
5460              End If
5470            Case 52  ' ** Bad file name or number.
                  ' ** If it's a CD, this is the error I get, so try vbVolume.
5480  On Error GoTo ERRH
5490  On Error Resume Next
5500              strTmp01 = Dir(strTargetPath, vbVolume)
5510              If ERR.Number <> 0 Then
5520                DoCmd.Hourglass False
5530                MsgBox "There is a problem with the chosen path." & vbCrLf & vbCrLf & _
                      "Error: " & CStr(ERR.Number) & vbCrLf & _
                      "Description: " & ERR.description, vbInformation + vbOKOnly, "Path Not Found"
5540  On Error GoTo ERRH
5550                blnContinue = False
5560              Else
5570  On Error GoTo ERRH
                    ' ** No error as volume.
5580                blnIsVol = True
5590                strTmp01 = strTargetPath
5600              End If
5610            End Select
5620          Else
5630  On Error GoTo ERRH
                ' ** No error, but may be empty.
5640            blnIsDir = True
5650            strTmp01 = strTargetPath
5660          End If

5670          If blnContinue = True Then

                ' ** Target path must exist; make it if necessary.
5680            If blnIsDir = True Then
5690              If Dir(strTmp01, vbDirectory) = vbNullString Then
5700                MkDir strTmp01
5710                strTargetPath = strTmp01  ' ** I know this just flips back-and-forth!
5720              End If
5730            ElseIf blnIsVol = True Then
5740              If Dir(strTmp01, vbVolume) = vbNullString Then
5750                If Len(strTmp01) = 2 Then
                      ' ** If it's a single letter, followed by a colon (and backslash), don't try to create it!
5760                  If Mid(strTmp01, 2, 1) = ":" And _
                          Asc(Left(strTmp01, 1)) >= 65 And Asc(Left(strTmp01, 1)) <= 90 Then
                        ' ** Let it go through.
5770                  Else
                        ' ** This is a bit iffy!
5780                    MkDir strTmp01
5790                    strTargetPath = strTmp01
5800                  End If
5810                Else
                      ' ** This is a bit iffy!
5820                  MkDir strTmp01
5830                  strTargetPath = strTmp01
5840                End If
5850              End If
5860            End If

5870            DoCmd.Hourglass True  ' ** Just to be sure.
5880            DoEvents

                ' ** Confirm and cleanup.
5890            intRetVal_Backup = BackupToDrive(strTargetPathFile)  ' ** Module Function: modBackup.
5900            DoCmd.Hourglass False
5910            Select Case intRetVal_Backup
                Case 0  ' ** All's well.
5920              MsgBox "Data was successfully backed up to:" & vbCrLf & vbCrLf & "  " & strTargetPathFile, vbInformation + vbOKOnly, ("Finished" & Space(40))
5930              DoCmd.Close acForm, THIS_NAME
5940              Select Case strCallersCallingForm
                  Case "frmMenu_Title"
5950                blnIsQuit = True
5960                QuitNow  ' ** Module Procedure: modStartupFuncs.
5970              Case Else
                    ' ** Don't return to frmBackupOptions if it's not a Cancel.
5980                blnDone = True
5990                DoCmd.Close acForm, THIS_NAME
6000              End Select
6010            Case 1  ' ** All's well CD.
                  ' ** This invoked the built-in CD Writing Wizard, which must be manually prompted to finish the copy process.
6020              MsgBox "Continue with the CD Writing Wizard to complete the backup process to:" & vbCrLf & vbCrLf & _
                    "  " & strTargetPathFile, vbExclamation + vbOKOnly, "Manually Complete Backup"
6030              DoCmd.Close acForm, THIS_NAME
6040              Select Case strCallersCallingForm
                  Case "frmMenu_Title"
6050                blnIsQuit = True
6060                QuitNow  ' ** Module Procedure: modStartupFuncs.
6070              Case Else
                    ' ** Don't return to frmBackupOptions if it's not a Cancel.
6080                blnDone = True
6090                DoCmd.Close acForm, THIS_NAME
6100              End Select
6110            Case -9  ' ** Error.
6120              Select Case strCallersCallingForm
                  Case "frmMenu_Title"
6130                blnIsQuit = True
6140                QuitNow  ' ** Module Procedure: modStartupFuncs.
6150              Case Else
6160                blnDone = True
6170                DoCmd.Close acForm, THIS_NAME
6180              End Select
6190            End Select

6200          End If  ' ** blnContinue.

6210        End If
6220      End If

6230    End With

6240    DoCmd.Hourglass False

EXITP:
6250    Exit Sub

ERRH:
6260    DoCmd.Hourglass False
6270    Select Case ERR.Number
        Case Else
6280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6290    End Select
6300    Resume EXITP

End Sub

Private Sub cmdBackup_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBackup_KeyDown"

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
6490          .cmdCancel.SetFocus
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab
6550        With Me
6560          intRetVal = 0
6570          .TargetFile.SetFocus
6580        End With
6590      End Select
6600    End If

EXITP:
6610    KeyCode = intRetVal
6620    Exit Sub

ERRH:
6630    intRetVal = 0
6640    Select Case ERR.Number
        Case Else
6650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6660    End Select
6670    Resume EXITP

End Sub
