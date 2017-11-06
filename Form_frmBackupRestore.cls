VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBackupRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmBackupRestore"

'VGC 04/16/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Drive:            D {opgLocBackup_optDrive}
' **   Drive:            D {opgLocRestore_optDrive}
' **   Backup Now:       N {cmdBackup}
' **   Restore Now:      N {cmdRestore}
' **   CD/DVD:           V {opgLocBackup_optDVD}
' **   CD/DVD:           V {opgLocRestore_optDVD}
' **   Exit:             X {cmdClose}

' ** Called by:
' **   frmMenu_Title.cmdQuit_Click()
' **   frmMenu_Utility.cmdBackupData_Click()
' **   frmMenu_Utility.cmdRestoreData_Click()

' ** Network, Hard Drive, External Drive, Flash Drive, DVD, CD

' ** A COMPLETELY EMPTY FLASH/THUMB/USB DRIVE IS NOT A DIRECTORY!
' ** Handled in frmBackup.

Private lngBackup_OptDVDLbl_Width As Long, lngRestore_OptDVDLbl_Width As Long
Private blnJustClose As Boolean, blnDone As Boolean
Private strMode As String, strModeCmd As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer

110     blnJustClose = False: blnDone = False

120     With Me
130       If IsNull(.OpenArgs) = False Then
140         If Trim(.OpenArgs) <> vbNullString Then
150           strCallingForm = Trim(.OpenArgs)
160           intPos01 = InStr(strCallingForm, "~")
170           If intPos01 > 0 Then

180             strModeCmd = Mid(strCallingForm, (intPos01 + 1))
190             strCallingForm = Left(strCallingForm, (intPos01 - 1))

200             Select Case strModeCmd
                Case "cmdBackupData_Click", "cmdQuit_Click"
210               strMode = "Backup"
220             Case "cmdRestoreData_Click"
230               strMode = "Restore"
240             End Select

250             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

260             lngBackup_OptDVDLbl_Width = .opgLocBackup_optDVD_lbl.Width  ' ** Non-Bold width.
270             lngRestore_OptDVDLbl_Width = .opgLocRestore_optDVD_lbl.Width

280             Select Case strMode
                Case "Backup"
290               If .opgLocBackup = 0 Then .opgLocBackup = .opgLocBackup_optDrive.OptionValue
300               opgLocBackup_AfterUpdate  ' ** Procedure: Below.
310               opgLocRestore_AfterUpdate  ' ** Procedure: Below.
320               .opgLocRestore.Enabled = False
330               .cmdRestore.Enabled = False
340               .opgLocRestore_box1.BackStyle = acBackStyleTransparent
350               .opgLocRestore_box2.BackStyle = acBackStyleTransparent
360               .opgLocRestore_hline03.BorderColor = .BackupRestore_box.BackColor
370             Case "Restore"
380               If .opgLocRestore = 0 Then .opgLocRestore = .opgLocRestore_optDrive.OptionValue
390               opgLocBackup_AfterUpdate  ' ** Procedure: Below.
400               opgLocRestore_AfterUpdate  ' ** Procedure: Below.
410               .opgLocBackup.Enabled = False
420               .cmdBackup.Enabled = False
430               .opgLocBackup_box1.BackStyle = acBackStyleTransparent
440               .opgLocBackup_box2.BackStyle = acBackStyleTransparent
450               .opgLocBackup_hline03.BorderColor = .BackupRestore_box.BackColor
460             End Select

470           Else
480             Cancel = -1
490             strCallingForm = vbNullString
500           End If
510         Else
520           Cancel = -1
530         End If
540       Else
550         Cancel = -1
560       End If
570     End With

580     If Cancel = -1 Then
590       Beep
600       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
610     End If

EXITP:
620     Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Drive:            D {opgLocBackup_optDrive}
        ' **   Drive:            D {opgLocRestore_optDrive}
        ' **   Backup Now:       N {cmdBackup}
        ' **   Restore Now:      N {cmdRestore}
        ' **   CD/DVD:           V {opgLocBackup_optDVD}
        ' **   CD/DVD:           V {opgLocRestore_optDVD}
        ' **   Exit:             X {cmdClose}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         intRetVal = 0
780         cmdClose_Click  ' ** Procedure: Below.
790       End Select
800     End If

        ' ** Alt keys.
810     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
820       Select Case intRetVal
          Case vbKeyD
830         With Me
840           intRetVal = 0
850           If .opgLocBackup.Enabled = True Then
860             .opgLocBackup = .opgLocBackup_optDrive.OptionValue
870             .opgLocBackup.SetFocus
880             opgLocBackup_AfterUpdate  ' ** Procedure: Below.
890           ElseIf .opgLocRestore.Enabled = True Then
900             .opgLocRestore = .opgLocRestore_optDrive.OptionValue
910             .opgLocRestore.SetFocus
920             opgLocRestore_AfterUpdate  ' ** Procedure: Below.
930           End If
940         End With
950       Case vbKeyN
960         With Me
970           intRetVal = 0
980           If .cmdBackup.Enabled = True Then
990             .cmdBackup.SetFocus
1000            cmdBackup_Click  ' ** Procedure: Below.
1010          ElseIf .cmdRestore.Enabled = True Then
1020            .cmdRestore.SetFocus
1030            cmdRestore_Click  ' ** Procedure: Below.
1040          End If
1050        End With
1060      Case vbKeyV
1070        With Me
1080          intRetVal = 0
1090          If .opgLocBackup.Enabled = True Then
1100            .opgLocBackup = .opgLocBackup_optDVD.OptionValue
1110            .opgLocBackup.SetFocus
1120            opgLocBackup_AfterUpdate  ' ** Procedure: Below.
1130          ElseIf .opgLocRestore.Enabled = True Then
1140            .opgLocRestore = .opgLocRestore_optDVD.OptionValue
1150            .opgLocRestore.SetFocus
1160            opgLocRestore_AfterUpdate  ' ** Procedure: Below.
1170          End If
1180        End With
1190      Case vbKeyX
1200        intRetVal = 0
1210        cmdClose_Click  ' ** Procedure: Below.
1220      End Select
1230    End If

EXITP:
1240    KeyCode = intRetVal
1250    Exit Sub

ERRH:
1260    intRetVal = 0
1270    Select Case ERR.Number
        Case Else
1280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1290    End Select
1300    Resume EXITP

End Sub

Private Sub Form_Timer()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1410    With Me
1420      .TimerInterval = 0&
1430      If gblnSetFocus = True Then
1440        gblnSetFocus = False
1450        strFormName = vbNullString
1460  On Error Resume Next
1470        Do Until strFormName = THIS_NAME
1480          DoCmd.SelectObject acForm, THIS_NAME, False
1490          strFormName = Screen.ActiveForm.Name
1500        Loop
1510  On Error GoTo ERRH
1520      End If
1530      Select Case strMode
          Case "backup"
1540        .cmdBackup.SetFocus
1550      Case "restore"
1560        .cmdRestore.SetFocus
1570      End Select
1580    End With

EXITP:
1590    Exit Sub

ERRH:
1600    Select Case ERR.Number
        Case Else
1610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1620    End Select
1630    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1710    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1720    If blnJustClose = False Then  ' ** True for frmBackupRestore_File.
1730      gstrReturningForm = THIS_NAME
1740      gblnSetFocus = True
1750      Select Case strCallingForm
          Case "frmMenu_Utility"
1760  On Error Resume Next
1770        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
1780  On Error GoTo ERRH
1790      Case "frmMenu_Title"
1800        Select Case blnDone
            Case True
1810          QuitNow  ' ** Module Procedure: modStartupFuncs.
1820        Case False
              ' ** Handled by frmBackupRestore_File, though blnJustClose keeps it from ever getting here.
1830        End Select
1840      Case Else
            ' ** Nothing at the moment.
1850      End Select
1860    End If

EXITP:
1870    Exit Sub

ERRH:
1880    Select Case ERR.Number
        Case 2486  ' ** You can't carry out this action at the present time.
          ' ** Ignore.
1890    Case Else
1900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1910    End Select
1920    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2010    blnDone = True  ' ** Only pertinent if coming from frmMenu_Title.
2020    DoCmd.Close acForm, THIS_NAME

EXITP:
2030    Exit Sub

ERRH:
2040    Select Case ERR.Number
        Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyTab
2170        With Me
2180          intRetVal = 0
2190          If .opgLocBackup.Enabled = True Then
2200            .opgLocBackup.SetFocus
2210          Else
2220            .opgLocRestore.SetFocus
2230          End If
2240        End With
2250      End Select
2260    End If

        ' ** Shift keys.
2270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2280      Select Case intRetVal
          Case vbKeyTab
2290        With Me
2300          intRetVal = 0
2310          If .cmdRestore.Enabled = True Then
2320            .cmdRestore.SetFocus
2330          Else
2340            .cmdBackup.SetFocus
2350          End If
2360        End With
2370      End Select
2380    End If

EXITP:
2390    KeyCode = intRetVal
2400    Exit Sub

ERRH:
2410    intRetVal = 0
2420    Select Case ERR.Number
        Case Else
2430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2440    End Select
2450    Resume EXITP

End Sub

Private Sub opgLocBackup_AfterUpdate()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocBackup_AfterUpdate"

2510    With Me
2520      Select Case .opgLocBackup
          Case .opgLocBackup_optDrive.OptionValue
2530        .opgLocBackup_optDrive_lbl.FontBold = True
2540        .opgLocBackup_optDVD_lbl.FontBold = False
2550      Case .opgLocBackup_optDVD.OptionValue
2560        .opgLocBackup_optDrive_lbl.FontBold = False
2570        .opgLocBackup_optDVD_lbl.FontBold = True
2580      Case Else
2590        .opgLocBackup_optDrive_lbl.FontBold = False
2600        .opgLocBackup_optDVD_lbl.FontBold = False
2610      End Select
2620    End With

EXITP:
2630    Exit Sub

ERRH:
2640    Select Case ERR.Number
        Case Else
2650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2660    End Select
2670    Resume EXITP

End Sub

Private Sub opgLocBackup_optDrive_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocBackup_optDrive_KeyDown"

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
2790          .cmdBackup.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2850        With Me
2860          intRetVal = 0
2870          .cmdClose.SetFocus
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

Private Sub opgLocBackup_optDVD_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocBackup_optDVD_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3020    intShiftDown = (Shift And acShiftMask) > 0
3030    intAltDown = (Shift And acAltMask) > 0
3040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3070        With Me
3080          intRetVal = 0
3090          .cmdBackup.SetFocus
3100        End With
3110      End Select
3120    End If

        ' ** Shift keys.
3130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3150        With Me
3160          intRetVal = 0
3170          .cmdClose.SetFocus
3180        End With
3190      End Select
3200    End If

EXITP:
3210    KeyCode = intRetVal
3220    Exit Sub

ERRH:
3230    intRetVal = 0
3240    Select Case ERR.Number
        Case Else
3250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3260    End Select
3270    Resume EXITP

End Sub

Private Sub cmdBackup_Click()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBackup_Click"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf As DAO.QueryDef, tdf As DAO.TableDef
        Dim strDataPathFile As String
        Dim strSQL As String
        Dim blnRetVal As Boolean

3310    With Me

          ' ** Set path independent of globals to be sure of accuracy.
3320      strDataPathFile = CurrentBackendPathFile("Ledger")  ' ** Module Function: modFileUtilities.

3330      Select Case .opgLocBackup
          Case .opgLocBackup_optDrive.OptionValue

3340        blnJustClose = True
3350        DoCmd.OpenForm "frmBackupRestore_File", , , , , , (THIS_NAME & "~" & strModeCmd & "~" & strCallingForm)
3360        DoCmd.Close acForm, THIS_NAME

3370      Case .opgLocBackup_optDVD.OptionValue

3380        Beep
3390        MsgBox "Temporarily unavailable." & vbCrLf & "Use the Drive option instead.", _
              vbInformation + vbOKOnly, "CD/DVD Option Temporarily Unavailable"
3400        blnRetVal = False

3410        If blnRetVal = True Then
              'THIS IS CURRENTLY NOT USED!
              'NEEDS TO BE COMPLETELY REPLACED!

              ' ** Temporarily bring the archive table local.
3420          strSQL = "SELECT LedgerArchive.* INTO LedgerArchive_Backup IN '" & strDataPathFile & "' FROM LedgerArchive;"
              ' ** VGC 07/30/08: A user is repeatedly getting this error:
              ' **   3010: Table 'LedgerArchive_Backup' already exists.
              ' ** It's in v2.1.47 (one of our odd ones), and the error message
              ' ** gives no location details. So, I'm covering the bases with this.

3430  On Error Resume Next
3440          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
3450          If ERR.Number <> 0 Then
3460  On Error GoTo ERRH
3470  On Error Resume Next
3480            Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
3490            If ERR.Number <> 0 Then
3500  On Error GoTo ERRH
3510  On Error Resume Next
3520              Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
3530              If ERR.Number <> 0 Then
3540  On Error GoTo ERRH
3550  On Error Resume Next
3560                Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
3570                If ERR.Number <> 0 Then
3580  On Error GoTo ERRH
3590  On Error Resume Next
3600                  Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
3610                  If ERR.Number <> 0 Then
3620  On Error GoTo ERRH
3630  On Error Resume Next
3640                    Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
3650                    If ERR.Number <> 0 Then
3660  On Error GoTo ERRH
3670  On Error Resume Next
3680                      Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
3690  On Error GoTo ERRH
3700                    Else
3710  On Error GoTo ERRH
3720                    End If
3730                  Else
3740  On Error GoTo ERRH
3750                  End If
3760                Else
3770  On Error GoTo ERRH
3780                End If
3790              Else
3800  On Error GoTo ERRH
3810              End If
3820            Else
3830  On Error GoTo ERRH
3840            End If
3850          Else
3860  On Error GoTo ERRH
3870          End If

3880          With wrk
3890            Set dbs = .OpenDatabase(strDataPathFile, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
3900            With dbs
3910              For Each tdf In .TableDefs
3920                With tdf
3930                  If .Name = "LedgerArchive_Backup" Then
3940                    dbs.TableDefs.Delete "LedgerArchive_Backup"
3950                    Exit For
3960                  End If
3970                End With
3980              Next
3990              .Close
4000            End With
4010            .Close
4020          End With

4030          Set dbs = CurrentDb
4040          With dbs
4050  On Error Resume Next
                ' ** This creates a table called LedgerArchive_Backup in the current TrustDta.mdb.
4060            .Execute strSQL
4070  On Error GoTo ERRH
4080            .Close
4090          End With
4100          Set dbs = Nothing
4110  On Error GoTo ERRH
              'blnRetVal = BackupToFloppy(strDataPathFile)  ' ** Module Function: modBackup.
4120  On Error Resume Next
4130          Set dbs = DBEngine.Workspaces(0).OpenDatabase(strDataPathFile)
              ' ** After the current TrustDta.mdb is backed-up to the CD/DVD, the LedgerArchive_Backup table is deleted.
4140          dbs.TableDefs.Delete "LedgerArchive_Backup"
4150          dbs.Close
4160  On Error GoTo ERRH

4170          blnDone = True
4180          DoCmd.Close acForm, THIS_NAME

4190        End If
4200      End Select

4210    End With

EXITP:
4220    Set qdf = Nothing
4230    Set tdf = Nothing
4240    Set dbs = Nothing
4250    Set wrk = Nothing
4260    Exit Sub

ERRH:
4270    Select Case ERR.Number
        Case Else
4280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4290    End Select
4300    Resume EXITP

End Sub

Private Sub cmdBackup_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdBackup_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4420    intShiftDown = (Shift And acShiftMask) > 0
4430    intAltDown = (Shift And acAltMask) > 0
4440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4460      Select Case intRetVal
          Case vbKeyTab
4470        With Me
4480          intRetVal = 0
4490          .cmdClose.SetFocus
4500        End With
4510      End Select
4520    End If

        ' ** Shift keys.
4530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4540      Select Case intRetVal
          Case vbKeyTab
4550        With Me
4560          intRetVal = 0
4570          .opgLocBackup.SetFocus
4580        End With
4590      End Select
4600    End If

EXITP:
4610    KeyCode = intRetVal
4620    Exit Sub

ERRH:
4630    intRetVal = 0
4640    Select Case ERR.Number
        Case Else
4650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4660    End Select
4670    Resume EXITP

End Sub

Private Sub opgLocRestore_AfterUpdate()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocRestore_AfterUpdate"

4710    With Me
4720      Select Case .opgLocRestore
          Case .opgLocRestore_optDrive.OptionValue
4730        .opgLocRestore_optDrive_lbl.FontBold = True
4740        .opgLocRestore_optDVD_lbl.FontBold = False
4750      Case .opgLocRestore_optDVD.OptionValue
4760        .opgLocRestore_optDrive_lbl.FontBold = False
4770        .opgLocRestore_optDVD_lbl.FontBold = True
4780      Case Else
4790        .opgLocRestore_optDrive_lbl.FontBold = False
4800        .opgLocRestore_optDVD_lbl.FontBold = False
4810      End Select
4820    End With

EXITP:
4830    Exit Sub

ERRH:
4840    Select Case ERR.Number
        Case Else
4850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860    End Select
4870    Resume EXITP

End Sub

Private Sub opgLocRestore_optDrive_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocRestore_optDrive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4920    intShiftDown = (Shift And acShiftMask) > 0
4930    intAltDown = (Shift And acAltMask) > 0
4940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4970        With Me
4980          intRetVal = 0
4990          .cmdRestore.SetFocus
5000        End With
5010      End Select
5020    End If

        ' ** Shift keys.
5030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5050        With Me
5060          intRetVal = 0
5070          .cmdClose.SetFocus
5080        End With
5090      End Select
5100    End If

EXITP:
5110    KeyCode = intRetVal
5120    Exit Sub

ERRH:
5130    intRetVal = 0
5140    Select Case ERR.Number
        Case Else
5150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5160    End Select
5170    Resume EXITP

End Sub

Private Sub opgLocRestore_optDVD_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgLocRestore_optDVD_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
5270        With Me
5280          intRetVal = 0
5290          .cmdRestore.SetFocus
5300        End With
5310      End Select
5320    End If

        ' ** Shift keys.
5330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5350        With Me
5360          intRetVal = 0
5370          .cmdClose.SetFocus
5380        End With
5390      End Select
5400    End If

EXITP:
5410    KeyCode = intRetVal
5420    Exit Sub

ERRH:
5430    intRetVal = 0
5440    Select Case ERR.Number
        Case Else
5450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5460    End Select
5470    Resume EXITP

End Sub

Private Sub cmdRestore_Click()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRestore_Click"

        Dim blnRetVal As Boolean

5510    With Me
5520      Select Case .opgLocRestore
          Case .opgLocRestore_optDrive.OptionValue
5530        blnRetVal = RestoreFromDrive  ' ** Module Procedure: modBackup.
5540      Case .opgLocRestore_optDVD.OptionValue
5550        Beep
5560        MsgBox "Temporarily unavailable." & vbCrLf & "Use the Drive option instead.", _
              vbInformation + vbOKOnly, "CD/DVD Option Temporarily Unavailable"
5570        blnRetVal = False
            'blnRetVal = RestoreFromFloppy THIS_NAME  ' ** Module Procedure: modBackup.
5580      End Select
5590    End With

5600    DoCmd.Hourglass False

5610    If blnRetVal = True Then
5620      DoCmd.Close acForm, THIS_NAME
5630    End If

EXITP:
5640    Exit Sub

ERRH:
5650    Select Case ERR.Number
        Case Else
5660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5670    End Select
5680    Resume EXITP

End Sub

Private Sub cmdRestore_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRestore_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5720    intShiftDown = (Shift And acShiftMask) > 0
5730    intAltDown = (Shift And acAltMask) > 0
5740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab
5770        With Me
5780          intRetVal = 0
5790          .cmdClose.SetFocus
5800        End With
5810      End Select
5820    End If

        ' ** Shift keys.
5830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5840      Select Case intRetVal
          Case vbKeyTab
5850        With Me
5860          intRetVal = 0
5870          .opgLocRestore.SetFocus
5880        End With
5890      End Select
5900    End If

EXITP:
5910    KeyCode = intRetVal
5920    Exit Sub

ERRH:
5930    intRetVal = 0
5940    Select Case ERR.Number
        Case Else
5950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5960    End Select
5970    Resume EXITP

End Sub
