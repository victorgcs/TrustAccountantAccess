VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_MICR_Adjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_MICR_Adjust"

'VGC 07/03/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Move Down:        D {cmdMoveDown}
' **   Move Left:        L {cmdMoveLeft}
' **   Start Over:       O {cmdStartOver}
' **   Move Right:       R {cmdMoveRight}
' **   Save:             S {cmdSaveAdjustments}
' **   Print Test:       T {cmdPrintChecksTest}
' **   Move Up:          U {cmdMoveUp}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSaveAdjustments}

Private lngOpenTop As Long, lngOpen_Left As Long
Private lngStart_Top As Long, lngAdjust_Top As Long
Private lngStart_Left As Long, lngAdjust_Left As Long
Private lngAdjustByAmount As Long
Private blnAdjustNotSet As Boolean, blnChanged As Boolean, blnGoingToReport As Boolean, blnChecksCalled As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnChanged = False: blnChecksCalled = False
160           MICRChanged False  ' ** Procedure: Below.

170           lngOpenTop = .Moveable_lbl.Top  ' ** 1 Pixel below Fixed_lbl.Top, because it's border is transparent.
180           lngOpen_Left = .Moveable_lbl.Left  ' ** 1 Pixel right of Fixed_lbl.Left, because it's border is transparent.

              ' ** Get the stored offset from DDTrust.INI and convert back to pixels.
190           lngAdjustByAmount = CLng(IniFile_Get("CheckMICROffset", "AdjustBy", RET_ERR, CurrentAppPath & LNK_SEP & gstrFile_INI))  ' ** Module Function: modStartupFuncs, modFileUtilities.
200           If lngAdjustByAmount = 0& Then
210             lngAdjustByAmount = GetTPP  ' ** Module Function: modWindowFunctions.
220             blnAdjustNotSet = True
230           Else
240             blnAdjustNotSet = False
250           End If
260           lngAdjust_Top = CLng(IniFile_Get("CheckMICROffset", "Top", RET_ERR, CurrentAppPath & LNK_SEP & gstrFile_INI))  ' ** Module Function: modStartupFuncs, modFileUtilities.
270           .Moveable_lbl.Top = .Moveable_lbl.Top + lngAdjust_Top
280           lngAdjust_Left = CLng(IniFile_Get("CheckMICROffset", "Left", RET_ERR, CurrentAppPath & LNK_SEP & gstrFile_INI))  ' ** Module Function: modStartupFuncs, modFileUtilities.
290           .Moveable_lbl.Left = .Moveable_lbl.Left + lngAdjust_Left
300           lngStart_Top = .Moveable_lbl.Top  ' ** Minus 15 Twips for real relationship to Fixed_lbl.
310           lngStart_Left = .Moveable_lbl.Left  ' ** Minus 15 Twips for real relationship to Fixed_lbl.

320           .AdjustTop_Saved = lngAdjust_Top
330           .AdjustLeft_Saved = lngAdjust_Left
340           .AdjustByAmount = lngAdjustByAmount
350           .AdjustTop_Current = lngAdjust_Top
360           .AdjustLeft_Current = lngAdjust_Left

370         Else
380           Cancel = -1
390         End If
400       Else
410         Cancel = -1
420       End If
430     End With

440     If Cancel = -1 Then
450       Beep
460       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
470     End If

480     DoCmd.Hourglass False

EXITP:
490     Exit Sub

ERRH:
500     DoCmd.Hourglass False
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

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSaveAdjustments}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyEscape
670         intRetVal = 0
680         cmdClose_Click  ' ** Procedure: Below.
690       Case vbKeyUp
700         With Me
710           intRetVal = 0
720           .cmdMoveUp.SetFocus
730           cmdMoveUp_Click  ' ** Procedure: Below.
740         End With
750       Case vbKeyLeft
760         With Me
770           intRetVal = 0
780           .cmdMoveLeft.SetFocus
790           cmdMoveLeft_Click  ' ** Procedure: Below.
800         End With
810       Case vbKeyDown
820         With Me
830           intRetVal = 0
840           .cmdMoveDown.SetFocus
850           cmdMoveDown_Click  ' ** Procedure: Below.
860         End With
870       Case vbKeyRight
880         With Me
890           intRetVal = 0
900           .cmdMoveRight.SetFocus
910           cmdMoveRight_Click  ' ** Procedure: Below.
920         End With
930       End Select
940     End If

        ' ** Alt keys.
950     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyX
970         intRetVal = 0
980         cmdClose_Click  ' ** Procedure: Below.
990       End Select
1000    End If

        ' ** Ctrl keys.
1010    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1020      Select Case intRetVal
          Case vbKeyS
1030        intRetVal = 0
1040        cmdSaveAdjustments_Click  ' ** Procedure: Below.
1050      End Select
1060    End If

        ' ** Ctrl-Shift keys.
1070    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1080      Select Case intRetVal
          Case vbKeyF
1090        With Me
1100          intRetVal = 0
1110          .FocusHolder.SetFocus
1120        End With
1130      End Select
1140    End If

EXITP:
1150    KeyCode = intRetVal
1160    Exit Sub

ERRH:
1170    intRetVal = 0
1180    Select Case ERR.Number
        Case Else
1190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1200    End Select
1210    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1310    With Me
1320      .TimerInterval = 0&
1330      If gblnSetFocus = True = True Then
1340        gblnSetFocus = False
1350        strFormName = vbNullString
1360  On Error Resume Next
1370        Do Until strFormName = THIS_NAME
1380          DoCmd.SelectObject acForm, THIS_NAME
1390          strFormName = Screen.ActiveForm.Name
1400        Loop
1410  On Error GoTo ERRH
1420      End If
1430      If gblnGoToReport = True Then
1440        Select Case blnGoingToReport
            Case True
1450          DoCmd.Hourglass False  ' ** It should already be off.
1460          blnGoingToReport = False
1470          gblnGoToReport = False
1480          .GoToReport_arw_micradj_img.Visible = False
1490          .cmdSaveAdjustments.Visible = True
1500          DoEvents
1510        Case False
1520          blnGoingToReport = True
1530          .cmdPrintChecksTest.SetFocus
1540          .cmdSaveAdjustments.Visible = False
1550          .GoToReport_arw_micradj_img.Visible = True
1560          DoEvents
1570          Beep
1580          DoCmd.Hourglass False
1590          DoEvents
1600          .TimerInterval = (GTR_WAIT * 2&)
1610        End Select
1620      End If
1630    End With

EXITP:
1640    Exit Sub

ERRH:
1650    DoCmd.Hourglass False
1660    Select Case ERR.Number
        Case Else
1670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1680    End Select
1690    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1710    If strCallingForm <> vbNullString Then
1720      gblnSetFocus = True
1730      Forms(strCallingForm).TimerInterval = 250&
1740    End If

EXITP:
1750    Exit Sub

ERRH:
1760    Select Case ERR.Number
        Case Else
1770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1780    End Select
1790    Resume EXITP

End Sub

Private Sub cmdClose_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim msgResponse As VbMsgBoxResult

1810    With Me

1820      If .Moveable_lbl.Top <> lngStart_Top Or .Moveable_lbl.Left <> lngStart_Left Then
1830        msgResponse = MsgBox("You have made changes that you have not saved." & vbCrLf & vbCrLf & _
              "Are you sure you want to exit without saving?", vbQuestion + vbYesNo + vbDefaultButton2, "Changes Not Saved")
1840      Else
1850        msgResponse = vbYes
1860      End If

1870      If msgResponse = vbYes Then

1880        If Forms(strCallingForm).Visible = False Then
1890          Forms(strCallingForm).Visible = True
1900        End If

1910        If blnChecksCalled = True Then
1920          Set dbs = CurrentDb
1930          With dbs
                ' ** Empty tmpJournal.
1940            Set qdf = .QueryDefs("qryRpt_Checks_MICR_Adjust_04")
1950            qdf.Execute
1960            Set qdf = Nothing
1970            .Close
1980          End With
1990        End If

2000        DoCmd.Close acForm, THIS_NAME, acSaveNo

2010      Else
2020        .cmdSaveAdjustments.SetFocus
2030      End If

2040    End With

EXITP:
2050    Set qdf = Nothing
2060    Set dbs = Nothing
2070    Exit Sub

ERRH:
2080    Select Case ERR.Number
        Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyTab
2270        With Me
2280          intRetVal = 0
2290          .cmdMoveUp.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab
2350        With Me
2360          intRetVal = 0
2370          .cmdStartOver.SetFocus
2380        End With
2390      End Select
2400    End If

EXITP:
2410    KeyCode = intRetVal
2420    Exit Sub

ERRH:
2430    intRetVal = 0
2440    Select Case ERR.Number
        Case Else
2450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2460    End Select
2470    Resume EXITP

End Sub

Private Sub cmdPrintChecksTest_Click()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintChecksTest_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String, strShortName As String, strJournalUser As String
        Dim strBankName As String, strBankCity As String, strBankState As String
        Dim strBankAcctNum As String, strBankRteNum As String, strFedNum1 As String, strFedNum2 As String
        Dim strRecurItem As String, strRecurAddress As String, strRecurCity As String, strRecurState As String, strRecurZip As String
        Dim datTransDate As Date, strJComment As String, strCheckMemo As String
        Dim lngJournalNo As Long, lngCheckNum As Long, curAmount As Currency
        Dim strDocName As String, strReportCriteria As String, strMICRFont As String, intMode As Integer

2510    With Me

2520      DoCmd.Hourglass True
2530      DoEvents

2540      blnChecksCalled = True

2550      If .cmdSaveAdjustments.Enabled = True Then
2560        cmdSaveAdjustments_Click  ' ** Procedure: Below.
2570        DoEvents
2580        .cmdPrintChecksTest.SetFocus
2590      End If

          ' ** Find a representative Ledger entry to base the test check on.
2600      Set dbs = CurrentDb
2610      With dbs

2620        strAccountNo = vbNullString: lngJournalNo = 0&: lngCheckNum = 0&

            ' ** qryRpt_Checks_MICR_Adjust_02a (Ledger, just CheckPaid = True,
            ' ** CheckNum <> Null), grouped by accountno, with cnt.
2630        Set qdf = .QueryDefs("qryRpt_Checks_MICR_Adjust_02b")
2640        Set rst = qdf.OpenRecordset
2650        With rst
2660          If .BOF = True And .EOF = True Then
                ' ** We'll have to invent something.
2670          Else
2680            .MoveFirst
2690            strAccountNo = ![accountno]
2700            lngJournalNo = ![journalno]
2710            Select Case IsNull(![CheckNum])
                Case True
2720              lngCheckNum = 12345
2730            Case False
2740              lngCheckNum = (![CheckNum] + 1&)
2750            End Select
2760          End If
2770          .Close
2780        End With  ' ** rst.
2790        Set rst = Nothing
2800        Set qdf = Nothing

2810        If strAccountNo = vbNullString Then
              ' ** Account, just accountno = 'SUSPENSE'.
2820          Set qdf = .QueryDefs("qryRpt_Checks_MICR_Adjust_02c")
2830          Set rst = qdf.OpenRecordset
2840          With rst
2850            .MoveFirst
2860            strAccountNo = ![accountno]
2870            lngJournalNo = ![journalno]  ' ** Max(journalno) + 1.
2880            lngCheckNum = ![CheckNum]  ' ** 12345.
2890            .Close
2900          End With  ' ** rst.
2910          Set rst = Nothing
2920          Set qdf = Nothing
2930        End If

            ' ** Account, just needed fields, by specified [actno].
2940        Set qdf = .QueryDefs("qryRpt_Checks_MICR_Adjust_03")
2950        With qdf.Parameters
2960          ![actno] = strAccountNo
2970        End With
2980        Set rst = qdf.OpenRecordset
2990        With rst
3000          .MoveFirst
3010          strShortName = ![shortname]
3020          Select Case IsNull(![Bank_Name])
              Case True
3030            strBankName = "North Fork Bank"
3040          Case False
3050            strBankName = ![Bank_Name]
3060          End Select
3070          Select Case IsNull(![Bank_City])
              Case True
3080            strBankCity = "North Fork"
3090          Case False
3100            strBankCity = ![Bank_City]
3110          End Select
3120          Select Case IsNull(![Bank_State])
              Case True
3130            strBankState = "MN"
3140          Case False
3150            strBankState = ![Bank_State]
3160          End Select
3170          Select Case IsNull(![Bank_AccountNumber])
              Case True
3180            strBankAcctNum = "1234567890"
3190          Case False
3200            strBankAcctNum = ![Bank_AccountNumber]
3210          End Select
3220          Select Case IsNull(![Bank_RoutingNumber])
              Case True
3230            strBankRteNum = "123456789"
3240          Case False
3250            strBankRteNum = ![Bank_RoutingNumber]
3260          End Select
3270          Select Case IsNull(![FedIFNum1])
              Case True
3280            strFedNum1 = "18-7002"
3290          Case False
3300            strFedNum1 = ![FedIFNum1]
3310          End Select
3320          Select Case IsNull(![FedIFNum2])
              Case True
3330            strFedNum2 = "3910"
3340          Case False
3350            strFedNum2 = ![FedIFNum2]
3360          End Select
3370          .Close
3380        End With  ' ** rst.
3390        Set rst = Nothing
3400        Set qdf = Nothing

3410        strRecurItem = "Hilcrest Care Center"
3420        strRecurAddress = "2341 East Main Street"
3430        strRecurCity = "North Fork"
3440        strRecurState = "MN"
3450        strRecurZip = "553241131"
3460        datTransDate = Date
3470        strJComment = "Trustee Health Care"
3480        strCheckMemo = "Payment for " & Format(DateAdd("m", 1, Date), "mmmm yyyy")
3490        strJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
3500        curAmount = CCur(-1234.56)

            ' ** Empty tmpJournal.
3510        Set qdf = .QueryDefs("qryRpt_Checks_MICR_Adjust_04")
3520        qdf.Execute
3530        Set qdf = Nothing

3540        Set rst = .OpenRecordset("tmpJournal", dbOpenDynaset, dbConsistent)
3550        With rst
3560          .AddNew
              ' ** ![jrnl_id] : AutoNumber.
3570          ![accountno] = strAccountNo
3580          ![shortname] = strShortName
3590          ![Bank_Name] = strBankName
3600          ![Bank_City] = strBankCity
3610          ![Bank_State] = strBankState
3620          ![Bank_AccountNumber] = strBankAcctNum
3630          ![Bank_RoutingNumber] = strBankRteNum
3640          ![FedIFNum1] = strFedNum1
3650          ![FedIFNum2] = strFedNum2
3660          ![transdate] = datTransDate
3670          ![amount] = curAmount
3680          ![CheckNum] = lngCheckNum
3690          ![RecurringItem] = strRecurItem
3700          ![RecurringItem_Address] = strRecurAddress
3710          ![RecurringItem_City] = strRecurCity
3720          ![RecurringItem_State] = strRecurState
3730          ![RecurringItem_Zip] = strRecurZip
3740          ![jcomment] = strJComment
3750          ![JrnlMemo_Memo] = strCheckMemo
3760          ![journal_USER] = strJournalUser
3770          ![jrnl_datemodified] = Now()
3780          .Update
3790          .Close
3800        End With
3810        Set rst = Nothing

3820        .Close
3830      End With  ' ** dbs.
3840      Set dbs = Nothing

          ' ** Sort order options:
          ' **   Account_Number
          ' **   Short_Name
          ' **   Payee
          ' **   Bank_Name
          ' **   Bank_Account_Number

3850      strMICRFont = IniFile_Get("General", "MICRFont", "#ERROR", CurrentAppPath & LNK_SEP & gstrFile_INI)  ' ** Module Procedure: modStartupFuncs, modFileUtilities.
3860      If strMICRFont <> "#ERROR" Then
3870        .MICRFontName = strMICRFont
3880      End If

3890      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
3900        intMode = acViewPreview
3910      Else
3920        intMode = acViewNormal
3930      End If

3940      DoCmd.Hourglass False
3950      strDocName = "rptChecks_Blank"
3960      strReportCriteria = strReportCriteria & "[accountno] = '" & strAccountNo & "'"  ' ** Not really necessary here.
          '##GTR_Ref: rptChecks_Blank
3970      DoCmd.OpenReport strDocName, intMode, , strReportCriteria, , "Test_Check" & "~" & "False~Test"
          ' ** OpenArgs: SortOrder ~ Descending.

3980      If intMode = acViewPreview Then
3990        DoCmd.Maximize
4000        DoCmd.RunCommand acCmdFitToWindow
4010        Forms(THIS_NAME).Visible = False
4020        Forms(strCallingForm).Visible = False
4030      End If

4040    End With

EXITP:
4050    Set rst = Nothing
4060    Set qdf = Nothing
4070    Set dbs = Nothing
4080    Exit Sub

ERRH:
4090    DoCmd.Hourglass False
4100    Select Case ERR.Number
        Case Else
4110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4120    End Select
4130    Resume EXITP

End Sub

Private Sub cmdPrintChecksTest_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintChecksTest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4220    intShiftDown = (Shift And acShiftMask) > 0
4230    intAltDown = (Shift And acAltMask) > 0
4240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4260      Select Case intRetVal
          Case vbKeyTab
4270        With Me
4280          intRetVal = 0
4290          If .cmdSaveAdjustments.Enabled = True Then
4300            .cmdSaveAdjustments.SetFocus
4310          ElseIf .cmdStartOver.Enabled = True Then
4320            .cmdStartOver.SetFocus
4330          Else
4340            .cmdClose.SetFocus
4350          End If
4360        End With
4370      End Select
4380    End If

        ' ** Shift keys.
4390    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4400      Select Case intRetVal
          Case vbKeyTab
4410        With Me
4420          intRetVal = 0
4430          .cmdMoveRight.SetFocus
4440        End With
4450      End Select
4460    End If

EXITP:
4470    KeyCode = intRetVal
4480    Exit Sub

ERRH:
4490    intRetVal = 0
4500    Select Case ERR.Number
        Case Else
4510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4520    End Select
4530    Resume EXITP

End Sub

Private Sub cmdSaveAdjustments_Click()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaveAdjustments_Click"

        Dim lngTmpAdjustTop As Long, lngTmpAdjustLeft As Long

        ' ** Convert from pixels to inches and store offset in DDTrust.INI.
        ' ** VGC 11/05/2010: Though this may have been the original intent,
        ' ** right now it's saving in Twips.
4610    With Me

4620      lngTmpAdjustTop = (.Moveable_lbl.Top - .Fixed_lbl.Top) - lngAdjustByAmount  ' ** Because of the border business.
4630      IniFile_Set "CheckMICROffset", "Top", CStr(lngTmpAdjustTop), CurrentAppPath & LNK_SEP & gstrFile_INI  ' ** Module Function: modStartupFuncs, modFileUtilities.
4640      lngTmpAdjustLeft = (.Moveable_lbl.Left - .Fixed_lbl.Left) - lngAdjustByAmount  ' ** Because of the border business.
4650      IniFile_Set "CheckMICROffset", "Left", CStr(lngTmpAdjustLeft), CurrentAppPath & LNK_SEP & gstrFile_INI  ' ** Module Function: modStartupFuncs, modFileUtilities.
4660      IniFile_Set "CheckMICROffset", "ScreenFontSize", CStr(.Moveable_lbl.FontSize), CurrentAppPath & LNK_SEP & gstrFile_INI  ' ** Module Function: modStartupFuncs, modFileUtilities.
4670      If blnAdjustNotSet = True Then
4680        IniFile_Set "CheckMICROffset", "AdjustBy", CStr(lngAdjustByAmount), CurrentAppPath & LNK_SEP & gstrFile_INI  ' ** Module Function: modStartupFuncs, modFileUtilities.
4690      End If

4700      lngStart_Top = .Moveable_lbl.Top
4710      lngStart_Left = .Moveable_lbl.Left
4720      lngAdjust_Top = lngTmpAdjustTop
4730      lngAdjust_Left = lngTmpAdjustLeft

4740      .AdjustTop_Saved = lngAdjust_Top
4750      .AdjustLeft_Saved = lngAdjust_Left

4760      MICRChanged False  ' ** Procedure: Below.
4770      If .cmdPrintChecksTest.Enabled = True Then
4780        .cmdPrintChecksTest.SetFocus
4790      Else
4800        .cmdClose.SetFocus
4810      End If
4820      .cmdSaveAdjustments.Enabled = False
4830      .cmdStartOver.Enabled = False

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

Private Sub cmdSaveAdjustments_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSaveAdjustments_KeyDown"

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
          Case vbKeyTab
4970        With Me
4980          intRetVal = 0
4990          If .cmdStartOver.Enabled = True Then
5000            .cmdStartOver.SetFocus
5010          Else
5020            .cmdClose.SetFocus
5030          End If
5040        End With
5050      End Select
5060    End If

        ' ** Shift keys.
5070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5080      Select Case intRetVal
          Case vbKeyTab
5090        With Me
5100          intRetVal = 0
5110          If .cmdPrintChecksTest.Enabled = True Then
5120            .cmdPrintChecksTest.SetFocus
5130          Else
5140            .cmdMoveRight.SetFocus
5150          End If
5160        End With
5170      End Select
5180    End If

EXITP:
5190    KeyCode = intRetVal
5200    Exit Sub

ERRH:
5210    intRetVal = 0
5220    Select Case ERR.Number
        Case Else
5230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5240    End Select
5250    Resume EXITP

End Sub

Private Sub cmdStartOver_Click()

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStartOver_Click"

5310    With Me
5320      .Moveable_lbl.Top = lngOpenTop
5330      .Moveable_lbl.Left = lngOpen_Left
5340      lngAdjust_Top = 0&
5350      lngAdjust_Left = 0&
5360      .AdjustTop_Current = lngAdjust_Top
5370      .AdjustLeft_Current = lngAdjust_Left
5380      MICRChanged False  ' ** Procedure: Below.
5390      If .cmdPrintChecksTest.Enabled = True Then
5400        .cmdPrintChecksTest.SetFocus
5410      Else
5420        .cmdClose.SetFocus
5430      End If
5440      .cmdSaveAdjustments.Enabled = False
5450      .cmdStartOver.Enabled = False
5460    End With

EXITP:
5470    Exit Sub

ERRH:
5480    Select Case ERR.Number
        Case Else
5490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5500    End Select
5510    Resume EXITP

End Sub

Private Sub cmdStartOver_KeyDown(KeyCode As Integer, Shift As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStartOver_KeyDown"

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
5670        With Me
5680          intRetVal = 0
5690          .cmdClose.SetFocus
5700        End With
5710      End Select
5720    End If

        ' ** Shift keys.
5730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5740      Select Case intRetVal
          Case vbKeyTab
5750        With Me
5760          intRetVal = 0
5770          If .cmdSaveAdjustments.Enabled = True Then
5780            .cmdSaveAdjustments.SetFocus
5790          ElseIf .cmdPrintChecksTest.Enabled = True Then
5800            .cmdPrintChecksTest.SetFocus
5810          Else
5820            .cmdMoveRight.SetFocus
5830          End If
5840        End With
5850      End Select
5860    End If

EXITP:
5870    KeyCode = intRetVal
5880    Exit Sub

ERRH:
5890    intRetVal = 0
5900    Select Case ERR.Number
        Case Else
5910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5920    End Select
5930    Resume EXITP

End Sub

Private Sub cmdMoveUp_Click()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveUp_Click"

6010    With Me
6020      .Moveable_lbl.Top = .Moveable_lbl.Top - lngAdjustByAmount
6030      .AdjustTop_Current = .AdjustTop_Current - lngAdjustByAmount
6040      MICRChanged True  ' ** Procedure: Below.
6050    End With

EXITP:
6060    Exit Sub

ERRH:
6070    Select Case ERR.Number
        Case Else
6080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6090    End Select
6100    Resume EXITP

End Sub

Private Sub cmdMoveUp_DblClick(Cancel As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveUp_DblClick"

6210    With Me
6220      .Moveable_lbl.Top = .Moveable_lbl.Top - (lngAdjustByAmount * 3&)
6230      .AdjustTop_Current = .AdjustTop_Current - (lngAdjustByAmount * 3&)
6240      MICRChanged True  ' ** Procedure: Below.
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

Private Sub cmdMoveUp_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveUp_KeyDown"

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
6490          .cmdMoveLeft.SetFocus
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab
6550        With Me
6560          intRetVal = 0
6570          .cmdClose.SetFocus
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

Private Sub cmdMoveUp_GotFocus()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveUp_GotFocus"

6710    With Me
6720      .cmdMoveUp_dots.Visible = True
6730      .cmdMoveLeft_dots.Visible = False
6740      .cmdMoveDown_dots.Visible = False
6750      .cmdMoveRight_dots.Visible = False
6760    End With

EXITP:
6770    Exit Sub

ERRH:
6780    Select Case ERR.Number
        Case Else
6790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6800    End Select
6810    Resume EXITP

End Sub

Private Sub cmdMoveUp_LostFocus()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveUp_LostFocus"

6910    With Me
6920      .cmdMoveUp_dots.Visible = False
6930    End With

EXITP:
6940    Exit Sub

ERRH:
6950    Select Case ERR.Number
        Case Else
6960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6970    End Select
6980    Resume EXITP

End Sub

Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveUp_MouseDown"

7010    With Me
7020      .cmdMoveUp_sunken_focus_dots_img.Visible = True
7030      .cmdMoveUp_raised_img.Visible = False
7040    End With

EXITP:
7050    Exit Sub

ERRH:
7060    Select Case ERR.Number
        Case Else
7070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7080    End Select
7090    Resume EXITP

End Sub

Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveUp_MouseUp"

7110    With Me
7120      .cmdMoveUp_raised_img.Visible = True
7130      .cmdMoveUp_sunken_focus_dots_img.Visible = False
7140    End With

EXITP:
7150    Exit Sub

ERRH:
7160    Select Case ERR.Number
        Case Else
7170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7180    End Select
7190    Resume EXITP

End Sub

Private Sub cmdMoveLeft_Click()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveLeft_Click"

7210    With Me
7220      .Moveable_lbl.Left = .Moveable_lbl.Left - lngAdjustByAmount
7230      .AdjustLeft_Current = .AdjustLeft_Current - lngAdjustByAmount
7240      MICRChanged True  ' ** Procedure: Below.
7250    End With

EXITP:
7260    Exit Sub

ERRH:
7270    Select Case ERR.Number
        Case Else
7280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7290    End Select
7300    Resume EXITP

End Sub

Private Sub cmdMoveLeft_DblClick(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveLeft_DblClick"

7410    With Me
7420      .Moveable_lbl.Left = .Moveable_lbl.Left - (lngAdjustByAmount * 3&)
7430      .AdjustLeft_Current = .AdjustLeft_Current - (lngAdjustByAmount * 3&)
7440      MICRChanged True  ' ** Procedure: Below.
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

Private Sub cmdMoveLeft_GotFocus()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveLeft_GotFocus"

7610    With Me
7620      .cmdMoveUp_dots.Visible = False
7630      .cmdMoveLeft_dots.Visible = True
7640      .cmdMoveDown_dots.Visible = False
7650      .cmdMoveRight_dots.Visible = False
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

Private Sub cmdMoveLeft_LostFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveLeft_LostFocus"

7810    With Me
7820      .cmdMoveLeft_dots.Visible = False
7830    End With

EXITP:
7840    Exit Sub

ERRH:
7850    Select Case ERR.Number
        Case Else
7860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7870    End Select
7880    Resume EXITP

End Sub

Private Sub cmdMoveLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveLeft_MouseDown"

7910    With Me
7920      .cmdMoveDown_sunken_focus_dots_img.Visible = True
7930      .cmdMoveLeft_raised_img.Visible = False
7940    End With

EXITP:
7950    Exit Sub

ERRH:
7960    Select Case ERR.Number
        Case Else
7970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7980    End Select
7990    Resume EXITP

End Sub

Private Sub cmdMoveLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveLeft_MouseUp"

8010    With Me
8020      .cmdMoveLeft_raised_img.Visible = True
8030      .cmdMoveLeft_sunken_focus_dots_img.Visible = False
8040    End With

EXITP:
8050    Exit Sub

ERRH:
8060    Select Case ERR.Number
        Case Else
8070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8080    End Select
8090    Resume EXITP

End Sub

Private Sub cmdMoveDown_Click()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveDown_Click"

8110    With Me
8120      .Moveable_lbl.Top = .Moveable_lbl.Top + lngAdjustByAmount
8130      .AdjustTop_Current = .AdjustTop_Current + lngAdjustByAmount
8140      MICRChanged True  ' ** Procedure: Below.
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

Private Sub cmdMoveDown_DblClick(Cancel As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveDown_DblClick"

8310    With Me
8320      .Moveable_lbl.Top = .Moveable_lbl.Top + (lngAdjustByAmount * 3&)
8330      .AdjustTop_Current = .AdjustTop_Current + (lngAdjustByAmount * 3&)
8340      MICRChanged True  ' ** Procedure: Below.
8350    End With

EXITP:
8360    Exit Sub

ERRH:
8370    Select Case ERR.Number
        Case Else
8380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8390    End Select
8400    Resume EXITP

End Sub

Private Sub cmdMoveDown_GotFocus()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveDown_GotFocus"

8510    With Me
8520      .cmdMoveUp_dots.Visible = False
8530      .cmdMoveLeft_dots.Visible = False
8540      .cmdMoveDown_dots.Visible = True
8550      .cmdMoveRight_dots.Visible = False
8560    End With

EXITP:
8570    Exit Sub

ERRH:
8580    Select Case ERR.Number
        Case Else
8590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8600    End Select
8610    Resume EXITP

End Sub

Private Sub cmdMoveDown_LostFocus()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveDown_LostFocus"

8710    With Me
8720      .cmdMoveDown_dots.Visible = False
8730    End With

EXITP:
8740    Exit Sub

ERRH:
8750    Select Case ERR.Number
        Case Else
8760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8770    End Select
8780    Resume EXITP

End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveDown_MouseDown"

8810    With Me
8820      .cmdMoveDown_sunken_focus_dots_img.Visible = True
8830      .cmdMoveDown_raised_img.Visible = False
8840    End With

EXITP:
8850    Exit Sub

ERRH:
8860    Select Case ERR.Number
        Case Else
8870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8880    End Select
8890    Resume EXITP

End Sub

Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveDown_MouseUp"

8910    With Me
8920      .cmdMoveDown_raised_img.Visible = True
8930      .cmdMoveDown_sunken_focus_dots_img.Visible = False
8940    End With

EXITP:
8950    Exit Sub

ERRH:
8960    Select Case ERR.Number
        Case Else
8970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8980    End Select
8990    Resume EXITP

End Sub

Private Sub cmdMoveRight_Click()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveRight_Click"

9010    With Me
9020      .Moveable_lbl.Left = .Moveable_lbl.Left + lngAdjustByAmount
9030      .AdjustLeft_Current = .AdjustLeft_Current + lngAdjustByAmount
9040      MICRChanged True  ' ** Procedure: Below.
9050    End With

EXITP:
9060    Exit Sub

ERRH:
9070    Select Case ERR.Number
        Case Else
9080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9090    End Select
9100    Resume EXITP

End Sub

Private Sub cmdMoveRight_DblClick(Cancel As Integer)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveRight_DblClick"

9210    With Me
9220      .Moveable_lbl.Left = .Moveable_lbl.Left + (lngAdjustByAmount * 3&)
9230      .AdjustLeft_Current = .AdjustLeft_Current + (lngAdjustByAmount * 3&)
9240      MICRChanged True  ' ** Procedure: Below.
9250    End With

EXITP:
9260    Exit Sub

ERRH:
9270    Select Case ERR.Number
        Case Else
9280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9290    End Select
9300    Resume EXITP

End Sub

Private Sub cmdMoveRight_KeyDown(KeyCode As Integer, Shift As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveRight_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9420    intShiftDown = (Shift And acShiftMask) > 0
9430    intAltDown = (Shift And acAltMask) > 0
9440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9460      Select Case intRetVal
          Case vbKeyTab
9470        With Me
9480          intRetVal = 0
9490          If .cmdPrintChecksTest.Enabled = True Then
9500            .cmdPrintChecksTest.SetFocus
9510          ElseIf .cmdSaveAdjustments.Enabled = True Then
9520            .cmdSaveAdjustments.SetFocus
9530          ElseIf .cmdStartOver.Enabled = True Then
9540            .cmdStartOver.SetFocus
9550          Else
9560            .cmdClose.SetFocus
9570          End If
9580        End With
9590      End Select
9600    End If

        ' ** Shift keys.
9610    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9620      Select Case intRetVal
          Case vbKeyTab
9630        With Me
9640          intRetVal = 0
9650          .cmdMoveDown.SetFocus
9660        End With
9670      End Select
9680    End If

EXITP:
9690    KeyCode = intRetVal
9700    Exit Sub

ERRH:
9710    intRetVal = 0
9720    Select Case ERR.Number
        Case Else
9730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9740    End Select
9750    Resume EXITP

End Sub

Private Sub cmdMoveRight_GotFocus()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveRight_GotFocus"

9810    With Me
9820      .cmdMoveUp_dots.Visible = False
9830      .cmdMoveLeft_dots.Visible = False
9840      .cmdMoveDown_dots.Visible = False
9850      .cmdMoveRight_dots.Visible = True
9860    End With

EXITP:
9870    Exit Sub

ERRH:
9880    Select Case ERR.Number
        Case Else
9890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9900    End Select
9910    Resume EXITP

End Sub

Private Sub cmdMoveRight_LostFocus()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveRight_LostFocus"

10010   With Me
10020     .cmdMoveRight_dots.Visible = False
10030   End With

EXITP:
10040   Exit Sub

ERRH:
10050   Select Case ERR.Number
        Case Else
10060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10070   End Select
10080   Resume EXITP

End Sub

Private Sub cmdMoveRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveRight_MouseDown"

10110   With Me
10120     .cmdMoveRight_sunken_focus_dots_img.Visible = True
10130     .cmdMoveRight_raised_img.Visible = False
10140   End With

EXITP:
10150   Exit Sub

ERRH:
10160   Select Case ERR.Number
        Case Else
10170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10180   End Select
10190   Resume EXITP

End Sub

Private Sub cmdMoveRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdMoveRight_MouseUp"

10210   With Me
10220     .cmdMoveRight_raised_img.Visible = True
10230     .cmdMoveRight_sunken_focus_dots_img.Visible = False
10240   End With

EXITP:
10250   Exit Sub

ERRH:
10260   Select Case ERR.Number
        Case Else
10270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10280   End Select
10290   Resume EXITP

End Sub

Private Sub MICRChanged(blnTmp01 As Boolean)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "MICRChanged"

10310   With Me
10320     Select Case blnTmp01
          Case True
10330       If .Moveable_lbl.Top = lngStart_Top And .Moveable_lbl.Left = lngStart_Left Then
              ' ** If they've moved back to the starting position, reset for no changes.
10340         blnChanged = False
10350         .cmdSaveAdjustments.Enabled = False
10360         .cmdStartOver.Enabled = False
10370       Else
10380         blnChanged = True
10390         .cmdSaveAdjustments.Enabled = True
10400         .cmdStartOver.Enabled = True
10410       End If
10420     Case False
10430       blnChanged = False
10440       If .cmdPrintChecksTest.Enabled = True Then
10450         .cmdPrintChecksTest.SetFocus
10460       Else
10470         .cmdClose.SetFocus
10480       End If
10490       .cmdSaveAdjustments.Enabled = False
10500       .cmdStartOver.Enabled = False
10510     End Select
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
