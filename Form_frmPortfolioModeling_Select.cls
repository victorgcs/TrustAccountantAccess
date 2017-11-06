VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPortfolioModeling_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmPortfolioModeling_Select"

'VGC 10/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** For GTR, shows arrows, then continues.

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Cancel:           C {cmdCancel}
' **   Name:             M {opgAccountSource_optName}
' **   OK:               O {cmdOK}
' **   Exit:             X {cmdCancel}

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long

Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long, blnFormOpening As Boolean
Private blnJustClose As Boolean
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long ', lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim lngDetail_HeightNew As Long, lngDetail_HeightDiff As Long
        Dim lngTmp01 As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True: blnFormOpening = False
150           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

160           strCallingForm = Trim(.OpenArgs)

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

190           blnJustClose = False
200           blnGTR_Emblem = False: lngGTR_Stat = 0&
210           gstrAccountNo = vbNullString

              ' ** Variables are fed empty, then populated ByRef.
220           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

230           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

240           If gblnGoToReport = True Then
250             If .GoToReport_arw_acct_img.Visible = False Then

260               lngDetail_HeightNew = ((.cmbAccounts.Top + .cmbAccounts.Height) + .GoToReport_arw_acct_img.Height)
270               lngDetail_HeightDiff = (lngDetail_HeightNew - .Detail.Height)
280               .Detail.Height = lngDetail_HeightNew
290               .GoToReport_arw_acct_img.Top = (.cmbAccounts.Top + .cmbAccounts.Height)
300               .GoToReport_arw_acct_img.Left = (.cmbAccounts.Left + (5& * lngTpp))
310               .GoToReport_arw_acct_img.Visible = True
320               .GoToReport_Emblem_01_img.Visible = True

330               lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
340               lngMonitorNum = 1&: lngTmp01 = 0&
350               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
360               If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

370               If lngMonitorNum = 1& Then lngTmp01 = lngFrm_Top
380               DoCmd.MoveSize lngFrm_Left, lngTmp01, lngFrm_Width, lngFrm_Height + lngDetail_HeightDiff  'lngFrm_Top
390               If lngMonitorNum > 1& Then
400                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
410               End If
420               DoEvents
430             End If
440           End If

              ' ** EVENT CHECK: chkRememberMe!
              ' ** NONE!

450           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
460           chkRememberMe_AfterUpdate  ' ** Procedure: Below.

470           Select Case IsNull(.cmbAccounts)
              Case True
480             .cmbAccounts.SetFocus
490           Case False
500             .cmdOK.SetFocus
510           End Select

520           DoCmd.Hourglass False

530           blnIsOpen = False

540         Else
550           Cancel = -1
560         End If
570       Else
580         Cancel = -1
590       End If
600     End With

610     If Cancel = -1 Then
620       Beep
630       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
640     End If

EXITP:
650     Exit Sub

ERRH:
660     DoCmd.Hourglass False
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

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
880         cmdCancel_Click  ' ** Procedure: below.
890       End Select
900     End If

        ' ** Alt keys.
910     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
920       Select Case intRetVal
          Case vbKeyX
930         intRetVal = 0
940         cmdCancel_Click  ' ** Procedure: below.
950       End Select
960     End If

EXITP:
970     KeyCode = intRetVal
980     Exit Sub

ERRH:
990     intRetVal = 0
1000    Select Case ERR.Number
        Case Else
1010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1020    End Select
1030    Resume EXITP

End Sub

Private Sub Form_Timer()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim lngTmp03 As Long

        Static lngTmp01 As Long, lngTmp02 As Long

1110    With Me
1120      .TimerInterval = 0&
1130      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
1140        lngGTR_Stat = lngGTR_Stat + 1&
1150        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1160        lngTmp01 = lngTmp01 + 1&
1170        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1180        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1190        If lngTmp01 > 1& Then
1200          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1210        End If
1220        If lngTmp01 = 24& Then
1230          lngTmp01 = 0&
1240          lngTmp02 = lngTmp02 + 1&
1250          If lngTmp02 < 5& Then
1260            .TimerInterval = 50&
1270          Else
1280            blnGTR_Emblem = False
1290          End If
1300        Else
1310          .TimerInterval = 50&
1320        End If
1330      End If  ' ** blnGTR_Emblem.
1340      If gblnSetFocus = True Then
1350        gblnSetFocus = False
1360        lngTmp03 = 0
1370        strFormName = vbNullString
1380  On Error Resume Next
1390        Do Until strFormName = THIS_NAME
1400          lngTmp03 = lngTmp03 + 1&
1410          DoCmd.SelectObject acForm, THIS_NAME, False
1420          strFormName = Screen.ActiveForm.Name
1430          If lngTmp03 > 100& Then Exit Do
1440        Loop
1450  On Error GoTo ERRH
1460      End If  ' ** gblnSetFocus.
1470      If gblnGoToReport = True Then
1480        DoCmd.Hourglass True  ' ** Make sure it's still running.
1490        If ((lngGTR_Stat > 0& And lngGTR_Stat < 36&) Or (lngGTR_Stat > 36& And lngGTR_Stat < 72&) Or (lngGTR_Stat > 72&)) Then
              ' ** Just let the emblem go.
1500        Else
1510          DoCmd.Hourglass True  ' ** Make sure it's still running.
1520          If blnGTR_Emblem = False Then
1530            blnGTR_Emblem = True
1540            lngTmp01 = 0&: lngTmp02 = 0&
1550            lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
1560            .GoToReport_arw_ok_img.Left = (.GoToReport_arw_ok_img.Left - lngTmp03)
1570            .cmdOK.Left = (.cmdOK.Left - lngTmp03)
1580            .cmdCancel.Left = (.cmdCancel.Left - lngTmp03)
1590          End If
1600          DoEvents
1610          Select Case blnGoingToReport
              Case True
1620            Select Case .GoToReport_arw_acct_img.Visible
                Case True
1630              .GoToReport_arw_acct_img.Visible = False
1640              .GoToReport_arw_ok_img.Visible = True
1650              .cmdOK.SetFocus
1660              DoEvents
1670              .TimerInterval = 50&
1680            Case False
1690              blnGoingToReport = False
1700              .TimerInterval = 50&
1710              cmdOK_Click  ' ** Procedure: Below
1720            End Select
1730          Case False
1740            blnGoingToReport = True
1750            If IsNull(.cmbAccounts) = True Then
1760              .cmbAccounts = .cmbAccounts.Column(0, 0)
1770              DoEvents
1780            End If
1790            .cmbAccounts.SetFocus
1800            .cmbAccounts.SelLength = 0
1810            .cmbAccounts.SelStart = 99
1820            DoEvents
1830            .TimerInterval = 50&
1840          End Select
1850        End If
1860      End If
1870    End With

EXITP:
1880    Exit Sub

ERRH:
1890    Select Case ERR.Number
        Case Else
1900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1910    End Select
1920    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2010    With Me
2020      If .chkRememberMe = False Then
2030        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
2040      End If
2050    End With

2060    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2070    If blnJustClose = False Then
2080      gstrReturningForm = THIS_NAME
2090      Select Case strCallingForm
          Case "frmMenu_Asset"
2100        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2110      End Select
2120    End If

EXITP:
2130    Exit Sub

ERRH:
2140    Select Case ERR.Number
        Case Else
2150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2160    End Select
2170    Resume EXITP

End Sub

Private Sub cmdOK_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngInvObjID As Long, lngCnt As Long
        Dim strDocName As String, strInvObj As String

2210    With Me

2220      DoCmd.Hourglass True
2230      DoEvents

2240      .TimerInterval = 0&

2250      If IsNull(.cmbAccounts) = True Then
2260        MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
2270        blnGoingToReport = False
2280        gblnGoToReport = False
2290      Else
2300        If .cmbAccounts = vbNullString Then
2310          MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
2320          blnGoingToReport = False
2330          gblnGoToReport = False
2340        Else

2350          gstrAccountNo = .cmbAccounts

2360          Set dbs = CurrentDb
2370          With dbs

                ' ** Get current number of AssetTypes.
2380            lngCnt = DCount("[assettype]", "AssetType")

                ' ** Empty tmpPortfolioModel.
2390            Set qdf = .QueryDefs("qryPortfolioModeling_03_01")
2400            qdf.Execute
2410            Set qdf = Nothing

                ' ** Empty tblPortfolioModeling.
2420            Set qdf = .QueryDefs("qryPortfolioModeling_03_02")
2430            qdf.Execute
2440            Set qdf = Nothing

                ' ** Empty tblPortfolioModeling2.
2450            Set qdf = .QueryDefs("qryPortfolioModeling_03_03")
2460            qdf.Execute
2470            Set qdf = Nothing

                ' ** Account, linked to InvestmentObjective, with invobj_id, by specified [actno].
2480            Set qdf = .QueryDefs("qryPortfolioModeling_11")
2490            With qdf.Parameters
2500              ![actno] = gstrAccountNo
2510            End With
2520            Set rst = qdf.OpenRecordset
2530            With rst  ' ** This shouldn't ever by empty.
2540              .MoveFirst
2550              lngInvObjID = ![invobj_id]
2560              strInvObj = ![investmentobj]
2570              .Close
2580            End With
2590            Set rst = Nothing
2600            Set qdf = Nothing

                ' ** PortfolioModel, by specified [actno], [inobid].
2610            Set qdf = .QueryDefs("qryPortfolioModeling_12")
2620            With qdf.Parameters
2630              ![actno] = gstrAccountNo
2640              ![inobid] = lngInvObjID
2650            End With
2660            Set rst = qdf.OpenRecordset
2670            If rst.BOF = True And rst.EOF = True Then
2680              Set rst = Nothing
2690              Set qdf = Nothing
                  ' ** Append full set of AssetType records to PortfolioModel, by specified [actno], [inobid].
2700              Set qdf = .QueryDefs("qryPortfolioModeling_05")
2710              With qdf.Parameters
2720                ![actno] = gstrAccountNo
2730                ![inobid] = lngInvObjID
2740              End With
2750              qdf.Execute
2760              Set qdf = Nothing
2770            Else
2780              rst.MoveFirst
2790              If rst![cnt_atyp] > lngCnt Then
                    ' ** Should we even check if there are too many?
2800              ElseIf rst![cnt_atyp] < lngCnt Then
2810                rst.Close
2820                Set rst = Nothing
2830                Set qdf = Nothing
                    ' ** Append qryPortfolioModeling_13_01 (AssetType, by specified [actno], [inobid]) to tmpPortfolioModel.
2840                Set qdf = .QueryDefs("qryPortfolioModeling_13_02")
2850                With qdf.Parameters
2860                  ![actno] = gstrAccountNo
2870                  ![inobid] = lngInvObjID
2880                End With
2890                qdf.Execute
2900                Set qdf = Nothing
                    ' ** Append qryPortfolioModeling_13_03 (tmpPortfolioModel, not in PortfolioModel) to PortfolioModel.
2910                Set qdf = .QueryDefs("qryPortfolioModeling_13_04")
2920                qdf.Execute
2930                Set qdf = Nothing
2940              Else
2950                rst.Close
2960                Set rst = Nothing
2970                Set qdf = Nothing
2980              End If
2990            End If

                ' ** Append qryPortfolioModAssetTotals_06 (qryPortfolioModAssetTotals_05 (PortfolioModel, linked to
                ' ** qryPortfolioModAssetTotals_04 (qryPortfolioModAssetTotals_03 (qryPortfolioModAssetTotals_02
                ' ** (qryPortfolioModAssetTotals_01 (Account, linked to ActiveAssets, with add'l fields, by specified
                ' ** [actno]), grouped and summed, by assetno, accountno), grouped and summed, by accountno, with
                ' ** TotalMarketValue), linked to qryPortfolioModAssetTotals_02 (qryPortfolioModAssetTotals_01
                ' ** (Account, linked to ActiveAssets, with add'l fields, by specified [actno]), grouped and summed,
                ' ** by assetno, accountno), with Current, TotalCurrent, CASH), Account, with Difference, Current,
                ' ** by specified [inobid]), grouped) to tblPortfolioModeling.
3000            Set qdf = .QueryDefs("qryPortfolioModeling_06")
3010            With qdf.Parameters
3020              ![actno] = gstrAccountNo
3030              ![inobid] = lngInvObjID
3040            End With
3050            qdf.Execute
3060            Set qdf = Nothing

                ' ** Append tblPortfolioModeling, grouped and summed, to tblPortfolioModeling2, by specified [inobid].
3070            Set qdf = .QueryDefs("qryPortfolioModeling_07")
3080            With qdf.Parameters
3090              ![inobid] = lngInvObjID
3100            End With
3110            qdf.Execute
3120            Set qdf = Nothing

                ' ** Zero-out Current in tblPortfolioModeling.  'AND WHY ARE WE DOING THIS?
3130            Set qdf = .QueryDefs("qryPortfolioModeling_08")
3140            qdf.Execute
3150            Set qdf = Nothing

3160            .Close
3170          End With  ' ** dbs.
3180          Set dbs = Nothing

3190          gblnSetFocus = True
3200          strDocName = "frmPortfolioModeling"
3210          gstrReturningForm = strDocName
3220          DoCmd.OpenForm strDocName, , , , , , THIS_NAME

3230          If gblnGoToReport = True Then
3240            Forms(strDocName).TimerInterval = 50&
3250            blnGoingToReport = False
3260            blnGTR_Emblem = False
3270          End If

3280          blnJustClose = True
3290          DoCmd.Close acForm, THIS_NAME

3300        End If
3310      End If

3320    End With

EXITP:
3330    Set rst = Nothing
3340    Set qdf = Nothing
3350    Set dbs = Nothing
3360    Exit Sub

ERRH:
3370    DoCmd.Hourglass False
3380    Select Case ERR.Number
        Case 6
          ' ** Overflow.
3390      MsgBox "Please check market values to continue!", vbInformation + vbOKOnly, "Error"
3400    Case Else
3410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3420    End Select
3430    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3520    intShiftDown = (Shift And acShiftMask) > 0
3530    intAltDown = (Shift And acAltMask) > 0
3540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3560      Select Case intRetVal
          Case vbKeyTab
3570        With Me
3580          intRetVal = 0
3590          .cmdCancel.SetFocus
3600        End With
3610      End Select
3620    End If

        ' ** Shift keys.
3630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3640      Select Case intRetVal
          Case vbKeyTab
3650        With Me
3660          intRetVal = 0
3670          .cmbAccounts.SetFocus
3680        End With
3690      End Select
3700    End If

EXITP:
3710    KeyCode = intRetVal
3720    Exit Sub

ERRH:
3730    intRetVal = 0
3740    Select Case ERR.Number
        Case Else
3750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3760    End Select
3770    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

3810    DoCmd.Close acForm, THIS_NAME

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3920    intShiftDown = (Shift And acShiftMask) > 0
3930    intAltDown = (Shift And acAltMask) > 0
3940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3960      Select Case intRetVal
          Case vbKeyTab
3970        With Me
3980          intRetVal = 0
3990          .cmbAccounts.SetFocus
4000        End With
4010      End Select
4020    End If

        ' ** Shift keys.
4030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4040      Select Case intRetVal
          Case vbKeyTab
4050        With Me
4060          intRetVal = 0
4070          .cmdOK.SetFocus
4080        End With
4090      End Select
4100    End If

        ' ** Ctrl keys.
4110    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4120      Select Case intRetVal
          Case vbKeyTab
4130        With Me
4140          intRetVal = 0
4150          .opgAccountSource.SetFocus
4160        End With
4170      End Select
4180    End If

EXITP:
4190    KeyCode = intRetVal
4200    Exit Sub

ERRH:
4210    intRetVal = 0
4220    Select Case ERR.Number
        Case Else
4230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4240    End Select
4250    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

4310    strAccountNo = vbNullString

4320    With Me
4330      If IsNull(.cmbAccounts) = False Then
4340        If Len(.cmbAccounts.Column(0)) > 0 Then
4350          strAccountNo = .cmbAccounts.Column(0)
4360        End If
4370      End If
4380      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
4390        .cmbAccounts.RowSource = "qryPortfolioModeling_Select_08"
4400        .opgAccountSource_optNumber_lbl.FontBold = True
4410        .opgAccountSource_optName_lbl.FontBold = False
4420      Case .opgAccountSource_optName.OptionValue
4430        .cmbAccounts.RowSource = "qryPortfolioModeling_Select_09"
4440        .opgAccountSource_optNumber_lbl.FontBold = False
4450        .opgAccountSource_optName_lbl.FontBold = True
4460      End Select
4470      DoEvents
4480      If strAccountNo <> vbNullString Then
4490        .cmbAccounts = strAccountNo
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

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4620    intShiftDown = (Shift And acShiftMask) > 0
4630    intAltDown = (Shift And acAltMask) > 0
4640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4670        With Me
4680          intRetVal = 0
4690          .chkRememberMe.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4750        With Me
4760          intRetVal = 0
4770          .cmdCancel.SetFocus
4780        End With
4790      End Select
4800    End If

EXITP:
4810    KeyCode = intRetVal
4820    Exit Sub

ERRH:
4830    intRetVal = 0
4840    Select Case ERR.Number
        Case Else
4850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860    End Select
4870    Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

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
4990          .chkRememberMe.SetFocus
5000        End With
5010      End Select
5020    End If

        ' ** Shift keys.
5030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5050        With Me
5060          intRetVal = 0
5070          .cmdCancel.SetFocus
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

Private Sub chkRememberMe_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

5210    With Me
5220      Select Case .chkRememberMe
          Case True
5230        .chkRememberMe_lbl.FontBold = True
5240      Case False
5250        .chkRememberMe_lbl.FontBold = False
5260      End Select
5270    End With

EXITP:
5280    Exit Sub

ERRH:
5290    Select Case ERR.Number
        Case Else
5300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5310    End Select
5320    Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
5470        With Me
5480          intRetVal = 0
5490          .cmbAccounts.SetFocus
5500        End With
5510      End Select
5520    End If

        ' ** Shift keys.
5530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5550        With Me
5560          intRetVal = 0
5570          .opgAccountSource.SetFocus
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

Private Sub cmbAccounts_GotFocus()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_GotFocus"

5710    With Me
5720      .cmbAccounts.SelLength = 0
5730      .cmbAccounts.SelStart = 99
5740      DoEvents
5750    End With

EXITP:
5760    Exit Sub

ERRH:
5770    Select Case ERR.Number
        Case Else
5780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5790    End Select
5800    Resume EXITP

End Sub

Private Sub cmbAccounts_NotInList(NewData As String, Response As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_NotInList"

5910    With Me
5920      MsgBox "The account you entered does not exist." & vbCrLf & vbCrLf & _
            "First create the account, then proceed to setup its model.", vbInformation + vbOKOnly, "Account Not Found"
5930      Response = acDataErrContinue
5940      .cmbAccounts.text = vbNullString
5950    End With

EXITP:
5960    Exit Sub

ERRH:
5970    Select Case ERR.Number
        Case Else
5980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5990    End Select
6000    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6120    intShiftDown = (Shift And acShiftMask) > 0
6130    intAltDown = (Shift And acAltMask) > 0
6140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6170        With Me
6180          intRetVal = 0
6190          .cmdOK.SetFocus
6200        End With
6210      End Select
6220    End If

        ' ** Shift keys.
6230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6250        With Me
6260          intRetVal = 0
6270          .cmdCancel.SetFocus
6280        End With
6290      End Select
6300    End If

        ' ** Ctrl-Shift keys.
6310    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6320      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6330        With Me
6340          intRetVal = 0
6350          .chkRememberMe.SetFocus
6360        End With
6370      End Select
6380    End If

EXITP:
6390    KeyCode = intRetVal
6400    Exit Sub

ERRH:
6410    intRetVal = 0
6420    Select Case ERR.Number
        Case Else
6430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6440    End Select
6450    Resume EXITP

End Sub

Public Sub GTREmblem_Off()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

6510    With Me
6520      blnGTR_Emblem = False
6530      For lngX = 1& To 24&
6540        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
6550      Next
6560    End With

EXITP:
6570    Exit Sub

ERRH:
6580    Select Case ERR.Number
        Case Else
6590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6600    End Select
6610    Resume EXITP

End Sub
