VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenu_Asset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMenu_Asset"

'VGC 10/11/2017: CHANGES!

' ** For GTR, shows arrows, sets focus, then continues.

' ** Shortcut Alt keys responsive from this form:
' **   Add/Edit:         A {cmdAddEditAsset}
' **   Close:            C {cmdClose}
' **   Load Timer:       D {chkLoadTimer}
' **   Holdings:         H {cmdHoldings}
' **   Portfolio Model:  M {cmdPortfolioModel}
' **   Pricing:          P {cmdAssetPricing}
' **   History:          S {cmdHistory}
' **   Locations:        T {cmdLocations}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}
' **   Timer Chk Vis:    T {Form_KeyDown}
' **   Close Form:       X {cmdCloseForm}

Private lngTpp As Long, blnDontEmpty As Boolean
Private blnJustClose As Boolean, blnFormOpening As Boolean
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           gblnSignal = False: blnDontEmpty = False

160           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

170           gstrCrtRpt_Version = vbNullString  ' ** Borrowing this variable from Court Reports.

180           blnJustClose = False: blnFormOpening = False
190           blnGTR_Emblem = False: lngGTR_Stat = 0&

              'lngTpp = GetTPP  ' ** Module Function: modWindowsFunctions.
200           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

210           Select Case .chkLoadTimer_vis
              Case True
220             .chkLoadTimer.Visible = True
230             .chkLoadTimer.Enabled = False
                '.chkLoadTimer_lbl1.Visible = True
240           Case False
250             .chkLoadTimer.Visible = False
                '.chkLoadTimer_lbl1.Visible = False
260           End Select
270           chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
280           DoEvents

290           DoCmd.Hourglass False

300         Else
310           Cancel = -1
320         End If
330       Else
340         Cancel = -1
350       End If
360     End With

370     If Cancel = -1 Then
380       Beep
390       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
400     End If

EXITP:
410     Exit Sub

ERRH:
420     DoCmd.Hourglass False
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer, Cancel As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Load Timer:       D {chkLoadTimer}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}
        ' **   Timer Chk Vis:    T {Form_KeyDown}
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         intRetVal = 0
580         cmdClose_Click  ' ** Procedure: Below.
590       End Select
600     End If

        ' ** Alt keys.
610     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
620       Select Case intRetVal
          Case vbKeyD
630         With Me
640           intRetVal = 0
650           If .chkLoadTimer.Visible = True Then
660             If .chkLoadTimer.Enabled = True Then
670               .chkLoadTimer.SetFocus
680             End If
690             .chkLoadTimer = (Not .chkLoadTimer)
700             chkLoadTimer_AfterUpdate
710           End If
720         End With
730       Case vbKeyX
740         intRetVal = 0
750         cmdClose_Click  ' ** Procedure: Below.
760       End Select
770     End If

        ' ** Ctrl-Shift keys.
780     If intCtrlDown And (Not intAltDown) And intShiftDown Then
790       Select Case intRetVal
          Case vbKeyF
800         With Me
810           intRetVal = 0
820           .FocusHolder.SetFocus
830         End With
840       Case vbKeyT
850         With Me
860           intRetVal = 0
870           If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
880             FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
890           ElseIf CurrentUser = "TAAdmin" Or CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
900             Select Case .chkLoadTimer.Visible
                Case True
910               .FocusHolder.SetFocus
920               DoEvents
930               .chkLoadTimer.Visible = False
                  '.chkLoadTimer_lbl1.Visible = False
940               gblnTimer = False
950             Case False
960               .chkLoadTimer.Visible = True
970               .chkLoadTimer.Enabled = False
                  '.chkLoadTimer_lbl1.Visible = True
980               chkLoadTimer_AfterUpdate  ' ** Procedure: Below.
990             End Select
1000          End If
1010          .chkLoadTimer_vis = .chkLoadTimer.Visible
1020        End With
1030      Case vbKeyX
1040        If GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
1050          intRetVal = 0
1060          blnJustClose = True
1070          SetOption_Dev  ' ** Module Function: modUtilities.
1080          gblnSignal = True  ' ** Borrowing this variable.
1090          DoCmd.Close acForm, THIS_NAME
1100          DoCmd.Hourglass False
1110        End If
1120      End Select
1130    End If

EXITP:
1140    KeyCode = intRetVal
1150    Exit Sub

ERRH:
1160    intRetVal = 0
1170    Select Case ERR.Number
        Case Else
1180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1190    End Select
1200    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim strFormName As String

        Static lngTmp01 As Long, lngTmp02 As Long

1310    With Me
1320      .TimerInterval = 0&
1330      If blnGTR_Emblem = True Then
1340        lngGTR_Stat = lngGTR_Stat + 1&
1350        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
1360        lngTmp01 = lngTmp01 + 1&
1370        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
1380        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
1390        If lngTmp01 > 1& Then
1400          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
1410        End If
1420        If lngTmp01 = 24& Then
1430          lngTmp01 = 0&
1440          lngTmp02 = lngTmp02 + 1&
1450          If lngTmp02 < 5& Then
1460            .TimerInterval = 50&
1470          Else
1480            blnGTR_Emblem = False
1490          End If
1500        Else
1510          .TimerInterval = 50&
1520        End If
1530      End If  ' ** blnGTR_Emblem.
1540      If gblnSetFocus = True Then
1550        gblnSetFocus = False
1560        strFormName = vbNullString
1570  On Error Resume Next
1580        Do Until strFormName = THIS_NAME
1590          DoCmd.SelectObject acForm, THIS_NAME, False
1600          strFormName = Screen.ActiveForm.Name
1610        Loop
1620  On Error GoTo ERRH
1630        DoEvents
1640      End If
1650      If blnGoingToReport = False Then
1660        Select Case gstrReturningForm
            Case "frmAssets"
1670          .cmdAddEditAsset.SetFocus
1680        Case "frmAssetPricing"
1690          .cmdAssetPricing.SetFocus
1700        Case "frmRpt_AssetHistory"
1710          .cmdHistory.SetFocus
1720        Case "frmRpt_Holdings"
1730          .cmdHoldings.SetFocus
1740        Case "frmRpt_Locations"
1750          .cmdLocations.SetFocus
1760        Case "frmPortfolioModeling_Select", "frmPortfolioModeling"
1770          .cmdPortfolioModel.SetFocus
1780        Case Else
1790          .cmdAddEditAsset.SetFocus
1800        End Select
1810        gstrReturningForm = vbNullString
1820        DoCmd.Hourglass False
1830      End If
1840      If gblnGoToReport = True Then
1850        DoCmd.Hourglass True  ' ** Make sure it's still running.
1860        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48&)) Then
              ' ** Just let the emblem go.
1870        Else
1880          If blnGTR_Emblem = False Then
1890            blnGTR_Emblem = True
1900            lngTmp01 = 0&: lngTmp02 = 0&
1910          End If
1920          DoEvents
1930          Select Case blnGoingToReport
              Case True
1940            Select Case garr_varGoToReport(GTR_INV1)
                Case True
1950              Select Case garr_varGoToReport(GTR_CTL1)
                  Case "cmdAddEditAsset"
1960                .TimerInterval = 50&
1970                cmdAddEditAsset_Click  ' ** Procedure: Below.
1980              Case "cmdAssetPricing"
1990                .TimerInterval = 50&
2000                cmdAssetPricing_Click  ' ** Procedure: Below.
2010              Case "cmdHistory"
2020                .TimerInterval = 50&
2030                cmdHistory_Click  ' ** Procedure: Below.
2040              Case "cmdHoldings"
2050                .TimerInterval = 50&
2060                cmdHoldings_Click  ' ** Procedure: Below.
2070              Case "cmdLocations"
2080                .TimerInterval = 50&
2090                cmdLocations_Click  ' ** Procedure: Below.
2100              Case "cmdPortfolioModel"
2110                .TimerInterval = 50&
2120                cmdPortfolioModel_Click  ' ** Procedure: Below.
2130              Case Else
                    ' ** Shouldn't happen.
2140                blnGoingToReport = False
2150                DoCmd.Hourglass False
2160                gblnGoToReport = False
2170                blnGTR_Emblem = False
2180                GTREmblem_Off  ' ** Procedure: Below.
2190                Beep
2200              End Select
2210            Case False
                  ' ** Just sit there, and we're done.
2220              blnGoingToReport = False
2230              DoCmd.Hourglass False
2240              gblnGoToReport = False
2250              blnGTR_Emblem = False
2260              GTREmblem_Off  ' ** Procedure: Below.
2270              Beep
2280            End Select
2290          Case False
2300            If garr_varGoToReport(GTR_FRM1) = THIS_NAME Then
                  ' ** This is first menu in path.
2310              blnGoingToReport = True
2320              Select Case IsNull(garr_varGoToReport(GTR_FRM2))
                  Case True
                    ' ** No more menus.
2330                .Controls(garr_varGoToReport(GTR_CTL1)).SetFocus
2340                lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
2350                lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
2360                .GoToReport_arw_menu_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
2370                .GoToReport_arw_menu_img.Visible = True
2380                DoEvents
2390                .TimerInterval = 50&
2400              Case False
                    ' ** Moving on.
2410                .Controls(garr_varGoToReport(GTR_CTL1)).SetFocus
2420                lngBtn_Top = .Controls(garr_varGoToReport(GTR_CTL1)).Top
2430                lngBtn_Left = .Controls(garr_varGoToReport(GTR_CTL1)).Left
2440                .GoToReport_arw_menu_img.Top = lngBtn_Top + (3& * lngTpp)  ' ** 45 Twips (3 pixels) beneath button top.
2450                .GoToReport_arw_menu_img.Visible = True
2460                Select Case garr_varGoToReport(GTR_INV1)
                    Case True
2470                  DoEvents
2480                  .TimerInterval = 50&
2490                Case False
                      ' ** Just sit there, and we're done.
2500                  blnGoingToReport = False
2510                  DoCmd.Hourglass False
2520                  gblnGoToReport = False
2530                  blnGTR_Emblem = False
2540                  GTREmblem_Off  ' ** Procedure: Below.
2550                  Beep
2560                End Select
2570              End Select
2580            Else
                  ' ** Shouldn't be here, then!
2590              blnGoingToReport = False
2600              DoCmd.Hourglass False
2610              gblnGoToReport = False
2620              blnGTR_Emblem = False
2630              GTREmblem_Off  ' ** Procedure: Below.
2640              Beep
2650            End If
2660          End Select
2670        End If  ' ** blnGTR_Emblem.
2680      End If  ' ** gblnGoToReport.
2690      If gblnTimer = True Then
2700        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
2710          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
2720          gblnTimer = False
2730          DoCmd.Hourglass False
2740        End If
2750      End If
2760    End With

        'HERE! 1  Form_Open()  frmMenu_Asset
        'HERE! 2  Form_Timer()  frmMenu_Asset

EXITP:
2770    Exit Sub

ERRH:
2780    Select Case ERR.Number
        Case Else
2790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2800    End Select
2810    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

2910    With Me

2920      If gblnSignal = False Then  ' ** Borrowing this variable.
2930        DoCmd.Hourglass True
2940        DoEvents
2950      End If

          ' ** Make sure all reports are closed.
2960      Do While Reports.Count > 0
2970        DoCmd.Close acReport, Reports(0).Name
2980      Loop

2990      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3000      If blnDontEmpty = False Then
3010        Set dbs = CurrentDb
            ' ** Empty tmpAccount.
3020        Set qdf = dbs.QueryDefs("qryPortfolioModeling_Select_10")
3030        qdf.Execute
3040        Set qdf = Nothing
3050        dbs.Close
3060        Set dbs = Nothing
3070        DoEvents
3080      End If  ' ** blnDontEmpty.

3090      blnGoingToReport = False

3100      gstrFormQuerySpec = vbNullString
3110      gblnAdmin = False

3120      If blnJustClose = False Then

3130        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3140        DoEvents

3150        gblnSetFocus = True
3160        gstrReturningForm = THIS_NAME
3170        Select Case strCallingForm
            Case "frmMenu_Main"
3180          DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Title"
3190        End Select

3200      End If

3210    End With

EXITP:
3220    Set qdf = Nothing
3230    Set dbs = Nothing
3240    Exit Sub

ERRH:
3250    DoCmd.Hourglass False
3260    Select Case ERR.Number
        Case Else
3270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3280    End Select
3290    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3310    With Me
3320      If .cmdClose.Caption <> "&Close" Then
3330        .cmdClose.Caption = "&Close"
3340      End If
3350    End With

EXITP:
3360    Exit Sub

ERRH:
3370    Select Case ERR.Number
        Case Else
3380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3390    End Select
3400    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

3510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3520    Exit Sub

ERRH:
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Sub

Private Sub Header_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_box_MouseMove"

3610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub Header_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl_MouseMove"

3710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub Header_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Header_lbl2_MouseMove"

3810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub cmdClose_Click()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3910    With Me
3920      If blnFormOpening = False Then

3930        DoCmd.Hourglass True
3940        DoEvents

3950        .TimerInterval = 0&

3960        blnJustClose = False
3970        DoCmd.Close acForm, THIS_NAME

3980      End If
3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    blnFormOpening = False
4020    DoCmd.Hourglass False
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4120    intShiftDown = (Shift And acShiftMask) > 0
4130    intAltDown = (Shift And acAltMask) > 0
4140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyTab
4170        With Me
4180          intRetVal = 0
4190          .cmdAddEditAsset.SetFocus
4200        End With
4210      End Select
4220    End If

        ' ** Shift keys.
4230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4240      Select Case intRetVal
          Case vbKeyTab
4250        With Me
4260          intRetVal = 0
4270          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
4280            .chkLoadTimer.SetFocus
4290          Else
4300            .cmdPortfolioModel.SetFocus
4310          End If
4320        End With
4330      End Select
4340    End If

EXITP:
4350    KeyCode = intRetVal
4360    Exit Sub

ERRH:
4370    intRetVal = 0
4380    Select Case ERR.Number
        Case Else
4390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4400    End Select
4410    Resume EXITP

End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_MouseMove"

4510    With Me
4520      .cmdClose.Caption = "< Ba&ck"
4530    End With

EXITP:
4540    Exit Sub

ERRH:
4550    Select Case ERR.Number
        Case Else
4560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4570    End Select
4580    Resume EXITP

End Sub

Private Sub cmdAddEditAsset_Click()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddEditAsset_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnNegShareface As Boolean, lngNegSharefaces As Long, arr_varNegShareface As Variant
        Dim strDocName As String, strMsg As String
        Dim lngX As Long

        ' ** Array: arr_varNegShareface().
        Const N_ACTNO As Integer = 0
        Const N_SNAM  As Integer = 1
        'Const N_ASTNO As Integer = 2
        Const N_DESC  As Integer = 3
        'Const N_ADATE As Integer = 4
        Const N_SHRS  As Integer = 5

4610    With Me
4620      If blnFormOpening = False Then

4630        DoCmd.Hourglass True
4640        DoEvents

4650        .TimerInterval = 0&

4660        blnFormOpening = True
4670        blnNegShareface = False

4680        Set dbs = CurrentDb
4690        With dbs

              ' ** MasterAsset, just marketvaluecurrent = Null, with marketvaluecurrent_new.
4700          Set qdf = dbs.QueryDefs("qryPricing_22")
4710          Set rst = qdf.OpenRecordset
4720          If rst.BOF = True And rst.EOF = True Then
                ' ** All's well.
4730            rst.Close
4740          Else
4750            rst.Close
                ' ** Update qryPricing_22 (MasterAsset, just marketvaluecurrent = Null, with marketvaluecurrent_new).
4760            Set qdf = dbs.QueryDefs("qryPricing_23")
4770            qdf.Execute
4780          End If
4790          Set rst = Nothing
4800          Set qdf = Nothing
4810          DoEvents

              ' ** ActiveAssets, just shareface < 0.
4820          Set qdf = dbs.QueryDefs("qryPricing_16")
4830          Set rst = qdf.OpenRecordset
4840          With rst
4850            If .BOF = True And .EOF = True Then
                  ' ** Everything's fine.
4860            Else
                  ' ** EEEK! Not another one?
4870              blnNegShareface = True
4880              .MoveLast
4890              lngNegSharefaces = .RecordCount
4900              .MoveFirst
4910              arr_varNegShareface = .GetRows(lngNegSharefaces)
                  ' ******************************************************
                  ' ** Array: arr_varNegShareface()
                  ' **
                  ' **   Field  Element  Name                 Constant
                  ' **   =====  =======  ===================  ==========
                  ' **     1       0     accountno            N_ACTNO
                  ' **     2       1     shortname            N_SNAM
                  ' **     3       2     assetno              N_ASTNO
                  ' **     4       3     asset_description    N_DESC
                  ' **     5       4     assetdate            N_ADATE
                  ' **     6       5     shareface            N_SHRS
                  ' **
                  ' ******************************************************
4920            End If
4930            .Close
4940          End With
4950          Set rst = Nothing
4960          Set qdf = Nothing
4970          DoEvents

              ' ** Empty tmpEdit09.
4980          Set qdf = .QueryDefs("qryAssets_16")
4990          qdf.Execute
5000          Set qdf = Nothing
5010          DoEvents

              ' ** Append qryAssets_03 (qryAssets_02 (MasterAsset, linked to assettype), with add'l fields),
              ' ** linked to qryAssets_04 (qryAssets_04c (Union of qryAssets_04a (Ledger, grouped by assetno, with cnt),
              ' ** qryAssets_04b (LedgerArchive, grouped by assetno, with cnt)), grouped and summed by assetno), to tmpEdit09.
5020          Set qdf = .QueryDefs("qryAssets_05")
5030          qdf.Execute
5040          Set qdf = Nothing
5050          DoEvents

5060          .Close
5070        End With
5080        Set dbs = Nothing

5090        If blnNegShareface = True And gblnGoToReport = False Then
5100          strMsg = "A check of currently held assets shows " & CStr(lngNegSharefaces) & " Tax Lot" & _
                IIf(lngNegSharefaces > 1&, "s", "") & " with negative shares." & vbCrLf
5110          strMsg = strMsg & vbCrLf  ' ** 2nd CrLf.
5120          If lngNegSharefaces <= 5& Then
5130            For lngX = 0& To (lngNegSharefaces - 1&)
5140              strMsg = strMsg & "Account: " & arr_varNegShareface(N_ACTNO, lngX) & "  " & arr_varNegShareface(N_SNAM, lngX) & vbCrLf
5150              strMsg = strMsg & "Asset: " & arr_varNegShareface(N_DESC, lngX) & vbCrLf
5160              strMsg = strMsg & "Share/Face: " & CStr(arr_varNegShareface(N_SHRS, lngX)) & vbCrLf & vbCrLf
5170            Next
5180          End If
5190          strMsg = strMsg & "Contact Delta Data, Inc., for assistance."
5200          DoCmd.Hourglass False
5210          MsgBox strMsg, vbInformation + vbOKOnly, "Negative Shares Present"
5220          DoCmd.Hourglass True
5230          DoEvents
5240        End If

5250        gblnSetFocus = True
5260        blnJustClose = True
5270        strDocName = "frmAssets"
5280        DoCmd.Close acForm, THIS_NAME
5290        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
5300        If gblnGoToReport = True Then
5310          Forms(strDocName).TimerInterval = 50&
5320          blnGoingToReport = False
5330          blnGTR_Emblem = False
5340        End If
5350        DoEvents

5360        DoCmd.Hourglass False

5370      End If
5380    End With

EXITP:
5390    Set rst = Nothing
5400    Set qdf = Nothing
5410    Set dbs = Nothing
5420    Exit Sub

ERRH:
5430    blnFormOpening = False
5440    DoCmd.Hourglass False
5450    Select Case ERR.Number
        Case Else
5460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5470    End Select
5480    Resume EXITP

End Sub

Private Sub cmdAddEditAsset_KeyDown(KeyCode As Integer, Shift As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddEditAsset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5520    intShiftDown = (Shift And acShiftMask) > 0
5530    intAltDown = (Shift And acAltMask) > 0
5540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5560      Select Case intRetVal
          Case vbKeyTab
5570        With Me
5580          intRetVal = 0
5590          .cmdAssetPricing.SetFocus
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab
5650        With Me
5660          intRetVal = 0
5670          .cmdClose.SetFocus
5680        End With
5690      End Select
5700    End If

EXITP:
5710    KeyCode = intRetVal
5720    Exit Sub

ERRH:
5730    intRetVal = 0
5740    Select Case ERR.Number
        Case Else
5750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5760    End Select
5770    Resume EXITP

End Sub

Private Sub cmdAssetPricing_Click()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAssetPricing_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnNegShareface As Boolean, lngNegSharefaces As Long, arr_varNegShareface As Variant
        Dim strDocName As String, strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean, blnShowAll As Boolean
        Dim lngX As Long

        ' ** Array: arr_varNegShareface().
        Const N_ACTNO As Integer = 0
        Const N_SNAM  As Integer = 1
        'Const N_ASTNO As Integer = 2
        Const N_DESC  As Integer = 3
        'Const N_ADATE As Integer = 4
        Const N_SHRS  As Integer = 5

5810    With Me
5820      If blnFormOpening = False Then

5830        DoCmd.Hourglass True
5840        DoEvents

5850        .TimerInterval = 0&
5860        blnContinue = True

5870        blnNegShareface = False: blnShowAll = False

5880        Set dbs = CurrentDb
5890        Set rst = dbs.OpenRecordset("masterasset", dbOpenDynaset, dbReadOnly)
5900        If rst.BOF = True And rst.EOF = True Then
              ' ** The default asset should always be present!
5910          blnContinue = False
5920          rst.Close
5930          dbs.Close
5940          DoCmd.Hourglass False
5950          Beep
5960          MsgBox "There are no assets to price.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
5970        Else

5980          rst.Close
5990          Set rst = Nothing

              ' ** ActiveAssets, just shareface < 0.
6000          Set qdf = dbs.QueryDefs("qryPricing_16")
6010          Set rst = qdf.OpenRecordset
6020          With rst
6030            If .BOF = True And .EOF = True Then
                  ' ** Everything's fine.
6040            Else
                  ' ** EEEK! Not another one?
6050              blnNegShareface = True
6060              .MoveLast
6070              lngNegSharefaces = .RecordCount
6080              .MoveFirst
6090              arr_varNegShareface = .GetRows(lngNegSharefaces)
                  ' ******************************************************
                  ' ** Array: arr_varNegShareface()
                  ' **
                  ' **   Field  Element  Name                 Constant
                  ' **   =====  =======  ===================  ==========
                  ' **     1       0     accountno            N_ACTNO
                  ' **     2       1     shortname            N_SNAM
                  ' **     3       2     assetno              N_ASTNO
                  ' **     4       3     asset_description    N_DESC
                  ' **     5       4     assetdate            N_ADATE
                  ' **     6       5     shareface            N_SHRS
                  ' **
                  ' ******************************************************
6100            End If
6110            .Close
6120          End With
6130          Set rst = Nothing
6140          Set qdf = Nothing

              ' ** MasterAsset, just marketvaluecurrent = Null, with marketvaluecurrent_new.
6150          Set qdf = dbs.QueryDefs("qryPricing_22")
6160          Set rst = qdf.OpenRecordset
6170          If rst.BOF = True And rst.EOF = True Then
                ' ** All's well.
6180            rst.Close
6190          Else
6200            rst.Close
                ' ** Update qryPricing_22 (MasterAsset, just marketvaluecurrent = Null, with marketvaluecurrent_new).
6210            Set qdf = dbs.QueryDefs("qryPricing_23")
6220            qdf.Execute
6230          End If
6240          Set rst = Nothing
6250          Set qdf = Nothing

              ' ** MasterAsset, just shareface > 0.
6260          Set qdf = dbs.QueryDefs("qryPricing_15")
6270          Set rst = qdf.OpenRecordset
6280          If rst.BOF = True And rst.EOF = True Then
6290            rst.Close
6300            Set rst = dbs.OpenRecordset("ActiveAssets", dbOpenDynaset, dbReadOnly)
6310            If rst.BOF = True And rst.EOF = True Then
                  ' ** Yup. No assets.
6320              If gblnGoToReport = True Then
6330                msgResponse = vbYes
6340              Else
6350                DoCmd.Hourglass False
6360                Beep
6370                msgResponse = MsgBox("There are no assets held by accounts." & vbCrLf & vbCrLf & _
                      "Do you wish to proceed anyway?", vbQuestion + vbYesNo, "No Currently Held Assets")
6380              End If
6390              If msgResponse <> vbYes Then
6400                blnContinue = False
6410                blnGoingToReport = False
6420                gblnGoToReport = False
6430              Else
6440                DoCmd.Hourglass True
6450                DoEvents
6460                blnShowAll = True
6470                blnFormOpening = False
6480              End If
6490            Else
                  ' ** Something's way off!
6500              If gblnGoToReport = False Then
6510                DoCmd.Hourglass False
6520                blnContinue = False
6530                Beep
6540                MsgBox "Though the Master Asset table shows no assets" & vbCrLf & _
                      "held by any account, there are active Tax Lots present." & vbCrLf & vbCrLf & _
                      "Contact Delta Data, Inc., for assistance.", vbCritical + vbOKOnly, "Asset List Error"
6550                blnFormOpening = False
6560              End If
6570            End If
6580            rst.Close
6590            dbs.Close
6600          Else
6610            rst.Close
6620            dbs.Close
6630          End If
6640        End If
6650        Set rst = Nothing
6660        Set qdf = Nothing
6670        Set dbs = Nothing

6680        If blnContinue = True Then

6690          If blnNegShareface = True And gblnGoToReport = False Then
6700            strMsg = "A check of currently held assets shows " & CStr(lngNegSharefaces) & " Tax Lot" & _
                  IIf(lngNegSharefaces > 1&, "s", "") & " with negative shares." & vbCrLf
6710            strMsg = strMsg & vbCrLf  ' ** 2nd CrLf.
6720            If lngNegSharefaces <= 5& Then
6730              For lngX = 0& To (lngNegSharefaces - 1&)
6740                strMsg = strMsg & "Account: " & arr_varNegShareface(N_ACTNO, lngX) & "  " & arr_varNegShareface(N_SNAM, lngX) & vbCrLf
6750                strMsg = strMsg & "Asset: " & arr_varNegShareface(N_DESC, lngX) & vbCrLf
6760                strMsg = strMsg & "Share/Face: " & CStr(arr_varNegShareface(N_SHRS, lngX)) & vbCrLf & vbCrLf
6770              Next
6780            End If
6790            strMsg = strMsg & "Contact Delta Data, Inc., for assistance."
6800            DoCmd.Hourglass False
6810            MsgBox strMsg, vbInformation + vbOKOnly, "Negative Shares Present"
6820            DoCmd.Hourglass True
6830            DoEvents
6840          End If

6850          Set dbs = CurrentDb

              ' ** Empty tblPricing_MasterAsset.
6860          Set qdf = dbs.QueryDefs("qryPricing_05")
6870          qdf.Execute
6880          Set qdf = Nothing

              ' ** Empty tblPricing_Cusip.
6890          Set qdf = dbs.QueryDefs("qryPricing_06")
6900          qdf.Execute
6910          Set qdf = Nothing

              ' ** Append MasterAsset to tblPricing_MasterAsset.
6920          Set qdf = dbs.QueryDefs("qryPricing_07")
6930          qdf.Execute
6940          Set qdf = Nothing

6950          dbs.Close
6960          Set dbs = Nothing

6970          gblnSetFocus = True
6980          blnJustClose = True
6990          strDocName = "frmAssetPricing"
7000          DoCmd.Close acForm, THIS_NAME
7010          DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & Format(blnShowAll, "True/False")
7020          If gblnGoToReport = True Then
7030            Forms(strDocName).TimerInterval = 50&
7040            blnGoingToReport = False
7050            blnGTR_Emblem = False
7060          End If
7070          DoEvents

7080        End If  ' ** blnContinue.

7090        DoCmd.Hourglass False

7100      End If
7110    End With

EXITP:
7120    Set rst = Nothing
7130    Set qdf = Nothing
7140    Set dbs = Nothing
7150    Exit Sub

ERRH:
7160    blnFormOpening = False
7170    DoCmd.Hourglass False
7180    Select Case ERR.Number
        Case Else
7190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7200    End Select
7210    Resume EXITP

End Sub

Private Sub cmdLocations_Click()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdLocations_Click"

        Dim varTmp00 As Variant, varTmp01 As Variant
        Dim strDocName As String

7310    With Me
7320      If blnFormOpening = False Then

7330        DoCmd.Hourglass True
7340        DoEvents

7350        .TimerInterval = 0&

7360        strDocName = "frmRpt_Locations"
7370        Select Case gblnGoToReport
            Case True
7380          gblnSetFocus = True
7390          blnJustClose = True
7400          DoCmd.Close acForm, THIS_NAME
7410          DoCmd.OpenForm strDocName, , , , , , THIS_NAME
7420          Forms(strDocName).TimerInterval = 50&
7430          blnGoingToReport = False
7440          blnGTR_Emblem = False
7450        Case False
              ' ** Location, grouped by Location_Name, with cnt.
7460          varTmp00 = DCount("*", "qryLocation_08")
7470          varTmp01 = DCount("*", "ActiveAssets")
7480          If varTmp01 > 0 Then
7490            gblnSetFocus = True
7500            blnJustClose = True
7510            blnFormOpening = True
7520            DoCmd.Close acForm, THIS_NAME
7530            DoCmd.OpenForm strDocName, , , , , , THIS_NAME
7540          Else
7550            DoCmd.Hourglass False
7560            Beep
7570            If varTmp00 > 0 Then
7580              MsgBox "There are no assets held by accounts.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
7590            Else
7600              MsgBox "There are no listed locations, and no assets held by accounts.", _
                    vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
7610            End If
7620          End If
7630        End Select

7640        DoCmd.Hourglass False

7650      End If
7660    End With

EXITP:
7670    Exit Sub

ERRH:
7680    blnFormOpening = False
7690    DoCmd.Hourglass False
7700    Select Case ERR.Number
        Case Else
7710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7720    End Select
7730    Resume EXITP

End Sub

Private Sub cmdHoldings_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHoldings_Click"

        Dim strDocName As String
        Dim varTmp00 As Variant

7810    With Me
7820      If blnFormOpening = False Then

7830        DoCmd.Hourglass True
7840        DoEvents

7850        .TimerInterval = 0&

7860        strDocName = "frmRpt_Holdings"
7870        Select Case gblnGoToReport
            Case True
7880          gblnSetFocus = True
7890          blnJustClose = True
7900          DoCmd.Close acForm, THIS_NAME
7910          DoCmd.OpenForm strDocName, , , , , , THIS_NAME
7920          Forms(strDocName).TimerInterval = 50&
7930          blnGoingToReport = False
7940          blnGTR_Emblem = False
7950        Case False
7960          varTmp00 = DCount("*", "ActiveAssets")
7970          If varTmp00 > 0 Then
7980            gblnSetFocus = True
7990            blnJustClose = True
8000            blnFormOpening = True
8010            DoCmd.Close acForm, THIS_NAME
8020            DoCmd.OpenForm strDocName, , , , , , THIS_NAME
8030          Else
8040            DoCmd.Hourglass False
8050            Beep
8060            MsgBox "There are no assets held by accounts.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
8070          End If
8080        End Select

8090        DoCmd.Hourglass False

8100      End If
8110    End With

EXITP:
8120    Exit Sub

ERRH:
8130    blnFormOpening = False
8140    DoCmd.Hourglass False
8150    Select Case ERR.Number
        Case Else
8160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8170    End Select
8180    Resume EXITP

End Sub

Private Sub cmdHistory_Click()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistory_Click"

        Dim varTmp00 As Variant
        Dim strDocName As String

8210    With Me
8220      If blnFormOpening = False Then

8230        DoCmd.Hourglass True
8240        DoEvents

8250        .TimerInterval = 0&

8260        strDocName = "frmRpt_AssetHistory"
8270        Select Case gblnGoToReport
            Case True
8280          gblnSetFocus = True
8290          blnJustClose = True
8300          DoCmd.Close acForm, THIS_NAME
8310          DoCmd.OpenForm strDocName, , , , , , THIS_NAME
8320          Forms(strDocName).TimerInterval = 50&
8330          blnGoingToReport = False
8340          blnGTR_Emblem = False
8350        Case False
              ' ** Union of qryAssetHistory_70a (Ledger, grouped by assetno, with cnt),
              ' ** qryAssetHistory_70b (LedgerArchive, grouped by assetno, with cnt).
8360          varTmp00 = DCount("*", "qryAssetHistory_70c")
8370          If varTmp00 > 0 Then
8380            gblnSetFocus = True
8390            blnJustClose = True
8400            blnFormOpening = True
8410            DoCmd.Close acForm, THIS_NAME
8420            DoCmd.OpenForm strDocName, , , , , , THIS_NAME
8430          Else
8440            DoCmd.Hourglass False
8450            Beep
8460            MsgBox "There are no asset transactions for accounts.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
8470          End If
8480        End Select

8490        DoCmd.Hourglass False

8500      End If
8510    End With

EXITP:
8520    Exit Sub

ERRH:
8530    blnFormOpening = False
8540    DoCmd.Hourglass False
8550    Select Case ERR.Number
        Case Else
8560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8570    End Select
8580    Resume EXITP

End Sub

Private Sub cmdPortfolioModel_Click()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPortfolioModel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long, arr_varRec As Variant
        Dim strDocName As String, strMsg As String
        Dim blnContinue As Boolean
        Dim lngX As Long

        ' ** Array: arr_varRec().
        'Const R_ASTNO  As Integer = 0
        Const R_CUSIP  As Integer = 1
        Const R_DESC   As Integer = 2
        'Const R_MKTVAL As Integer = 3

8610    With Me
8620      If blnFormOpening = False Then

8630        DoCmd.Hourglass True
8640        DoEvents

8650        .TimerInterval = 0&

8660        blnContinue = True

8670        Set dbs = CurrentDb
8680        With dbs
              ' ** MasterAsset, just marketvaluecurrent = Null.
8690          Set qdf = .QueryDefs("qryPortfolioModeling_Select_02")
8700          Set rst = qdf.OpenRecordset
8710          With rst
8720            If .BOF = True And .EOF = True Then
                  ' ** All's well.
8730              lngRecs = 0&
8740            Else
8750              blnContinue = False
8760              .MoveLast
8770              lngRecs = .RecordCount
8780              .MoveFirst
8790              arr_varRec = .GetRows(lngRecs)
                  ' *******************************************************
                  ' ** Array: arr_varRec()
                  ' **
                  ' **   Field  Element  Name                  Constant
                  ' **   =====  =======  ====================  ==========
                  ' **     1       0     assetno               R_ASTNO
                  ' **     2       1     cusip                 R_CUSIP
                  ' **     3       2     description           R_DESC
                  ' **     4       3     marketvaluecurrent    R_MKVAL
                  ' **
                  ' *******************************************************
8800            End If
8810            .Close
8820          End With
8830          Set rst = Nothing
8840          Set qdf = Nothing
              ' ** Empty tmpAccount.
8850          Set qdf = .QueryDefs("qryPortfolioModeling_Select_10")
8860          qdf.Execute
8870          Set qdf = Nothing
8880          DoEvents
              ' ** Append qryAccountMenu_01_10 (qryAccountProfile_01_01 (Account, linked to qryAccountProfile_01_02
              ' ** (Ledger, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_03
              ' ** (LedgerArchive, grouped by accountno, for ledger_HIDDEN = True, with cnt), qryAccountProfile_01_04
              ' ** (ActiveAssets, grouped, with cnt, by accountno), with S_PQuotes, L_PQuotes, ActiveAssets cnt),
              ' ** linked to qryAccountProfile_01_08 (qryAccountProfile_01_07 (qryAccountProfile_01_05 (Account,
              ' ** with IsNum), grouped, just IsNum = False, with cnt_acct), linked to qryAccountProfile_01_06
              ' ** (qryAccountProfile_01_05 (Account, with IsNum), grouped, just IsNum = True, with cnt_acct),
              ' ** with IsNum, cnt_num), just accountno, with acct_sort) to tmpAccount.
8890          Set qdf = .QueryDefs("qryPortfolioModeling_Select_11")
8900          qdf.Execute
8910          Set qdf = Nothing
8920          DoEvents
8930          .Close
8940        End With
8950        Set dbs = Nothing
8960        DoEvents

8970        Select Case blnContinue
            Case True
8980          gblnSetFocus = True: blnDontEmpty = True
8990          blnJustClose = True
9000          blnFormOpening = True
9010          strDocName = "frmPortfolioModeling_Select"
9020          DoCmd.Close acForm, THIS_NAME  ' ** Close.
9030          DoCmd.OpenForm strDocName, , , , , , THIS_NAME
9040          If gblnGoToReport = True Then
9050            Forms(strDocName).TimerInterval = 50&
9060            blnGoingToReport = False
9070            blnGTR_Emblem = False
9080          End If
9090        Case False
9100          If lngRecs = 1& Then
9110            strMsg = "There is 1 asset with no Market Unit Value whatsoever!" & vbCrLf
9120            strMsg = strMsg & "CUSIP: " & arr_varRec(R_CUSIP, 0) & vbCrLf
9130            strMsg = strMsg & arr_varRec(R_DESC, 0) & vbCrLf & vbCrLf
9140          Else
9150            strMsg = "There are " & CStr(lngRecs) & " assets with no Market Unit Value whatsoever!" & vbCrLf
9160            If lngRecs < 5& Then
9170              For lngX = 0& To (lngRecs - 1&)
9180                strMsg = strMsg & "CUSIP: " & arr_varRec(R_CUSIP, lngX) & vbCrLf
9190                strMsg = strMsg & arr_varRec(R_DESC, lngX) & vbCrLf
9200              Next
9210              strMsg = strMsg & vbCrLf
9220            Else
9230              strMsg = strMsg & "  (see Asset Pricing)" & vbCrLf & vbCrLf
9240            End If
9250          End If
9260          strMsg = strMsg & "All assets must have a zero or positive Market Unit Value." & vbCrLf
9270          strMsg = strMsg & "Correct the omission using Asset Pricing on this menu."
9280          DoCmd.Hourglass False
9290          Beep
9300          MsgBox strMsg, vbOKOnly, "Asset Has Null Market Value"
9310        End Select

9320        DoCmd.Hourglass False

9330      End If
9340    End With

EXITP:
9350    Set rst = Nothing
9360    Set qdf = Nothing
9370    Set dbs = Nothing
9380    Exit Sub

ERRH:
9390    blnFormOpening = False
9400    DoCmd.Hourglass False
9410    Select Case ERR.Number
        Case Else
9420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9430    End Select
9440    Resume EXITP

End Sub

Private Sub cmdPortfolioModel_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPortfolioModel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9520    intShiftDown = (Shift And acShiftMask) > 0
9530    intAltDown = (Shift And acAltMask) > 0
9540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9560      Select Case intRetVal
          Case vbKeyTab
9570        With Me
9580          intRetVal = 0
9590          If .chkLoadTimer.Visible = True And .chkLoadTimer.Enabled = True Then
9600            .chkLoadTimer.SetFocus
9610          Else
9620            .cmdClose.SetFocus
9630          End If
9640        End With
9650      End Select
9660    End If

        ' ** Shift keys.
9670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9680      Select Case intRetVal
          Case vbKeyTab
9690        With Me
9700          intRetVal = 0
9710          .cmdLocations.SetFocus
9720        End With
9730      End Select
9740    End If

EXITP:
9750    KeyCode = intRetVal
9760    Exit Sub

ERRH:
9770    intRetVal = 0
9780    Select Case ERR.Number
        Case Else
9790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9800    End Select
9810    Resume EXITP

End Sub

Private Sub chkLoadTimer_AfterUpdate()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_AfterUpdate"

9910    With Me
9920      If IsLoaded("frmLoadTimer", acForm) = False Then  ' ** Module Function: modFileUtilities.
            ' ** Don't turn it off if it's still playing out a timing sequence.
9930        gblnTimer = False
9940      End If
9950      Select Case .chkLoadTimer
          Case True
9960        .chkLoadTimer_lbl.FontBold = True
9970        If .chkLoadTimer.Visible = True Then  ' ** This may get turned off by the Form_Timer() sequence.
9980          gblnTimer = True
9990        End If
10000     Case False
10010       .chkLoadTimer_lbl.FontBold = False
10020     End Select
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

Private Sub chkLoadTimer_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLoadTimer_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10120   intShiftDown = (Shift And acShiftMask) > 0
10130   intAltDown = (Shift And acAltMask) > 0
10140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10160     Select Case intRetVal
          Case vbKeyTab
10170       With Me
10180         intRetVal = 0
10190         .cmdClose.SetFocus
10200       End With
10210     End Select
10220   End If

        ' ** Shift keys.
10230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10240     Select Case intRetVal
          Case vbKeyTab
10250       With Me
10260         intRetVal = 0
10270         .cmdPortfolioModel.SetFocus
10280       End With
10290     End Select
10300   End If

EXITP:
10310   KeyCode = intRetVal
10320   Exit Sub

ERRH:
10330   intRetVal = 0
10340   Select Case ERR.Number
        Case Else
10350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10360   End Select
10370   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

10410   With Me
10420     blnGTR_Emblem = False
10430     For lngX = 1& To 24&
10440 On Error Resume Next
10450       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
10460       If ERR.Number <> 0 Then
10470 On Error GoTo ERRH
10480         Exit For
10490       Else
10500 On Error GoTo ERRH
10510       End If
10520     Next
10530   End With

EXITP:
10540   Exit Sub

ERRH:
10550   Select Case ERR.Number
        Case Else
10560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10570   End Select
10580   Resume EXITP

End Sub
