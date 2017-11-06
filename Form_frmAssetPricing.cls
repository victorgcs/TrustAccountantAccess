VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssetPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssetPricing"

'VGC 10/02/2017: CHANGES!
' ** TimerInterval: Set by GoToReport in frmMenu_Asset.

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   IncludeCurrency()
' #######################################

' ** For GTR, either just opens and stops, or shows arrows, sets focus, and continues.

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   New Date:         D {Date_New}
' **   Reset:            E {cmdReset}
' **   Show Hidden:      H {chkShowHidden}
' **   History:          I {cmdHistory}
' **   Include Currency: L {chkIncludeCurrency}  (cmdUpdate has to be Alt+U)
' **   Market Value:     M {marketvalue on frmAssetPricing_Sub}
' **   Show AssetNo:     N {chkShowAssetNo}
' **   Only Assets Held: O {opgShow_optHeldOnly}
' **   Print:            P {cmdPrintReport}
' **   Price:            R {cmdPrice}
' **   Shorten:          S {cmdShorten on frmAssetPricing_Sub}
' **   All Assets:       T {opgShow_optAll}
' **   Update:           U {cmdUpdate}
' **   New Value:        V {marketvalue on frmAssetPricing_Sub}
' **   Widen:            W {cmdWiden on frmAssetPricing_Sub}
' **   Exit:             X {cmdClose}
' **   Yield:            Y {yield_entry on frmAssetPricing_Sub}

' ** Shortcut F-Keys responsive from this form:
' **   Recalc:          F5 {ForExTots}
' **   Date Picker:     F7 {cmdCalendar}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmAssetPricing_Sub}

' ** At Open:
' **   Enabled:
' **     True:
' **
' **     False:
' **       cusip
' **       totdesc
' **       marketvaluecurrent
' **       yield
' **       marketvalue
' **   Locked:
' **     True:
' **       cusip
' **       totdesc
' **       marketvaluecurrent
' **       yield
' **       marketvalue
' **     False:
' **
' ** Disabled:
' **   BackColor:   8421440   CLR_LTTEAL
' **   ForeColor:   6052956   CLR_DKGRY
' **   BorderColor: 8421504   CLR_GRY
' ** Enabled:
' **   BackColor:   16777215  CLR_WHT
' **   ForeColor:   0         CLR_BLK
' ** Updated:
' **   ForeColor:   16774128  CLR_LTBLU  (ForeColor of solid Terminal font: marketvalue_upd, yield_upd, which
' **                                      marketvalue_upd, yield_upd depend on records being in tblPricing_Cusip.)

' ** When a valid date is entered:
' **   Me.yield.Enabled = True
' **   Me.yield.Locked = False
' **   Me.marketvalue.Enabled = True
' **   Me.marketvalue.Locked = False

' ** When Update is finished:
' **   Me.yield.Enabled = False
' **   Me.yield.Locked = True
' **   Me.marketvalue.Enabled = False
' **   Me.marketvalue.Locked = True

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngPrintReportLbl_Left As Long

Private datStartDate As Date, datEndDate As Date, blnFromCalendar As Boolean
Private blnHasForEx_Chk As Boolean, blnForExTots As Boolean, lngIsHidSum As Long
Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnSetPricingInvoked As Boolean, blnViewOnly As Boolean
Private blnTabBack As Boolean, blnKeyDown_Return As Boolean
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim usr As DAO.User, grp As DAO.Group
        Dim intPos01 As Integer
        Dim blnRetVal As Boolean, blnShowAll As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           intPos01 = InStr(.OpenArgs, "~")
150           If intPos01 > 0 Then

160             blnIsOpen = True

170             strCallingForm = Trim(.OpenArgs)
180             blnShowAll = CBool(Mid(strCallingForm, (intPos01 + 1)))
190             strCallingForm = Left(strCallingForm, (intPos01 - 1))

200             lngForm_Width = .Width
210             lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

                ' ** Just get these out of the way.
220             .form_width_line.Left = 0&
230             .GoToReport_Emblem_lbl.Left = 0&

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
240             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

250             CLR_DISABLED_FG = CLR_DKGRY
260             CLR_DISABLED_BG = CLR_LTTEAL

270             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
280             blnCalendar1_Focus = False: blnCalendar1_MouseDown = False

290             blnTabBack = False: blnKeyDown_Return = False: blnSetPricingInvoked = False: blnFromCalendar = False
300             blnGTR_Emblem = False: lngGTR_Stat = 0&
310             gblnMessage = False  ' ** Used for indicating cmdReset.
320             gblnHasForExThis = False: blnForExTots = False: blnHasForEx_Chk = True  ' ** Yes, True.

330             If gblnGoToReport = True Then
340               If IsNull(garr_varGoToReport(GTR_CTL2)) = True Then
350                 .cmdPrintReport_lbl2.Left = (lngPrintReportLbl_Left + .GoToReport_arw_pr_img_lbl.Width)
360                 .GoToReport_arw_pr_img.Visible = True
370                 DoEvents
380               Else
390                 Select Case garr_varGoToReport(GTR_CTL2)
                    Case "cmdPrice"
400                   If .GoToReport_arw_date_img.Visible = False Then
410                     .GoToReport_Emblem_01_img.Visible = True
420                     DoEvents
430                     .cmdCalendar_raised_img.Visible = False
440                     .cmdCalendar.Enabled = False
450                     .GoToReport_arw_date_img.Visible = True
460                     DoEvents
470                   End If
480                 Case "cmdHistory"
490                   .GoToReport_Emblem_01_img.Visible = True
500                   .GoToReport_arw_hist_img.Top = (.cmdHistory.Top + lngTpp)
510                   .GoToReport_arw_hist_img.Left = ((.cmdHistory.Left - .GoToReport_arw_hist_img.Width) + (2& * lngTpp))
520                   .GoToReport_arw_hist_img.Visible = True
530                   DoEvents
540                   .cmdReset.Enabled = False
550                   .cmdUpdate.Enabled = False
560                   If .cmdHistory.Enabled = True Then
570                     .cmdHistory.SetFocus
580                   Else
590                     .cmdClose.SetFocus
600                   End If
610                   DoEvents
620                 End Select
630               End If
640             End If

                ' ** Copyright Lebans Holdings 1999 Ltd.
                ' ** Create an instance of the Calendar class.
650             Set clsMonthClass = New clsMonthCal
                ' ** You MUST SET the class hWndForm prop!!!
660             clsMonthClass.hWndForm = .hwnd
                ' ** Let's default to PositionAtCursor.
670             clsMonthClass.PositionAtCursor = True

                ' ** Form_Load() happens between here...
680             OpeningGambit  ' ** Procedure: Below.
                ' ** And here...

690             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
700             If lngMonitorCnt > 1& Then
710               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
720               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
730               If lngMonitorNum > 1& Then
740                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
750               End If
760             End If

770             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

780             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
790             DoEvents

                ' ** This will set the gblnPricingAllowed variable.
800             blnRetVal = Security_PricingChk  ' ** Module Function: modSecurityFunctions.

810             blnViewOnly = False
820             For Each grp In DBEngine.Workspaces(0).Groups
830               If grp.Name = "ViewOnly" Then
840                 For Each usr In grp.Users
850                   If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
860                     blnViewOnly = True
870                     Exit For
880                   End If
890                 Next
900               End If
910             Next
920             DoEvents

930             .chkUpdated = False

940             If blnViewOnly = True Then
950               .cmdClose.SetFocus
960               .cmdUpdate.Enabled = False
970               .cmdReset.Enabled = False
980               .Date_New.Enabled = False
990               .Date_New.BorderColor = WIN_CLR_DISR
1000            End If

1010            Set dbs = CurrentDb
                ' ** Update tblPricing_MasterAsset yield = rate if rate > 0.
1020            Set qdf = dbs.QueryDefs("qryPricing_02")
1030            qdf.Execute
1040            Set qdf = Nothing
1050            DoEvents
                ' ** Update MasterAsset yield = rate if rate > 0.
1060            Set qdf = dbs.QueryDefs("qryPricing_02a")
1070            qdf.Execute
1080            Set qdf = Nothing
1090            DoEvents
                ' ** tblPricing_MasterAsset_History, grouped by currentDate.
1100            Set qdf = dbs.QueryDefs("qryPricing_History_06")
1110            Set rst = qdf.OpenRecordset
1120            If rst.BOF = True And rst.EOF = True Then
1130              If .GoToReport_arw_hist_img.Visible = True Then
1140                .cmdClose.SetFocus
1150              End If
1160              .cmdHistory.Enabled = False
1170            End If
1180            rst.Close
1190            dbs.Close
1200            Set rst = Nothing
1210            Set qdf = Nothing
1220            Set dbs = Nothing
1230            DoEvents

1240            If blnShowAll = True Then
1250              .opgShow = .opgShow_optAll.OptionValue
1260              .opgShow_optHeldOnly.Enabled = False
1270            End If
1280            opgShow_AfterUpdate  ' ** Procedure: Below.
1290            DoEvents
1300            chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
1310            DoEvents
1320            chkShowAssetNo_AfterUpdate  ' ** Procedure: Below.
1330            DoEvents
1340            chkShowHidden_AfterUpdate  ' ** Procedure: Below.
1350            DoEvents

1360            If lngMonitorCnt > 1& And blnIsOpen = False Then
1370              EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
1380              lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
1390              If lngMonitorNum > 1& Then
1400                LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
1410              End If
1420            End If

1430            Select Case blnViewOnly
                Case True
1440              .cmdClose.SetFocus
1450            Case False
1460              Select Case gblnGoToReport
                  Case True
1470                If ((garr_varGoToReport(GTR_CTL2) = "cmdPrice") Or _
                        (garr_varGoToReport(GTR_CTL2) = "cmdHistory" And .cmdHistory.Enabled = False)) Then
1480                  .Date_New.SetFocus
1490                  .cmdUpdate.Enabled = False  ' ** Don't enable till they've made changes.
1500                  .cmdReset.Enabled = False
1510                End If
1520              Case False
1530                .Date_New.SetFocus
1540                .cmdUpdate.Enabled = False  ' ** Don't enable till they've made changes.
1550                .cmdReset.Enabled = False
1560              End Select
1570            End Select

1580            DoCmd.Hourglass False

1590            blnIsOpen = False
1600            .frmAssetPricing_Sub.Form.IsOpen_Set blnIsOpen  ' ** Form Procedure: frmAssetPricing_Sub.

1610          Else
1620            Cancel = -1
1630          End If
1640        Else
1650          Cancel = -1
1660        End If
1670      Else
1680        Cancel = -1
1690      End If
1700    End With  ' ** Me.

1710    If gblnGoToReport = False Then
1720      DoCmd.Hourglass False
1730    End If

1740    If Cancel = -1 Then
1750      Beep
1760      MsgBox "This form is not to be opened independently.", _
            vbCritical + vbOKOnly, (Left(("Form Inaccessible" & Space(55)), 55) & "A01")
1770    End If

EXITP:
1780    Set rst = Nothing
1790    Set qdf = Nothing
1800    Set dbs = Nothing
1810    Set usr = Nothing
1820    Set grp = Nothing
1830    Exit Sub

ERRH:
1840    DoCmd.Hourglass False
1850    Select Case ERR.Number
        Case Else
1860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1870    End Select
1880    Resume EXITP

End Sub

Private Sub Form_Load()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

1910    With Me

1920      Set dbs = CurrentDb
1930      With dbs
            ' ** qryAssetPricing_04_01 (MasterAsset, just hidden assets), grouped, with cnt_astno.
1940        Set qdf = .QueryDefs("qryAssetPricing_04_02")
1950        Set rst = qdf.OpenRecordset
1960        With rst
1970          If .BOF = True And .EOF = True Then
1980            lngIsHidSum = 0&
1990          Else
2000            .MoveFirst
2010            If IsNull(![cnt_astno]) = True Then
2020              lngIsHidSum = 0&
2030            Else
2040              lngIsHidSum = ![cnt_astno]
2050            End If
2060          End If
2070          .Close
2080        End With
2090        Set rst = Nothing
2100        Set qdf = Nothing
2110        .Close
2120      End With
2130      Set dbs = Nothing

2140      If lngIsHidSum = 0& Then
2150        .chkShowHidden = False
2160        .chkShowHidden_lbl.Visible = False
2170        .chkShowHidden.Enabled = False
2180        .chkShowHidden_lbl2_dim.Visible = True
2190        .chkShowHidden_lbl2_dim_hi.Visible = True
2200        .chkShowHidden_lbl3_dim.Visible = True
2210        .chkShowHidden_lbl3_dim_hi.Visible = True
2220      Else
2230        .chkShowHidden_lbl.Visible = True
2240        .chkShowHidden.Enabled = True
2250        .chkShowHidden_lbl2_dim.Visible = False
2260        .chkShowHidden_lbl2_dim_hi.Visible = False
2270        .chkShowHidden_lbl3_dim.Visible = False
2280        .chkShowHidden_lbl3_dim_hi.Visible = False
2290      End If

2300    End With

EXITP:
2310    Set rst = Nothing
2320    Set qdf = Nothing
2330    Set dbs = Nothing
2340    Exit Sub

ERRH:
2350    THAT_PROC = THIS_PROC
2360    That_Erl = Erl
2370    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2380    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Market Value:     M {marketvalue on frmAssetPricing_Sub}
        ' **   Yield:            Y {yield_entry on frmAssetPricing_Sub}

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Recalc:          F5 {ForExTots}
        ' **   Date Picker:     F7 {cmdCalendar}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAssetPricing_Sub}

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyF5
2470        With Me
2480          intRetVal = 0
2490          blnForExTots = True
2500          .TimerInterval = 400&
2510        End With
2520      Case vbKeyF7
2530        With Me
2540          intRetVal = 0
2550          .cmdCalendar.SetFocus
2560          cmdCalendar_Click  ' ** Procedure: Below.
2570        End With
2580      Case vbKeyEscape
2590        intRetVal = 0
2600        cmdClose_Click  ' ** Procedure: Below.
2610      End Select
2620    End If

        ' ** Alt keys.
2630    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2640      Select Case intRetVal
          Case vbKeyM
2650        With Me
2660          intRetVal = 0
2670          .frmAssetPricing_Sub.SetFocus
2680          If .frmAssetPricing_Sub.Form.marketvalue.Enabled = True Then
2690            .frmAssetPricing_Sub.Form.marketvalue.SetFocus
2700          End If
2710        End With
2720      Case vbKeyX
2730        intRetVal = 0
2740        cmdClose_Click  ' ** Procedure: Below.
2750      Case vbKeyY
2760        With Me
2770          intRetVal = 0
2780          .frmAssetPricing_Sub.SetFocus
2790          If .frmAssetPricing_Sub.Form.yield_entry.Enabled = True Then
2800            .frmAssetPricing_Sub.Form.yield_entry.SetFocus
2810          End If
2820        End With
2830      End Select
2840    End If

        ' ** Ctrl keys.
2850    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyS
2870        With Me
2880          intRetVal = 0
2890          lngRecsCur = .frmAssetPricing_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_Sub.
2900          If lngRecsCur > 0& Then
2910            .frmAssetPricing_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmAssetPricing_Sub.
2920          Else
2930            Beep
2940          End If
2950        End With
2960      End Select
2970    End If

        ' ** Ctrl-Shift keys.
2980    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2990      Select Case intRetVal
          Case vbKeyF
3000        With Me
3010          intRetVal = 0
3020          .FocusHolder.SetFocus
3030        End With
3040      End Select
3050    End If

EXITP:
3060    KeyCode = intRetVal
3070    Exit Sub

ERRH:
3080    intRetVal = 0
3090    THAT_PROC = THIS_PROC
3100    That_Erl = Erl
3110    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
3120    Resume EXITP

End Sub

Private Sub Form_Timer()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim varTmp00 As Variant

        Static lngTmp01 As Long, lngTmp02 As Long

3210    With Me
3220      .TimerInterval = 0&
3230      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
3240        lngGTR_Stat = lngGTR_Stat + 1&
3250        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
3260        lngTmp01 = lngTmp01 + 1&
3270        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
3280        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
3290        If lngTmp01 > 1& Then
3300          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
3310        End If
3320        If lngTmp01 = 24& Then
3330          lngTmp01 = 0&
3340          lngTmp02 = lngTmp02 + 1&
3350          If lngTmp02 < 5& Then
3360            .TimerInterval = 50&
3370          Else
3380            blnGTR_Emblem = False
3390          End If
3400        Else
3410          .TimerInterval = 50&
3420        End If
3430      End If  ' ** blnGTR_Emblem.
3440      If gblnSetFocus = True Then
3450        gblnSetFocus = False
3460        strFormName = vbNullString
3470  On Error Resume Next
3480        Do Until Screen.ActiveForm.Name = strFormName
3490          DoCmd.SelectObject acForm, THIS_NAME, False
3500          strFormName = Screen.ActiveForm.Name
3510        Loop
3520  On Error GoTo ERRH
3530        DoCmd.Hourglass False
3540      End If  ' ** gblnSetFocus.
3550      If gblnGoToReport = True Then
3560        DoCmd.Hourglass True  ' ** Make sure it's still running.
3570        If ((lngGTR_Stat > 0& And lngGTR_Stat < 24&) Or (lngGTR_Stat > 24& And lngGTR_Stat < 48&) Or _
                (lngGTR_Stat > 48& And lngGTR_Stat < 72&) Or (lngGTR_Stat > 72&)) Then
              ' ** Just let the emblem go.
3580        Else
3590          If blnGTR_Emblem = False And IsNull(garr_varGoToReport(GTR_CTL2)) = False Then
3600            blnGTR_Emblem = True
3610            lngTmp01 = 0&: lngTmp02 = 0&
3620          End If
3630          DoEvents
3640          Select Case blnGoingToReport
              Case True
3650            Select Case blnGoingToReport2
                Case True
3660              .TimerInterval = 50&
3670              cmdPrice_Click  ' ** Procedure: Below.
3680            Case False
3690              Select Case IsNull(garr_varGoToReport(GTR_CTL2))
                  Case True
3700                DoCmd.Hourglass False  ' ** It should already be off.
3710                blnGoingToReport = False
3720                gblnGoToReport = False
3730                blnGTR_Emblem = False
3740                .GoToReport_arw_pr_img.Visible = False
3750                .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
3760                DoEvents
3770              Case False
3780                Select Case garr_varGoToReport(GTR_CTL2)
                    Case "cmdPrice"
3790                  blnGoingToReport2 = True
3800                  .Date_New = Date
3810                  Form_Click  ' ** Procedure: Below.
3820                  DoEvents
3830                  .cmdPrice.SetFocus
3840                  blnSetPricingInvoked = True
3850                  .GoToReport_arw_price_img.Visible = True
3860                  .GoToReport_arw_date_img.Visible = False
3870                  .cmdCalendar_raised_img.Visible = True
3880                  .cmdCalendar.Enabled = True
3890                  DoEvents
3900                  .TimerInterval = 50&  ' ** It should only pop out of the emblem loop at certain intervals.
3910                Case "cmdHistory"
3920                  .TimerInterval = 50&
3930                  cmdHistory_Click  ' ** Procedure: Below.
3940                End Select
3950              End Select
3960            End Select
3970          Case False
3980            DoCmd.Hourglass True  ' ** Make sure it's still running.
3990            DoEvents
4000            blnGoingToReport = True: blnGoingToReport2 = False
4010            Select Case IsNull(garr_varGoToReport(GTR_CTL2))
                Case True
4020              .GoToReport_arw_pr_img.Visible = True  ' ** Should already be visible.
4030              DoEvents
4040              If .cmdPrintReport.Enabled = True Then
4050                .cmdPrintReport.SetFocus
4060              Else
4070                .cmdClose.SetFocus
4080              End If
4090              DoEvents
4100              Beep
4110              DoCmd.Hourglass False
4120              DoEvents
4130              .TimerInterval = (GTR_WAIT * 2&)
4140            Case False
4150              Select Case garr_varGoToReport(GTR_CTL2)
                  Case "cmdPrice"
                    ' ** Don't turn it off so soon; arrow already on.
4160                .Date_New.SetFocus
4170                DoEvents
4180                .TimerInterval = GTR_WAIT
4190              Case "cmdHistory"
4200                Select Case .cmdHistory.Enabled
                    Case True
                      ' ** Arrow is already on.
4210                  .cmdHistory.SetFocus
4220                  DoEvents
4230                  .TimerInterval = 50&  ' ** It should only pop out of the emblem loop at certain intervals.
4240                Case False
                      ' ** Should be enabled if they're here!
4250                  blnGoingToReport = False
4260                  gblnGoToReport = False
4270                  blnGTR_Emblem = False
4280                  GTREmblem_Off  ' ** Procedure: Below.
4290                  Beep
4300                  DoCmd.Hourglass False
4310                  DoEvents
4320                End Select
4330              Case Else
                    ' ** Nothing else
4340                blnGoingToReport = False
4350                DoCmd.Hourglass False
4360                gblnGoToReport = False
4370                blnGTR_Emblem = False
4380                GTREmblem_Off  ' ** Procedure: Below.
4390                Beep
4400              End Select
4410            End Select
4420          End Select
4430        End If
4440      End If  ' ** gblnGoToReport.
4450      If blnHasForEx_Chk = True Then
4460        blnHasForEx_Chk = False
4470        If gblnGoToReport = False Then
4480          gblnHasForExThis = gblnHasForEx
4490          Select Case gblnHasForEx
              Case True
4500            Select Case .chkIncludeCurrency
                Case True
4510              .chkIncludeCurrency_lbl2.Visible = False
4520            Case False
4530              If Pref_HasPref(THIS_NAME, "chkIncludeCurrency") = False Then  ' ** Module Procedure: modPreferenceFuncs.
4540                varTmp00 = .frmAssetPricing_Sub.Form.ForExTot_Get  ' ** Form Function: frmAssetPricing_Sub.
4550                If IsNull(varTmp00) = True Then
4560                  .chkIncludeCurrency = False
4570                  .chkIncludeCurrency_lbl2.Visible = False
4580                Else
4590                  If varTmp00 = 0 Then
4600                    .chkIncludeCurrency = False
4610                    .chkIncludeCurrency_lbl2.Visible = False
4620                  Else
4630                    .chkIncludeCurrency = False
4640                    .chkIncludeCurrency_lbl2.Visible = True
4650                  End If
4660                End If
4670              Else
4680                If IsNull(varTmp00) = True Then
4690                  .chkIncludeCurrency_lbl2.Visible = False
4700                Else
4710                  If varTmp00 = 0 Then
4720                    .chkIncludeCurrency_lbl2.Visible = False
4730                  Else
4740                    .chkIncludeCurrency_lbl2.Visible = True
4750                  End If
4760                End If
4770              End If
4780            End Select
4790            .chkIncludeCurrency.Enabled = True
4800            chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
4810          Case False
4820            .chkIncludeCurrency = False
4830            .chkIncludeCurrency_lbl2.Visible = False
4840            chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
4850            .chkIncludeCurrency.Enabled = False
4860          End Select
4870          Select Case blnViewOnly
              Case True
4880            .cmdClose.SetFocus
4890          Case False
4900            .Date_New.SetFocus
4910          End Select
4920        End If
4930      End If  ' ** blnHasForEx_Chk.
4940      If blnForExTots = True Then
4950        blnForExTots = False
4960        If .chkIncludeCurrency = False Then
4970          .frmAssetPricing_Sub.Form.IsForEx_Tot.Requery
4980          DoEvents
4990          varTmp00 = .frmAssetPricing_Sub.Form.ForExTot_Get  ' ** Form Function: frmAssetPricing_Sub.
5000          If IsNull(varTmp00) = True Then
5010            .chkIncludeCurrency_lbl2.Visible = False
5020          Else
5030            If varTmp00 = 0 Then
5040              .chkIncludeCurrency_lbl2.Visible = False
5050            Else
5060              .chkIncludeCurrency_lbl2.Visible = True
5070            End If
5080          End If
5090        Else
5100          .chkIncludeCurrency_lbl2.Visible = False
5110        End If
5120        If blnGoingToReport = True Then
5130          .TimerInterval = (GTR_WAIT * 2&)
5140        End If
5150      End If  ' ** blnForExTots.
5160    End With

EXITP:
5170    Exit Sub

ERRH:
5180    DoCmd.Hourglass False
5190    blnGoingToReport = False
5200    gblnGoToReport = False
5210    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
5220    Case Else
5230      THAT_PROC = THIS_PROC
5240      That_Erl = Erl
5250      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
5260    End Select
5270    Resume EXITP

End Sub

Private Sub Form_Click()
' ** If a user put in a date, then clicked elsewhere, rather than hitting Enter/Return,
' ** SetPricing() will not have been activated. Since I don't want to activate it if
' ** they're just looking, I've removed the OnExit event that used to activate everything
' ** just by leaving the Date_New field. This way, they can open the form, then close
' ** it without setting anything off.
' ** If they're clicking around, trying to get into things, this event will trigger.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Click"

5310    SetPricing True  ' ** Function: Below.

EXITP:
5320    Exit Sub

ERRH:
5330    THAT_PROC = THIS_PROC
5340    That_Erl = Erl
5350    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
5360    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

5410    If DataErr <> 0 Then
5420      DoCmd.Hourglass False
5430      Select Case DataErr
          Case 13, 2108, 2116, 2169, 2501, 3020, 3101, 3314, 7753, 8519
            ' ** 2501: The '|' action was Canceled.
5440        Response = acDataErrContinue
5450      Case 2046  ' ** The command or action '|' isn't available now.
            ' ** Ignore.
5460        Response = acDataErrContinue
5470      Case 2110  ' ** Microsoft Access can't move the focus to the control |.
            ' ** Ignore.
5480        Response = acDataErrContinue
5490      Case 2113  ' ** The value you entered isn't valid for this field.
5500        Select Case Me.ActiveControl.Name
            Case "Date_New"
5510          If Me.Date_New.text <> "__/__/____" Then
5520            MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, (Left(("Invalid Date" & Space(55)), 55) & "FE01")
5530          Else
5540            Me.Date_New.text = vbNullString
5550          End If
5560          Response = acDataErrContinue
5570        End Select
5580      Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
            ' ** Referring to Date_New.Text.
5590        Response = acDataErrContinue
5600      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
5610        If Me.Date_New.text <> "__/__/____" Then
5620          MsgBox "Please enter a valid date (MM/DD/YYYY), Or" & vbCrLf & _
                "clear the contents of the date field to exit.", vbInformation + vbOKOnly, (Left(("Invalid Date" & Space(55)), 55) & "FE02")
5630        Else
5640          Me.Date_New.text = vbNullString
5650        End If
5660        Response = acDataErrContinue
5670      Case Else
5680        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
5690        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
5700      End Select
5710    End If

EXITP:
5720    Exit Sub

ERRH:
5730    DoCmd.Hourglass False
5740    Select Case ERR.Number
        Case Else
5750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5760    End Select
5770    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean

5810    DoCmd.Hourglass True
5820    DoEvents

        ' ** Check to make sure History has the current numbers.
5830    HistoryChk  ' ** Procedure: Below.

5840    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

5850    If Not clsMonthClass Is Nothing Then
5860      If clsMonthClass.IsCalendar Then
5870        Cancel = -1
5880      Else
5890        Set clsMonthClass = Nothing
5900      End If
5910    End If

5920    blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel

        ' ** Make sure the import form is closed.
5930    strDocName = "frmAssetPricing_Import"
5940    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
5950      DoCmd.Close acForm, strDocName
5960    End If

        ' ** Make sure the history form is closed.
5970    strDocName = "frmAssetPricing_History"
5980    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
5990      DoCmd.Close acForm, strDocName
6000    End If

        ' ** Make sure all reports are closed.
6010    Do While Reports.Count > 0
6020      DoCmd.Close acReport, Reports(0).Name
6030    Loop

6040    gstrReturningForm = THIS_NAME
6050    gblnSetFocus = True
6060    Select Case strCallingForm
        Case "frmMenu_Asset"
6070      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
6080    End Select

EXITP:
6090    Exit Sub

ERRH:
6100    DoCmd.Hourglass False
6110    Select Case ERR.Number
        Case Else
6120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6130    End Select
6140    Resume EXITP

End Sub

Private Sub FormHeader_Click()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_Click"

6210    Form_Click  ' ** Procedure: Above.

EXITP:
6220    Exit Sub

ERRH:
6230    THAT_PROC = THIS_PROC
6240    That_Erl = Erl
6250    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6260    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

6310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub Detail_Click()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Click"

6410    Form_Click  ' ** Procedure: Above.

EXITP:
6420    Exit Sub

ERRH:
6430    THAT_PROC = THIS_PROC
6440    That_Erl = Erl
6450    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6460    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

6510    With Me
6520      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
6530        Select Case blnPrintReport_Focus
            Case True
6540          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
6550          .cmdPrintReport_raised_img.Visible = False
6560        Case False
6570          .cmdPrintReport_raised_img.Visible = True
6580          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6590        End Select
6600        .cmdPrintReport_raised_focus_img.Visible = False
6610        .cmdPrintReport_raised_focus_dots_img.Visible = False
6620        .cmdPrintReport_sunken_focus_dots_img.Visible = False
6630        .cmdPrintReport_raised_img_dis.Visible = False
6640      End If
6650      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
6660        Select Case blnCalendar1_Focus
            Case True
6670          .cmdCalendar_raised_semifocus_dots_img.Visible = True
6680          .cmdCalendar_raised_img.Visible = False
6690        Case False
6700          .cmdCalendar_raised_img.Visible = True
6710          .cmdCalendar_raised_semifocus_dots_img.Visible = False
6720        End Select
6730        .cmdCalendar_raised_focus_img.Visible = False
6740        .cmdCalendar_raised_focus_dots_img.Visible = False
6750        .cmdCalendar_sunken_focus_dots_img.Visible = False
6760        .cmdCalendar_raised_img_dis.Visible = False
6770      End If
6780    End With

EXITP:
6790    Exit Sub

ERRH:
6800    Select Case ERR.Number
        Case Else
6810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6820    End Select
6830    Resume EXITP

End Sub

Private Sub FormFooter_Click()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_Click"

6910    Form_Click  ' ** Procedure: Above.

EXITP:
6920    Exit Sub

ERRH:
6930    THAT_PROC = THIS_PROC
6940    That_Erl = Erl
6950    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6960    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

7010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Public Sub cmdClose_Click()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim blnContinue As Boolean

7110    blnContinue = True

7120    With Me
7130      If .chkUpdated = True Then
7140        If MsgBox("You have made changes, but have not" & vbCrLf & _
                "confirmed them by clicking the Update button." & vbCrLf & vbCrLf & _
                "Are you sure you want to abandon changes?", vbQuestion + vbYesNo + vbDefaultButton2, _
                (Left(("Changes Not Saved" & Space(55)), 55) & "B01")) = vbNo Then
7150          blnContinue = False
7160        End If
7170      End If
7180    End With

7190    If blnContinue = True Then

7200      DoCmd.Close acForm, THIS_NAME

7210    End If

EXITP:
7220    Exit Sub

ERRH:
7230    THAT_PROC = THIS_PROC
7240    That_Erl = Erl
7250    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7260    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
7370        With Me
7380          intRetVal = 0
7390          lngRecsCur = .frmAssetPricing_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_Sub.
7400          If lngRecsCur > 0& Then
7410            .frmAssetPricing_Sub.SetFocus
7420            .frmAssetPricing_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Function: frmAssetPricing_Sub.
7430            If .frmAssetPricing_Sub.Form.marketvalue.Enabled = True Then
7440              .frmAssetPricing_Sub.Form.marketvalue.SetFocus
7450            Else
7460              .frmAssetPricing_Sub.Form.FocusHolder.SetFocus
7470            End If
7480          Else
7490            .opgShow.SetFocus
7500          End If
7510        End With
7520      End Select
7530    End If

        ' ** Shift keys.
7540    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7550      Select Case intRetVal
          Case vbKeyTab
7560        With Me
7570          intRetVal = 0
7580          If .cmdPrice.Enabled = True Then
7590            .cmdPrice.SetFocus
7600          Else
7610            .Date_New.SetFocus
7620          End If
7630        End With
7640      End Select
7650    End If

        ' ** Ctrl-Shift keys.
7660    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7670      Select Case intRetVal
          Case vbKeyTab
7680        With Me
7690          intRetVal = 0
7700          If .cmdPrice.Enabled = False Then
7710            .cmdCalendar.SetFocus
7720          Else
7730            Beep
7740          End If
7750        End With
7760      End Select
7770    End If

EXITP:
7780    KeyCode = intRetVal
7790    Exit Sub

ERRH:
7800    intRetVal = 0
7810    THAT_PROC = THIS_PROC
7820    That_Erl = Erl
7830    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7840    Resume EXITP

End Sub

Private Sub Date_Current_cmd_DblClick(Cancel As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_Current_cmd_DblClick"

7910    With Me
7920      .Date_New = .Date_Current
7930      .Date_New.SetFocus
7940    End With

EXITP:
7950    Exit Sub

ERRH:
7960    THAT_PROC = THIS_PROC
7970    That_Erl = Erl
7980    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7990    Resume EXITP

End Sub

Private Sub Date_New_BeforeUpdate(Cancel As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_New_BeforeUpdate"

8010    With Me
8020      If IsNull(.Date_New) = False Then
8030        If IsDate(.Date_New) = True Then
8040          If Compare_DateA_DateB(CDate(.Date_New), ">", Date) = True Then  ' ** Module Function: modStringFuncs.
8050            Cancel = -1
8060            MsgBox "Please enter a date no later than today." & vbCrLf & vbCrLf & _
                  "Future dates are not allowed.", vbInformation + vbOKOnly, (Left(("Invalid Date" & Space(55)), 55) & "C01")
8070          End If
8080        End If
8090      End If
8100    End With

EXITP:
8110    Exit Sub

ERRH:
8120    THAT_PROC = THIS_PROC
8130    That_Erl = Erl
8140    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8150    Resume EXITP

End Sub

Private Sub Date_New_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_New_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8210    intRetVal = KeyCode
8220    blnKeyDown_Return = False

        ' ** Use bit masks to determine which key was pressed.
8230    intShiftDown = (Shift And acShiftMask) > 0
8240    intAltDown = (Shift And acAltMask) > 0
8250    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8260    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8270      Select Case intRetVal
          Case vbKeyTab
8280        With Me
8290          intRetVal = 0
8300          If .cmdPrice.Enabled = True Then
8310            .cmdPrice.SetFocus
8320          Else
8330            .cmdClose.SetFocus
8340          End If
8350        End With
8360      Case vbKeyReturn
8370        blnKeyDown_Return = True
8380      End Select
8390    End If

        ' ** Shift keys.
8400    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8410      Select Case intRetVal
          Case vbKeyTab
8420        With Me
8430          intRetVal = 0
8440          If .cmdUpdate.Enabled = True Then
8450            .cmdUpdate.SetFocus
8460          ElseIf .cmdHistory.Enabled = True Then
8470            .cmdHistory.SetFocus
8480          ElseIf .cmdReset.Enabled = True Then
8490            .cmdReset.SetFocus
8500          ElseIf .cmdPrintReport.Enabled = True Then
8510            .cmdPrintReport.SetFocus
8520          ElseIf .chkShowHidden.Enabled = True Then
8530            .chkShowHidden.SetFocus
8540          ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
8550            .chkShowAssetNo.SetFocus
8560          ElseIf .chkIncludeCurrency.Enabled = True Then
8570            .chkIncludeCurrency.SetFocus
8580          Else
8590            .opgShow.SetFocus
8600          End If
8610        End With
8620      Case vbKeyReturn
8630        blnKeyDown_Return = True
8640      End Select
8650    End If

        ' ** Ctrl keys.
8660    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8670      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8680        With Me
8690          intRetVal = 0
8700          .cmdCalendar.SetFocus
8710        End With
8720      End Select
8730    End If

EXITP:
8740    KeyCode = intRetVal
8750    Exit Sub

ERRH:
8760    intRetVal = 0
8770    THAT_PROC = THIS_PROC
8780    That_Erl = Erl
8790    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8800    Resume EXITP

End Sub

Public Sub Date_New_Exit(Cancel As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_New_Exit"

8910    With Me
8920      If blnSetPricingInvoked = False Then
            ' ** If SetPricing() is being invoked via the above KeyDown() event, don't do anything else here.
            ' ** If, however, they've moved on via the mouse, then SetPricing() should be invoked.
8930        Select Case IsNull(.Date_New)
            Case True
8940          If blnKeyDown_Return = True Then
8950            blnKeyDown_Return = False
8960            .Date_New = Date
8970            SetPricing  ' ** Function: Below.
8980          Else
8990            If IsNull(.Date_New.text) = False Then
9000              If Trim(.Date_New.text) <> vbNullString Then
9010                SetPricing  ' ** Function: Below.
9020              Else
                    ' ** Empty, do nothing.
9030              End If
9040            Else
                  ' ** Empty, do nothing.
9050            End If
9060          End If
9070        Case False
9080          SetPricing  ' ** Function: Below.
9090        End Select
9100      End If
9110    End With

EXITP:
9120    Exit Sub

ERRH:
9130    THAT_PROC = THIS_PROC
9140    That_Erl = Erl
9150    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9160    Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

9210    With Me
9220      blnCalendar1_Focus = True
9230      .cmdCalendar_raised_semifocus_dots_img.Visible = True
9240      .cmdCalendar_raised_img.Visible = False
9250      .cmdCalendar_raised_focus_img.Visible = False
9260      .cmdCalendar_raised_focus_dots_img.Visible = False
9270      .cmdCalendar_sunken_focus_dots_img.Visible = False
9280      .cmdCalendar_raised_img_dis.Visible = False
9290    End With

EXITP:
9300    Exit Sub

ERRH:
9310    Select Case ERR.Number
        Case Else
9320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9330    End Select
9340    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

9410    With Me
9420      blnCalendar1_MouseDown = True
9430      .cmdCalendar_sunken_focus_dots_img.Visible = True
9440      .cmdCalendar_raised_img.Visible = False
9450      .cmdCalendar_raised_semifocus_dots_img.Visible = False
9460      .cmdCalendar_raised_focus_img.Visible = False
9470      .cmdCalendar_raised_focus_dots_img.Visible = False
9480      .cmdCalendar_raised_img_dis.Visible = False
9490    End With

EXITP:
9500    Exit Sub

ERRH:
9510    Select Case ERR.Number
        Case Else
9520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9530    End Select
9540    Resume EXITP

End Sub

Private Sub cmdCalendar_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim blnRetVal As Boolean

9610    With Me
9620      datStartDate = Date
9630      datEndDate = 0
9640      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
9650      If blnRetVal = True Then
            ' ** Don't allow future dates.
9660        If datStartDate > Date Then
9670          MsgBox "Please enter a date no later than today." & vbCrLf & vbCrLf & _
                "Future dates are not allowed.", vbInformation + vbOKOnly, (Left(("Invalid Date" & Space(55)), 55) & "C01")
9680        Else
9690          blnFromCalendar = True
9700          .Date_New = datStartDate
9710        End If
9720      Else
9730        blnFromCalendar = True
9740        .Date_New = Date
9750      End If
9760      .Date_New.SetFocus
9770    End With

EXITP:
9780    Exit Sub

ERRH:
9790    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
9800    Case Else
9810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9820    End Select
9830    Resume EXITP

End Sub

Private Sub cmdCalendar_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9910    intRetVal = KeyCode
9920    blnKeyDown_Return = False

        ' ** Use bit masks to determine which key was pressed.
9930    intShiftDown = (Shift And acShiftMask) > 0
9940    intAltDown = (Shift And acAltMask) > 0
9950    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9960    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9970      Select Case intRetVal
          Case vbKeyTab
9980        With Me
9990          intRetVal = 0
10000         If .cmdPrice.Enabled = True Then
10010           .cmdPrice.SetFocus
10020         Else
10030           .cmdClose.SetFocus
10040         End If
10050       End With
10060     End Select
10070   End If

        ' ** Shift keys.
10080   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10090     Select Case intRetVal
          Case vbKeyTab
10100       With Me
10110         intRetVal = 0
10120         .Date_New.SetFocus
10130       End With
10140     End Select
10150   End If

EXITP:
10160   KeyCode = intRetVal
10170   Exit Sub

ERRH:
10180   intRetVal = 0
10190   THAT_PROC = THIS_PROC
10200   That_Erl = Erl
10210   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10220   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

10310   With Me
10320     If blnCalendar1_MouseDown = False Then
10330       Select Case blnCalendar1_Focus
            Case True
10340         .cmdCalendar_raised_focus_dots_img.Visible = True
10350         .cmdCalendar_raised_focus_img.Visible = False
10360       Case False
10370         .cmdCalendar_raised_focus_img.Visible = True
10380         .cmdCalendar_raised_focus_dots_img.Visible = False
10390       End Select
10400       .cmdCalendar_raised_img.Visible = False
10410       .cmdCalendar_raised_semifocus_dots_img.Visible = False
10420       .cmdCalendar_sunken_focus_dots_img.Visible = False
10430       .cmdCalendar_raised_img_dis.Visible = False
10440     End If
10450   End With

EXITP:
10460   Exit Sub

ERRH:
10470   Select Case ERR.Number
        Case Else
10480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10490   End Select
10500   Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

10610   With Me
10620     .cmdCalendar_raised_focus_dots_img.Visible = True
10630     .cmdCalendar_raised_img.Visible = False
10640     .cmdCalendar_raised_semifocus_dots_img.Visible = False
10650     .cmdCalendar_raised_focus_img.Visible = False
10660     .cmdCalendar_sunken_focus_dots_img.Visible = False
10670     .cmdCalendar_raised_img_dis.Visible = False
10680     blnCalendar1_MouseDown = False
10690   End With

EXITP:
10700   Exit Sub

ERRH:
10710   Select Case ERR.Number
        Case Else
10720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10730   End Select
10740   Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

10810   With Me
10820     .cmdCalendar_raised_img.Visible = True
10830     .cmdCalendar_raised_semifocus_dots_img.Visible = False
10840     .cmdCalendar_raised_focus_img.Visible = False
10850     .cmdCalendar_raised_focus_dots_img.Visible = False
10860     .cmdCalendar_sunken_focus_dots_img.Visible = False
10870     .cmdCalendar_raised_img_dis.Visible = False
10880     blnCalendar1_Focus = False
10890   End With

EXITP:
10900   Exit Sub

ERRH:
10910   Select Case ERR.Number
        Case Else
10920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10930   End Select
10940   Resume EXITP

End Sub

Private Sub cmdPrice_Click()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrice_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

11010   With Me

11020     DoCmd.Hourglass True
11030     DoEvents

11040     Set dbs = CurrentDb
11050     With dbs
            ' ** Empty tblPricing_Staging.
11060       Set qdf = .QueryDefs("qryPricing_04e")
11070       qdf.Execute
11080       Set qdf = Nothing
            ' ** Append qryPricing_04c (qryPricing_04b (tblPricing_Import, all fields), linked to
            ' ** qryPricing_04a (tblForm_Graphics, just 'frmAssetPricing_Import')) to tblPricing_Staging.
11090       Set qdf = .QueryDefs("qryPricing_04d")
11100       qdf.Execute
11110       Set qdf = Nothing
11120       .Close
11130     End With
11140     Set dbs = Nothing

11150     gblnSetFocus = True
11160     strDocName = "frmAssetPricing_Import"
11170     DoCmd.OpenForm strDocName, , , , , , THIS_NAME & "~" & strCallingForm
11180     If gblnGoToReport = True Then
11190       .TimerInterval = 0&
11200       Forms(strDocName).TimerInterval = 50&
11210       blnGoingToReport = False
11220       blnGoingToReport2 = False
11230       blnGTR_Emblem = False
11240     End If

11250   End With

EXITP:
11260   Set qdf = Nothing
11270   Set dbs = Nothing
11280   Exit Sub

ERRH:
11290   DoCmd.Hourglass False
11300   THAT_PROC = THIS_PROC
11310   That_Erl = Erl
11320   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11330   Resume EXITP

End Sub

Private Sub cmdPrice_KeyDown(KeyCode As Integer, Shift As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrice_KeyDown"

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
11570         .Date_New.SetFocus
11580       End With
11590     End Select
11600   End If

        ' ** Ctrl-Shift keys.
11610   If intCtrlDown And (Not intAltDown) And intShiftDown Then
11620     Select Case intRetVal
          Case vbKeyTab
11630       With Me
11640         intRetVal = 0
11650         .cmdCalendar.SetFocus
11660       End With
11670     End Select
11680   End If

EXITP:
11690   KeyCode = intRetVal
11700   Exit Sub

ERRH:
11710   intRetVal = 0
11720   THAT_PROC = THIS_PROC
11730   That_Erl = Erl
11740   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11750   Resume EXITP

End Sub

Public Sub opgShow_AfterUpdate()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgShow_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim strTmp01 As String

11810   With Me
11820     Set frm = .frmAssetPricing_Sub.Form
11830     Select Case .opgShow
          Case .opgShow_optAll.OptionValue
11840       .opgShow_optAll_lbl.FontBold = True
11850       .opgShow_optHeldOnly_lbl.FontBold = False
11860       DoEvents
11870       If .frmAssetPricing_Sub.Form.RecordSource <> "qryPricing_01_02_04" Then
              ' ** tblPricing_MasterAsset, all assets.
11880         .frmAssetPricing_Sub.Form.RecordSource = "qryPricing_01_02_04"
11890       End If
11900       .frmAssetPricing_Sub.Form.totdesc_lbl.Caption = "Asset"
11910       Select Case .chkShowHidden
            Case True
11920         strFilter = vbNullString
11930         frm.Filter = strFilter
11940         frm.FilterOn = False
11950       Case False
11960         strFilter = "[IsHid] = False"
11970         frm.Filter = strFilter
11980         frm.FilterOn = True
11990       End Select
12000     Case .opgShow_optHeldOnly.OptionValue
12010       .opgShow_optAll_lbl.FontBold = False
12020       .opgShow_optHeldOnly_lbl.FontBold = True
12030       DoEvents
12040       If .frmAssetPricing_Sub.Form.RecordSource <> "qryPricing_01_01_04" Then
              ' ** tblPricing_MasterAsset, just shareface > 0.
12050         .frmAssetPricing_Sub.Form.RecordSource = "qryPricing_01_01_04"
12060       End If
12070       .frmAssetPricing_Sub.Form.totdesc_lbl.Caption = "Asset Owned"
12080       Select Case .chkShowHidden
            Case True
12090         strFilter = vbNullString
12100         frm.Filter = strFilter
12110         frm.FilterOn = False
12120       Case False
12130         strFilter = "[IsHid] = False"
12140         frm.Filter = strFilter
12150         frm.FilterOn = True
12160       End Select
12170     End Select
12180     DoEvents
12190     frm.Refresh
12200     chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
12210     strTmp01 = frm.SortNow_Get  ' ** Form Function: frmAssetPricing_Sub.
12220     frm.OrderBy = strTmp01
12230     frm.OrderByOn = True
          'If blnFromLoad = False Then
12240     lngRecsCur = .frmAssetPricing_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_Sub.
12250     If lngRecsCur > 0& Then
12260       .frmAssetPricing_Sub.SetFocus
12270       DoEvents
12280       blnForExTots = True
12290       .TimerInterval = 500&
12300     End If
          'Else
          '  blnFromLoad = False
          'End If
12310   End With

EXITP:
12320   Set frm = Nothing
12330   Exit Sub

ERRH:
12340   THAT_PROC = THIS_PROC
12350   That_Erl = Erl
12360   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
12370   Resume EXITP

End Sub

Private Sub opgShow_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgShow_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12420   intShiftDown = (Shift And acShiftMask) > 0
12430   intAltDown = (Shift And acAltMask) > 0
12440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12470       With Me
12480         intRetVal = 0
12490         If .chkIncludeCurrency.Enabled = True Then
12500           .chkIncludeCurrency.SetFocus
12510         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
12520           .chkShowAssetNo.SetFocus
12530         ElseIf .chkShowHidden.Enabled = True Then
12540           .chkShowHidden.SetFocus
12550         ElseIf .cmdPrintReport.Enabled = True Then
12560           .cmdPrintReport.SetFocus
12570         ElseIf .cmdReset.Enabled = True Then
12580           .cmdReset.SetFocus
12590         ElseIf .cmdHistory.Enabled = True Then
12600           .cmdHistory.SetFocus
12610         ElseIf .cmdUpdate.Enabled = True Then
12620           .cmdUpdate.SetFocus
12630         Else
12640           .Date_New.SetFocus
12650         End If
12660       End With
12670     End Select
12680   End If

        ' ** Shift keys.
12690   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12700     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12710       With Me
12720         intRetVal = 0
12730         lngRecsCur = .frmAssetPricing_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_Sub.
12740         If lngRecsCur > 0& Then
12750           .frmAssetPricing_Sub.SetFocus
12760           .frmAssetPricing_Sub.Form.FocusHolder.SetFocus
12770         Else
12780           .cmdClose.SetFocus
12790         End If
12800       End With
12810     End Select
12820   End If

EXITP:
12830   KeyCode = intRetVal
12840   Exit Sub

ERRH:
12850   intRetVal = 0
12860   THAT_PROC = THIS_PROC
12870   That_Erl = Erl
12880   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12890   Resume EXITP

End Sub

Private Sub opgShow_optHeldOnly_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgShow_optHeldOnly_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12920   intShiftDown = (Shift And acShiftMask) > 0
12930   intAltDown = (Shift And acAltMask) > 0
12940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12970       With Me
12980         intRetVal = 0
12990         If .chkIncludeCurrency.Enabled = True Then
13000           .chkIncludeCurrency.SetFocus
13010         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
13020           .chkShowAssetNo.SetFocus
13030         ElseIf .chkShowHidden.Enabled = True Then
13040           .chkShowHidden.SetFocus
13050         ElseIf .cmdPrintReport.Enabled = True Then
13060           .cmdPrintReport.SetFocus
13070         ElseIf .cmdReset.Enabled = True Then
13080           .cmdReset.SetFocus
13090         ElseIf .cmdHistory.Enabled = True Then
13100           .cmdHistory.SetFocus
13110         ElseIf .cmdUpdate.Enabled = True Then
13120           .cmdUpdate.SetFocus
13130         Else
13140           .Date_New.SetFocus
13150         End If
13160       End With
13170     End Select
13180   End If

        ' ** Shift keys.
13190   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13200     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13210       With Me
13220         intRetVal = 0
13230         lngRecsCur = .frmAssetPricing_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_Sub.
13240         If lngRecsCur > 0& Then
13250           .frmAssetPricing_Sub.SetFocus
13260           .frmAssetPricing_Sub.Form.FocusHolder.SetFocus
13270         Else
13280           .cmdClose.SetFocus
13290         End If
13300       End With
13310     End Select
13320   End If

EXITP:
13330   KeyCode = intRetVal
13340   Exit Sub

ERRH:
13350   intRetVal = 0
13360   THAT_PROC = THIS_PROC
13370   That_Erl = Erl
13380   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13390   Resume EXITP

End Sub

Public Sub chkIncludeCurrency_AfterUpdate()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

        Dim varTmp00 As Variant

13410   With Me
13420     DoCmd.Hourglass True
13430     DoEvents
13440     Select Case .chkIncludeCurrency
          Case True
13450       .chkIncludeCurrency_lbl.FontBold = True
13460       DoEvents
13470       If .frmAssetPricing_Sub.Form.curr_id.Visible = False Then
13480         .frmAssetPricing_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAssetPricing_Sub.
13490         DoEvents
13500         IncludeCurrency True  ' ** Procedure: Below.
13510       End If
13520     Case False
13530       .chkIncludeCurrency_lbl.FontBold = False
13540       DoEvents
13550 On Error Resume Next
13560       varTmp00 = .frmAssetPricing_Sub.Form.curr_id.Visible
13570       If ERR.Number <> 0 Then
13580 On Error GoTo ERRH
13590       Else
13600         If .frmAssetPricing_Sub.Form.curr_id.Visible = True Then
13610           .frmAssetPricing_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAssetPricing_Sub.
13620           DoEvents
13630           IncludeCurrency False  ' ** Procedure: Below.
13640         End If
13650       End If
13660     End Select
13670     DoEvents
13680 On Error Resume Next
13690     .frmAssetPricing_Sub.SetFocus
13700     .frmAssetPricing_Sub.Form.FocusHolder.SetFocus
13710 On Error GoTo ERRH
13720     Select Case .chkIncludeCurrency
          Case True
13730       .chkIncludeCurrency_lbl2.Visible = False
13740     Case False
13750 On Error Resume Next
13760       varTmp00 = .frmAssetPricing_Sub.Form.ForExTot_Get  ' ** Form Function: frmAssetPricing_Sub.
13770       If ERR.Number = 0 Then
13780 On Error GoTo ERRH
13790         If IsNull(varTmp00) = True Then
13800           .chkIncludeCurrency_lbl2.Visible = False
13810         Else
13820           If varTmp00 = 0 Then
13830             .chkIncludeCurrency_lbl2.Visible = False
13840           Else
13850             .chkIncludeCurrency_lbl2.Visible = True
13860           End If
13870         End If
13880       Else
13890 On Error GoTo ERRH
13900       End If
13910     End Select
13920     DoCmd.Hourglass False
13930   End With

EXITP:
13940   Exit Sub

ERRH:
13950   DoCmd.Hourglass False
13960   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
13970   Case Else
13980     THAT_PROC = THIS_PROC
13990     That_Erl = Erl
14000     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14010   End Select
14020   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14120   intShiftDown = (Shift And acShiftMask) > 0
14130   intAltDown = (Shift And acAltMask) > 0
14140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14170       With Me
14180         intRetVal = 0
14190         If .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
14200           .chkShowAssetNo.SetFocus
14210         ElseIf .chkShowHidden.Enabled = True Then
14220           .chkShowHidden.SetFocus
14230         ElseIf .cmdPrintReport.Enabled = True Then
14240           .cmdPrintReport.SetFocus
14250         ElseIf .cmdReset.Enabled = True Then
14260           .cmdReset.SetFocus
14270         ElseIf .cmdHistory.Enabled = True Then
14280           .cmdHistory.SetFocus
14290         ElseIf .cmdUpdate.Enabled = True Then
14300           .cmdUpdate.SetFocus
14310         Else
14320           .Date_New.SetFocus
14330         End If
14340       End With
14350     End Select
14360   End If

        ' ** Shift keys.
14370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14390       With Me
14400         intRetVal = 0
14410         .opgShow.SetFocus
14420       End With
14430     End Select
14440   End If

EXITP:
14450   KeyCode = intRetVal
14460   Exit Sub

ERRH:
14470   intRetVal = 0
14480   THAT_PROC = THIS_PROC
14490   That_Erl = Erl
14500   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14510   Resume EXITP

End Sub

Public Sub chkShowAssetNo_AfterUpdate()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetNo_AfterUpdate"

14610   With Me
14620     Select Case .chkShowAssetNo
          Case True
14630       .chkShowAssetNo_lbl.FontBold = True
14640       DoEvents
14650       .frmAssetPricing_Sub.Form.ShowAssetNo True  ' ** Form Procedure: frmAssetPricing_Sub.
14660     Case False
14670       .chkShowAssetNo_lbl.FontBold = False
14680       DoEvents
14690       .frmAssetPricing_Sub.Form.ShowAssetNo False  ' ** Form Procedure: frmAssetPricing_Sub.
14700     End Select
14710     DoEvents
14720 On Error Resume Next
14730     .frmAssetPricing_Sub.SetFocus
14740     .frmAssetPricing_Sub.Form.FocusHolder.SetFocus
14750 On Error GoTo ERRH
14760   End With

EXITP:
14770   Exit Sub

ERRH:
14780   THAT_PROC = THIS_PROC
14790   That_Erl = Erl
14800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14810   Resume EXITP

End Sub

Private Sub chkShowAssetNo_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14920   intShiftDown = (Shift And acShiftMask) > 0
14930   intAltDown = (Shift And acAltMask) > 0
14940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14970       With Me
14980         intRetVal = 0
14990         If .chkShowHidden.Enabled = True Then
15000           .chkShowHidden.SetFocus
15010         ElseIf .cmdPrintReport.Enabled = True Then
15020           .cmdPrintReport.SetFocus
15030         ElseIf .cmdReset.Enabled = True Then
15040           .cmdReset.SetFocus
15050         ElseIf .cmdHistory.Enabled = True Then
15060           .cmdHistory.SetFocus
15070         ElseIf .cmdUpdate.Enabled = True Then
15080           .cmdUpdate.SetFocus
15090         Else
15100           .Date_New.SetFocus
15110         End If
15120       End With
15130     End Select
15140   End If

        ' ** Shift keys.
15150   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15170       With Me
15180         intRetVal = 0
15190         If .chkIncludeCurrency.Enabled = True Then
15200           .chkIncludeCurrency.SetFocus
15210         Else
15220           .opgShow.SetFocus
15230         End If
15240       End With
15250     End Select
15260   End If

EXITP:
15270   KeyCode = intRetVal
15280   Exit Sub

ERRH:
15290   intRetVal = 0
15300   THAT_PROC = THIS_PROC
15310   That_Erl = Erl
15320   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15330   Resume EXITP

End Sub

Private Sub chkShowHidden_AfterUpdate()

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowHidden_AfterUpdate"

        Dim frm As Access.Form
        Dim strFilter As String
        Dim strTmp01 As String

15410   With Me
15420     Set frm = .frmAssetPricing_Sub.Form
15430     Select Case .chkShowHidden
          Case True
15440       DoEvents
15450       strFilter = vbNullString
15460       frm.Filter = strFilter
15470       frm.FilterOn = False
15480       .chkShowHidden_lbl.FontBold = True
15490     Case False
15500       DoEvents
15510       strFilter = "[IsHid] = False"
15520       frm.Filter = strFilter
15530       frm.FilterOn = True
15540       .chkShowHidden_lbl.FontBold = False
15550     End Select
15560     DoEvents
15570     frm.Refresh
15580     strTmp01 = frm.SortNow_Get  ' ** Form Function: frmAssets_Sub.
15590     frm.OrderBy = strTmp01
15600     frm.OrderByOn = True
15610 On Error Resume Next
15620     .frmAssetPricing_Sub.SetFocus
15630     .frmAssetPricing_Sub.Form.FocusHolder.SetFocus
15640 On Error GoTo ERRH
15650   End With

EXITP:
15660   Set frm = Nothing
15670   Exit Sub

ERRH:
15680   THAT_PROC = THIS_PROC
15690   That_Erl = Erl
15700   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15710   Resume EXITP

End Sub

Private Sub chkShowHidden_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowHidden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15820   intShiftDown = (Shift And acShiftMask) > 0
15830   intAltDown = (Shift And acAltMask) > 0
15840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15870       With Me
15880         intRetVal = 0
15890         If .cmdPrintReport.Enabled = True Then
15900           .cmdPrintReport.SetFocus
15910         ElseIf .cmdReset.Enabled = True Then
15920           .cmdReset.SetFocus
15930         ElseIf .cmdHistory.Enabled = True Then
15940           .cmdHistory.SetFocus
15950         ElseIf .cmdUpdate.Enabled = True Then
15960           .cmdUpdate.SetFocus
15970         Else
15980           .Date_New.SetFocus
15990         End If
16000       End With
16010     End Select
16020   End If

        ' ** Shift keys.
16030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16050       With Me
16060         intRetVal = 0
16070         If .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
16080           .chkShowAssetNo.SetFocus
16090         ElseIf .chkIncludeCurrency.Enabled = True Then
16100           .chkIncludeCurrency.SetFocus
16110         Else
16120           .opgShow.SetFocus
16130         End If
16140       End With
16150     End Select
16160   End If

EXITP:
16170   KeyCode = intRetVal
16180   Exit Sub

ERRH:
16190   intRetVal = 0
16200   THAT_PROC = THIS_PROC
16210   That_Erl = Erl
16220   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16230   Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

16310   With Me
16320     blnPrintReport_Focus = True
16330     .cmdPrintReport_raised_semifocus_dots_img.Visible = True
16340     .cmdPrintReport_raised_img.Visible = False
16350     .cmdPrintReport_raised_focus_img.Visible = False
16360     .cmdPrintReport_raised_focus_dots_img.Visible = False
16370     .cmdPrintReport_sunken_focus_dots_img.Visible = False
16380     .cmdPrintReport_raised_img_dis.Visible = False
16390   End With

EXITP:
16400   Exit Sub

ERRH:
16410   Select Case ERR.Number
        Case Else
16420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16430   End Select
16440   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

16510   With Me
16520     blnPrintReport_MouseDown = True
16530     .cmdPrintReport_sunken_focus_dots_img.Visible = True
16540     .cmdPrintReport_raised_img.Visible = False
16550     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
16560     .cmdPrintReport_raised_focus_img.Visible = False
16570     .cmdPrintReport_raised_focus_dots_img.Visible = False
16580     .cmdPrintReport_raised_img_dis.Visible = False
16590   End With

EXITP:
16600   Exit Sub

ERRH:
16610   Select Case ERR.Number
        Case Else
16620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16630   End Select
16640   Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

16710   With Me
16720     Select Case .chkIncludeCurrency
          Case True
16730       strDocName = "rptAssetPricing_ForEx"
16740     Case False
16750       strDocName = "rptAssetPricing"
16760     End Select
16770     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
16780       DoCmd.OpenReport strDocName, acViewPreview
16790       DoCmd.Maximize
16800       DoCmd.RunCommand acCmdFitToWindow
16810       .Visible = False
16820     Else
            '##GTR_Ref: rptAssetPricing
            '##GTR_Ref: rptAssetPricing_ForEx
16830       DoCmd.OpenReport strDocName, acViewNormal
16840     End If
16850   End With

EXITP:
16860   Exit Sub

ERRH:
16870   DoCmd.Restore
16880   Me.Visible = True
16890   THAT_PROC = THIS_PROC
16900   That_Erl = Erl
16910   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16920   Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17020   intShiftDown = (Shift And acShiftMask) > 0
17030   intAltDown = (Shift And acAltMask) > 0
17040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17060     Select Case intRetVal
          Case vbKeyTab
17070       With Me
17080         intRetVal = 0
17090         If .cmdReset.Enabled = True Then
17100           .cmdReset.SetFocus
17110         ElseIf .cmdHistory.Enabled = True Then
17120           .cmdHistory.SetFocus
17130         ElseIf .cmdUpdate.Enabled = True Then
17140           .cmdUpdate.SetFocus
17150         Else
17160           .Date_New.SetFocus
17170         End If
17180       End With
17190     End Select
17200   End If

        ' ** Shift keys.
17210   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17220     Select Case intRetVal
          Case vbKeyTab
17230       With Me
17240         intRetVal = 0
17250         If .chkShowHidden.Enabled = True Then
17260           .chkShowHidden.SetFocus
17270         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
17280           .chkShowAssetNo.SetFocus
17290         ElseIf .chkIncludeCurrency.Enabled = True Then
17300           .chkIncludeCurrency.SetFocus
17310         Else
17320           .opgShow.SetFocus
17330         End If
17340       End With
17350     End Select
17360   End If

EXITP:
17370   KeyCode = intRetVal
17380   Exit Sub

ERRH:
17390   intRetVal = 0
17400   THAT_PROC = THIS_PROC
17410   That_Erl = Erl
17420   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17430   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

17510   With Me
17520     If blnPrintReport_MouseDown = False Then
17530       Select Case blnPrintReport_Focus
            Case True
17540         .cmdPrintReport_raised_focus_dots_img.Visible = True
17550         .cmdPrintReport_raised_focus_img.Visible = False
17560       Case False
17570         .cmdPrintReport_raised_focus_img.Visible = True
17580         .cmdPrintReport_raised_focus_dots_img.Visible = False
17590       End Select
17600       .cmdPrintReport_raised_img.Visible = False
17610       .cmdPrintReport_raised_semifocus_dots_img.Visible = False
17620       .cmdPrintReport_sunken_focus_dots_img.Visible = False
17630       .cmdPrintReport_raised_img_dis.Visible = False
17640     End If
17650   End With

EXITP:
17660   Exit Sub

ERRH:
17670   Select Case ERR.Number
        Case Else
17680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17690   End Select
17700   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

17810   With Me
17820     .cmdPrintReport_raised_focus_dots_img.Visible = True
17830     .cmdPrintReport_raised_img.Visible = False
17840     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
17850     .cmdPrintReport_raised_focus_img.Visible = False
17860     .cmdPrintReport_sunken_focus_dots_img.Visible = False
17870     .cmdPrintReport_raised_img_dis.Visible = False
17880     blnPrintReport_MouseDown = False
17890   End With

EXITP:
17900   Exit Sub

ERRH:
17910   Select Case ERR.Number
        Case Else
17920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17930   End Select
17940   Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

18010   With Me
18020     .cmdPrintReport_raised_img.Visible = True
18030     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
18040     .cmdPrintReport_raised_focus_img.Visible = False
18050     .cmdPrintReport_raised_focus_dots_img.Visible = False
18060     .cmdPrintReport_sunken_focus_dots_img.Visible = False
18070     .cmdPrintReport_raised_img_dis.Visible = False
18080     blnPrintReport_Focus = False
18090   End With

EXITP:
18100   Exit Sub

ERRH:
18110   Select Case ERR.Number
        Case Else
18120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18130   End Select
18140   Resume EXITP

End Sub

Private Sub cmdReset_Click()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim msgResponse As VbMsgBoxResult

18210   With Me
18220     If .chkUpdated = True Or .cmdUpdate.Enabled = True Then
18230       msgResponse = MsgBox("This will restore the screen to the existing, current Asset Pricing values." & vbCrLf & _
              "Any changes you have made will be abandoned." & vbCrLf & vbCrLf & "Proceed?", vbQuestion + vbYesNo, "Abandon Changes Made")
18240       If msgResponse = vbYes Then

18250         DoCmd.Hourglass True
18260         DoEvents

18270         gblnMessage = True

18280         .frmAssetPricing_Sub.SetFocus
18290         .frmAssetPricing_Sub.Form.FocusHolder2.SetFocus
18300         DoEvents

18310         Set dbs = CurrentDb
18320         With dbs
                ' ** Empty tblPricing_MasterAsset.
18330           Set qdf = .QueryDefs("qryPricing_05")
18340           qdf.Execute
                ' ** Empty tblPricing_Cusip.
18350           Set qdf = .QueryDefs("qryPricing_06")
18360           qdf.Execute
                ' ** Append MasterAsset to tblPricing_MasterAsset.
18370           Set qdf = .QueryDefs("qryPricing_07")
18380           qdf.Execute
18390           .Close
18400         End With

18410         .frmAssetPricing_Sub.Form.Requery
18420         DoEvents

18430         .Date_Current = .frmAssetPricing_Sub.Form.currentDate

18440         .Date_New.SetFocus
18450         .chkUpdated = False
18460         .cmdReset.Enabled = False
18470         .cmdUpdate.Enabled = False

18480         gblnMessage = False

18490         DoCmd.Hourglass False

18500       End If
18510     Else
18520       MsgBox "No changes have been made since the last update.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
18530     End If
18540   End With

EXITP:
18550   Set qdf = Nothing
18560   Set dbs = Nothing
18570   Exit Sub

ERRH:
18580   DoCmd.Hourglass False
18590   THAT_PROC = THIS_PROC
18600   That_Erl = Erl
18610   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18620   Resume EXITP

End Sub

Private Sub cmdReset_KeyDown(KeyCode As Integer, Shift As Integer)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18720   intShiftDown = (Shift And acShiftMask) > 0
18730   intAltDown = (Shift And acAltMask) > 0
18740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18760     Select Case intRetVal
          Case vbKeyTab
18770       With Me
18780         intRetVal = 0
18790         If .cmdHistory.Enabled = True Then
18800           .cmdHistory.SetFocus
18810         ElseIf .cmdUpdate.Enabled = True Then
18820           .cmdUpdate.SetFocus
18830         Else
18840           .Date_New.SetFocus
18850         End If
18860       End With
18870     End Select
18880   End If

        ' ** Shift keys.
18890   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18900     Select Case intRetVal
          Case vbKeyTab
18910       With Me
18920         intRetVal = 0
18930         If .cmdPrintReport.Enabled = True Then
18940           .cmdPrintReport.SetFocus
18950         ElseIf .chkShowHidden.Enabled = True Then
18960           .chkShowHidden.SetFocus
18970         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
18980           .chkShowAssetNo.SetFocus
18990         ElseIf .chkIncludeCurrency.Enabled = True Then
19000           .chkIncludeCurrency.SetFocus
19010         Else
19020           .opgShow.SetFocus
19030         End If
19040       End With
19050     End Select
19060   End If

EXITP:
19070   KeyCode = intRetVal
19080   Exit Sub

ERRH:
19090   intRetVal = 0
19100   THAT_PROC = THIS_PROC
19110   That_Erl = Erl
19120   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19130   Resume EXITP

End Sub

Private Sub cmdHistory_Click()

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistory_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

19210   With Me
19220     Select Case .cmdHistory.Enabled
          Case True
19230       DoCmd.Hourglass True
19240       DoEvents

19250       Set dbs = CurrentDb
19260       With dbs
              ' ** Empty tblMark.
19270         Set qdf = .QueryDefs("qryPricing_History_06a")
19280         qdf.Execute
19290         Set qdf = Nothing
              ' ** Append tblPricing_MasterAsset_History to tblMark.
19300         Set qdf = .QueryDefs("qryPricing_History_06b")
19310         qdf.Execute
19320         Set qdf = Nothing
19330         .Close
19340       End With
19350       Set dbs = Nothing
19360       DoEvents

19370       gblnSetFocus = True
19380       strDocName = "frmAssetPricing_History"
19390       DoCmd.OpenForm strDocName, , , , , , THIS_NAME
19400       If gblnGoToReport = True Then
19410         .TimerInterval = 0&
19420         Forms(strDocName).TimerInterval = 50&
19430         blnGoingToReport = False
19440         blnGoingToReport2 = False
19450         blnGTR_Emblem = False
19460         GTREmblem_Off  ' ** Procedure: Below.
19470         .GoToReport_arw_hist_img.Visible = False
19480         DoEvents
19490       End If
19500     Case False
19510       If gblnGoToReport = True Then
19520         .TimerInterval = 0&
19530         DoCmd.Hourglass False
19540         gblnGoToReport = False
19550         blnGoingToReport = False
19560         blnGoingToReport2 = False
19570         blnGTR_Emblem = False
19580         GTREmblem_Off  ' ** Procedure: Below.
19590         .GoToReport_arw_hist_img.Visible = False
19600       End If
19610       Beep
19620     End Select
19630   End With

EXITP:
19640   Set qdf = Nothing
19650   Set dbs = Nothing
19660   Exit Sub

ERRH:
19670   DoCmd.Hourglass False
19680   THAT_PROC = THIS_PROC
19690   That_Erl = Erl
19700   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19710   Resume EXITP

End Sub

Private Sub cmdHistory_KeyDown(KeyCode As Integer, Shift As Integer)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistory_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19820   intShiftDown = (Shift And acShiftMask) > 0
19830   intAltDown = (Shift And acAltMask) > 0
19840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19860     Select Case intRetVal
          Case vbKeyTab
19870       With Me
19880         intRetVal = 0
19890         If .cmdUpdate.Enabled = True Then
19900           .cmdUpdate.SetFocus
19910         Else
19920           .Date_New.SetFocus
19930         End If
19940       End With
19950     End Select
19960   End If

        ' ** Shift keys.
19970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19980     Select Case intRetVal
          Case vbKeyTab
19990       With Me
20000         intRetVal = 0
20010         If .cmdReset.Enabled = True Then
20020           .cmdReset.SetFocus
20030         ElseIf .cmdPrintReport.Enabled = True Then
20040           .cmdPrintReport.SetFocus
20050         ElseIf .chkShowHidden.Enabled = True Then
20060           .chkShowHidden.SetFocus
20070         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
20080           .chkShowAssetNo.SetFocus
20090         ElseIf .chkIncludeCurrency.Enabled = True Then
20100           .chkIncludeCurrency.SetFocus
20110         Else
20120           .opgShow.SetFocus
20130         End If
20140       End With
20150     End Select
20160   End If

EXITP:
20170   KeyCode = intRetVal
20180   Exit Sub

ERRH:
20190   intRetVal = 0
20200   THAT_PROC = THIS_PROC
20210   That_Erl = Erl
20220   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20230   Resume EXITP

End Sub

Private Sub cmdUpdate_Click()

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim usr As DAO.User, grp As DAO.Group, frm As Access.Form
        Dim datNewDate As Date
        Dim lngAssetNo As Long, dblYield As Double, dblMktVal As Double
        Dim lngChanges As Long, arr_varChange() As Variant
        Dim lngRecs As Long
        Dim intMsgNum As Integer
        Dim blnContinue As Boolean, blnFound As Boolean, blnIncludeCurrency As Boolean
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

        ' ** Array: arr_varChange().
        Const C_ELEMS As Integer = 3
        Const C_ASTNO  As Integer = 0
        Const C_YIELD  As Integer = 1
        Const C_MKTVAL As Integer = 2
        Const C_CURDAT As Integer = 3

20310   blnContinue = True
20320   blnFound = False
20330   intMsgNum = 0
20340   strTmp01 = vbNullString

20350   For Each grp In DBEngine.Workspaces(0).Groups
20360     If grp.Name = "ViewOnly" Then
20370       For Each usr In grp.Users
20380         If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
20390           blnFound = True
20400           Exit For
20410         End If
20420       Next
20430     End If
20440   Next

20450   With Me

20460     If blnFound = True Then
20470       blnContinue = False
20480       MsgBox gstrNoPermission, vbInformation + vbOKOnly, (Left(("Access Denied" & Space(55)), 55) & "E01")
20490     Else

20500       Set frm = .frmAssetPricing_Sub.Form

            ' ** Assure the record they were on is saved.
20510       frm.cmdSave_Click  ' ** Form Procedure: frmAssetPricing_Sub.

20520       If IsNull(.Date_New) = True Then
20530         .Date_New.SetFocus
20540         If IsNull(.Date_New.text) = True Then
20550           blnContinue = False
20560           intMsgNum = 1
20570         Else
20580           If Trim(.Date_New.text) = vbNullString Then
20590             blnContinue = False
20600             intMsgNum = 1
20610           Else
20620             If IsDate(.Date_New.text) = False Then
20630               blnContinue = False
20640               intMsgNum = 2
20650             Else
20660               datNewDate = .Date_New.text
20670             End If
20680           End If
20690         End If
20700       Else
20710         If Trim(.Date_New) = vbNullString Then
20720           .Date_New.SetFocus
20730           If IsNull(.Date_New.text) = True Then
20740             blnContinue = False
20750             intMsgNum = 1
20760           Else
20770             If Trim(.Date_New.text) = vbNullString Then
20780               blnContinue = False
20790               intMsgNum = 1
20800             Else
20810               If IsDate(.Date_New.text) = False Then
20820                 blnContinue = False
20830                 intMsgNum = 2
20840               Else
20850                 datNewDate = .Date_New.text
20860               End If
20870             End If
20880           End If
20890         Else
20900           If IsDate(.Date_New) = False Then
20910             blnContinue = False
20920             intMsgNum = 2
20930           Else
20940             datNewDate = .Date_New
20950           End If
20960         End If
20970       End If

20980       lngChanges = 0&
20990       ReDim arr_varChange(C_ELEMS, 0)

21000       If blnContinue = True Then

21010         If frm.yield_entry.Enabled = True And frm.marketvalue.Enabled = True Then

21020           DoCmd.Hourglass True
21030           DoEvents

                ' ****************************************************
                ' ** This is where MasterAsset finally gets updated!
                ' ****************************************************

                ' ** Move the subform focus (if they edited the first record, we'll do it again below)
21040           frm.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAssetPricing_Sub.
21050           Set dbs = CurrentDb
21060           Set rst1 = frm.RecordsetClone
21070           With rst1
21080             .MoveLast
21090             lngRecs = .RecordCount
21100             .MoveFirst

21110             For lngX = 1& To lngRecs
21120               If ![AssetPricing_Changed] = True Or IsNull(![marketvalue]) = False Then
21130                 lngAssetNo = ![assetno]
21140                 If IsNull(![yield_entry]) = False And IsNull(![yield]) = False Then
21150                   If Round((![yield_entry] / 100#), 6) = Round(![yield], 6) Then
21160                     dblYield = ![yield]
21170                   Else
21180                     dblYield = Round((![yield_entry] / 100#), 6)  ' ** Yield should've been updated!
21190                   End If
21200                 Else
21210                   If IsNull(![yield]) = False Then
21220                     dblYield = ![yield]
21230                   ElseIf IsNull(![yield_entry]) = False Then
21240                     dblYield = Round((![yield_entry] / 100#), 6)
21250                   Else
21260                     dblYield = 0#
21270                   End If
21280                 End If
21290                 If IsNull(![marketvalue]) = False Then
21300                   dblMktVal = ![marketvalue]
21310                 ElseIf IsNull(![marketvaluecurrent]) = False Then
21320                   dblMktVal = ![marketvaluecurrent]
21330                 Else
21340                   dblMktVal = 0#
21350                 End If
                      ' ** MasterAsset, by specified [astno].
21360                 Set qdf = dbs.QueryDefs("qryPricing_11")
21370                 With qdf.Parameters
21380                   ![astno] = lngAssetNo
21390                 End With
21400                 Set rst2 = qdf.OpenRecordset
21410                 With rst2
21420                   .Edit
21430                   ![yield] = dblYield
21440                   ![marketvaluecurrent] = dblMktVal
21450                   ![marketvalue] = Null
21460                   ![currentDate] = datNewDate
21470                   .Update
21480                   .Close
21490                 End With  ' ** rst2.
21500                 Set rst2 = Nothing
21510                 Set qdf = Nothing
21520                 lngChanges = lngChanges + 1&
21530                 lngE = lngChanges - 1&
21540                 ReDim Preserve arr_varChange(C_ELEMS, lngE)
21550                 arr_varChange(C_ASTNO, lngE) = lngAssetNo
21560                 arr_varChange(C_YIELD, lngE) = dblYield
21570                 arr_varChange(C_MKTVAL, lngE) = dblMktVal
21580                 arr_varChange(C_CURDAT, lngE) = datNewDate
21590               End If
21600               If lngX < lngRecs Then .MoveNext
21610             Next

                  ' ** Just to be sure.
21620             frm.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAssetPricing_Sub.
21630             .MoveFirst

21640             For lngX = 1 To lngRecs
21650               .Edit
21660               ![marketvalue] = Null
21670               .Update
21680               If lngX < lngRecs Then .MoveNext
21690             Next
21700             .MoveFirst

                  ' ** Make sure this is updated before saving it to tblPricing_MasterAsset_History.
21710             For lngX = 0& To (lngChanges - 1&)
21720               .FindFirst "[assetno] = " & CStr(arr_varChange(C_ASTNO, lngX))
21730               If .NoMatch = False Then
21740                 .Edit
21750                 ![yield] = arr_varChange(C_YIELD, lngX)
21760                 ![marketvaluecurrent] = arr_varChange(C_MKTVAL, lngX)
21770                 ![currentDate] = arr_varChange(C_CURDAT, lngX)
21780                 .Update
21790               End If
21800             Next

21810             .Close
21820           End With  ' ** rst1.
21830           Set rst1 = Nothing

                ' ** If it's coming from a Restore, yield_entry and marketvalue will already be locked.
21840           With frm
21850             .FocusHolder.SetFocus
21860             .yield_entry.Enabled = False
21870             .yield_entry.Locked = True
21880             .yield_box.BackColor = CLR_DISABLED_BG
21890             .yield_entry.ForeColor = CLR_DISABLED_FG
21900             .marketvalue.Enabled = False
21910             .marketvalue.Locked = True
21920             .marketvalue_box.BackColor = CLR_DISABLED_BG
21930             .marketvalue_loc_usd_box.ForeColor = CLR_DISABLED_BG
21940             .currsym_symbol_loc_usd_box.ForeColor = CLR_DISABLED_BG
21950             .marketvalue.ForeColor = CLR_DISABLED_FG
21960           End With

                ' ** Update tblPricing_MasterAsset, by specified [curdat].
21970           Set qdf = dbs.QueryDefs("qryPricing_12")
21980           With qdf.Parameters
21990             ![curdat] = datNewDate
22000           End With
22010           qdf.Execute

                ' ** Update MasterAsset, by specified [curdat].
22020           Set qdf = dbs.QueryDefs("qryPricing_13")
22030           With qdf.Parameters
22040             ![curdat] = datNewDate
22050           End With
22060           qdf.Execute

22070           frm.Requery
22080           DoEvents

22090           Select Case .chkRestored
                Case True
                  ' ** Don't replace the original.
22100             .chkRestored = False
22110           Case False

                  ' ** Delete pre-existing tblPricing_MasterAsset_History records for this CurrentDate.
22120             Set qdf = dbs.QueryDefs("qryPricing_History_03")
22130             qdf.Execute

                  ' ** Append qryPricing_History_04 (tblPricing_MasterAsset, with add'l fields) to tblPricing_MasterAsset_History.
22140             Set qdf = dbs.QueryDefs("qryPricing_History_05")
22150             With qdf.Parameters
22160               ![datnow] = Now()
22170             End With
22180             qdf.Execute dbFailOnError

22190           End Select

22200           blnIncludeCurrency = .chkIncludeCurrency
22210           With frm
22220             .yield_upd.Visible = False  ' ** Set visible by frmAssetPricing_Import.
22230             .marketvalue_upd.Visible = False
22240             If blnIncludeCurrency = True Then
22250               .marketvalue_usd_upd.Visible = False
22260             End If
                  ' ** Make these invisible while tblPricing_MasterAsset is repopulated.
22270             .cusip.Visible = False
22280             .totdesc.Visible = False
22290             .yield_entry.Visible = False
22300             .marketvaluecurrent_str.Visible = False
22310             .marketvalue.Visible = False
22320           End With

                ' ** Empty tblPricing_MasterAsset.
22330           Set qdf = dbs.QueryDefs("qryPricing_05")
22340 On Error Resume Next
22350           qdf.Execute dbFailOnError
22360           If ERR.Number <> 0 Then
22370 On Error GoTo ERRH
22380             strTmp01 = frm.RecordSource
22390             frm.RecordSource = vbNullString
22400             DoEvents
22410             qdf.Execute dbFailOnError
22420           Else
22430 On Error GoTo ERRH
22440           End If

                ' ** Empty tblPricing_Cusip.
22450           Set qdf = dbs.QueryDefs("qryPricing_06")
22460           qdf.Execute dbFailOnError

                ' ** Append MasterAsset to tblPricing_MasterAsset.
22470           Set qdf = dbs.QueryDefs("qryPricing_07")
22480           qdf.Execute dbFailOnError

                ' ** MasterAsset, grouped, with Max(currentDate).
22490           Set qdf = dbs.QueryDefs("qryPricing_18")
22500           Set rst1 = qdf.OpenRecordset
22510           With rst1
22520             If .BOF = True And .EOF = True Then
                    ' ** Something's seriously wrong!
22530             Else
22540               .MoveFirst
22550               datNewDate = ![currentDate]
22560             End If
22570             .Close
22580           End With

22590           dbs.Close

22600           .Requery
22610           DoEvents

22620           With frm
22630             .cusip.Visible = True
22640             .totdesc.Visible = True
22650             .yield_entry.Visible = True
22660             .marketvaluecurrent_str.Visible = True
22670             .marketvalue.Visible = True
22680             If strTmp01 <> vbNullString Then
22690               .RecordSource = strTmp01
22700             End If
22710           End With

22720           .Date_Current = datNewDate
22730           .Date_New = Null
22740           DoEvents

22750           frm.Requery
22760           frm.Refresh
22770           DoEvents

                ' ** tblPricing_MasterAsset, grouped, for marketvalue <> Null, with cnt.
22780           varTmp00 = DLookup("[cnt]", "qryPricing_19")
22790           If IsNull(varTmp00) = False Then
22800             If varTmp00 > 0 Then
                    'MsgBox "The MarketValue field isn't empty, and it should be!", vbCritical + vbOKOnly, "SNAFU"
22810             End If
22820           End If

22830           DoCmd.Hourglass False
22840           DoEvents

22850           Beep
22860           MsgBox "Finished updating assets.", vbInformation + vbOKOnly, _
                  (Left(("Update Successful" & Space(40)), 40) & "E02")

22870         End If  ' ** Yield enabled.

22880         .cmdClose.SetFocus
22890         .cmdPrice.Enabled = False
22900         .cmdUpdate.Enabled = False
22910         .chkUpdated = False

22920       Else
22930         Select Case intMsgNum
              Case 1
22940           MsgBox "Please enter a date to continue.", _
                  vbInformation + vbOKOnly, (Left(("Entry Required" & Space(55)), 55) & "E03")
22950         Case 2
22960           MsgBox "Please enter a valid date (MM/DD/YYYY).", _
                  vbInformation + vbOKOnly, (Left(("Invalid Date" & Space(55)), 55) & "E04")
22970         Case Else
                ' ** Shouldn't be an else.
22980         End Select
22990       End If  ' ** blnContinue.

23000     End If  ' ** blnFound.

23010   End With  ' ** Me.

EXITP:
23020   Set frm = Nothing
23030   Set usr = Nothing
23040   Set grp = Nothing
23050   Set rst1 = Nothing
23060   Set rst2 = Nothing
23070   Set qdf = Nothing
23080   Set dbs = Nothing
23090   Exit Sub

ERRH:
23100   DoCmd.Hourglass False
23110   THAT_PROC = THIS_PROC
23120   That_Erl = Erl
23130   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23140   Resume EXITP

End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdUpdate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23220   intShiftDown = (Shift And acShiftMask) > 0
23230   intAltDown = (Shift And acAltMask) > 0
23240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23260     Select Case intRetVal
          Case vbKeyTab
23270       With Me
23280         intRetVal = 0
23290         .Date_New.SetFocus
23300       End With
23310     End Select
23320   End If

        ' ** Shift keys.
23330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23340     Select Case intRetVal
          Case vbKeyTab
23350       With Me
23360         intRetVal = 0
23370         If .cmdHistory.Enabled = True Then
23380           .cmdHistory.SetFocus
23390         ElseIf .cmdReset.Enabled = True Then
23400           .cmdReset.SetFocus
23410         ElseIf .cmdPrintReport.Enabled = True Then
23420           .cmdPrintReport.SetFocus
23430         ElseIf .chkShowHidden.Enabled = True Then
23440           .chkShowHidden.SetFocus
23450         ElseIf .chkIncludeCurrency.Enabled = True Then
23460           .chkIncludeCurrency.SetFocus
23470         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
23480           .chkShowAssetNo.SetFocus
23490         Else
23500           .opgShow.SetFocus
23510         End If
23520       End With
23530     End Select
23540   End If

EXITP:
23550   KeyCode = intRetVal
23560   Exit Sub

ERRH:
23570   intRetVal = 0
23580   THAT_PROC = THIS_PROC
23590   That_Erl = Erl
23600   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23610   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

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

Public Function SetPricing(Optional varEnableOnly As Variant) As Boolean

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "SetPricing"

        Dim blnEnableOnly As Boolean, blnSetDefault As Boolean
        Dim datNewDate As Date
        Dim strControlName As String
        Dim blnRetVal As Boolean

23810   blnRetVal = True
23820   blnSetDefault = False

23830   If IsMissing(varEnableOnly) = True Then
23840     blnEnableOnly = False
23850   Else
          ' ** This allows checking for mere presense of date, as well as pricing permission.
23860     blnEnableOnly = CBool(varEnableOnly)
23870   End If

23880 On Error Resume Next
23890   strControlName = Screen.ActiveControl.Name  ' ** If there is no active control, this errors.
23900 On Error GoTo ERRH

23910   With Me

23920     If blnEnableOnly = False Then
23930       If IsNull(.Date_New) = False Then
23940         If IsDate(.Date_New) = True Then
                ' ** Looks OK.
23950           datNewDate = .Date_New
23960         Else
23970           If Trim(.Date_New) <> vbNullString Then
23980             blnRetVal = False
23990             MsgBox "Please enter a valid date (MM/DD/YYYY).", _
                    vbInformation + vbOKOnly, (Left(("Invalid Date" & Space(55)), 55) & "F01")
24000           Else
24010             blnSetDefault = True
24020           End If
24030         End If
24040       Else
24050         .Date_New.SetFocus
24060         If IsNull(.Date_New.text) = False Then
24070           If Trim(.Date_New.text) <> vbNullString Then
24080             If IsDate(.Date_New.text) = True Then
                    ' ** Looks OK.
24090               datNewDate = .Date_New.text
24100             Else
24110               If .Date_New.text <> "__/__/____" Then
                      ' ** It might be right, but just a 2-digit year or something.
24120               Else
24130                 .Date_New.text = vbNullString
24140                 blnSetDefault = True
24150               End If
24160             End If
24170           Else
24180             blnSetDefault = True
24190           End If
24200         Else
24210           blnSetDefault = True
24220         End If
24230       End If
24240       If blnSetDefault = True Then
24250         .Date_New = Date
24260         datNewDate = Date
24270       End If
24280     Else
24290       If IsNull(.Date_New) = False Then
24300         If IsDate(.Date_New) = True Then
                ' ** Continue.
24310           datNewDate = .Date_New
24320           If .cmdPrice.Enabled = True Or .frmAssetPricing_Sub.Form.marketvalue.Enabled = True Then
                  ' ** If everything's already activated, don't do it again.
24330             blnRetVal = False
24340           End If
24350         Else
24360           blnRetVal = False
24370         End If
24380       Else
24390         .Date_New.SetFocus
24400         If IsNull(.Date_New.text) = False Then
24410           If Trim(.Date_New.text) <> vbNullString Then
24420             If IsDate(.Date_New.text) = True Then
                    ' ** Looks OK.
24430               datNewDate = .Date_New.text
24440             Else
24450               blnRetVal = False
24460             End If
24470           Else
24480             blnRetVal = False
24490           End If
24500         Else
24510           blnRetVal = False
24520         End If
24530       End If
24540     End If

24550     If blnRetVal = True Then
            ' ** If pricing is allowed, enable the button after they've entered a Date_New.
24560       .cmdPrice.Enabled = gblnPricingAllowed
24570       lngRecsCur = .frmAssetPricing_Sub.Form.RecCnt  ' ** Forms Function: frmAssetPricing_Sub.
24580       If lngRecsCur > 0& Then
24590         With .frmAssetPricing_Sub.Form
24600           .yield_entry.Enabled = True
24610           .yield_entry.Locked = False
24620           .yield_box.BackColor = CLR_WHT
24630           .yield_box.BorderColor = CLR_LTBLU2
24640           .yield_entry.ForeColor = CLR_BLK
24650           .marketvalue.Enabled = True
24660           .marketvalue.Locked = False
24670           .marketvalue_box.BackColor = CLR_WHT
24680           .marketvalue_loc_usd_box.ForeColor = CLR_WHT
24690           .currsym_symbol_loc_usd_box.ForeColor = CLR_WHT
24700           .marketvalue_box.BorderColor = CLR_LTBLU2
24710           .marketvalue.ForeColor = CLR_BLK
24720 On Error Resume Next
24730           .yield_entry.SetFocus
24740 On Error GoTo ERRH
24750         End With
24760       Else
24770         blnRetVal = False
24780         .cmdUpdate.Enabled = False
24790         .cmdPrice.Enabled = False
24800         .cmdReset.Enabled = False
24810         If blnTabBack = False Then
24820           MsgBox "No assets held by accounts.", vbInformation + vbOKOnly, _
                  (Left(("Nothing To Do" & Space(40)), 40) & "F03")
24830         Else
24840           blnTabBack = False
24850         End If
24860         .cmdClose.SetFocus
24870       End If
24880       If blnRetVal = True Then
24890         If blnTabBack = False Then
24900           .frmAssetPricing_Sub.SetFocus
24910           .frmAssetPricing_Sub.Form.marketvalue.SetFocus
24920         Else
24930           blnTabBack = False
24940           If .cmdUpdate.Enabled = True Then
24950             .cmdUpdate.SetFocus
24960           ElseIf .cmdHistory.Enabled = True Then
24970             .cmdHistory.SetFocus
24980           ElseIf .cmdReset.Enabled = True Then
24990             .cmdReset.SetFocus
25000           ElseIf .cmdPrintReport.Enabled = True Then
25010             .cmdPrintReport.SetFocus
25020           Else
25030             .cmdClose.SetFocus
25040           End If
25050         End If
25060       End If
25070     End If
25080   End With

EXITP:
25090   SetPricing = blnRetVal
25100   Exit Function

ERRH:
25110   blnRetVal = False
25120   THAT_PROC = THIS_PROC
25130   That_Erl = Erl
25140   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
25150   Resume EXITP

End Function

Private Sub IncludeCurrency(blnShow As Boolean)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim frm As Access.Form
        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long, lngTmp05 As Long, lngTmp06 As Long
        Dim lngX As Long

25210   With Me

25220     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
25230       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
25240     End If

          ' ** Variables are fed empty, then populated ByRef.
25250     GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

25260     FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

25270     Set frm = .frmAssetPricing_Sub.Form

25280     lngTmp04 = (frm.curr_id.Width + (4& * lngTpp))
25290     lngTmp04 = (lngTmp04 + (frm.marketvaluecurrent_usd_str_lbl.Width + (4& * lngTpp)))
25300     lngTmp04 = (lngTmp04 + (frm.marketvalue_usd_lbl.Width + (4& * lngTpp)))

25310     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
25320     lngMonitorNum = 1&: lngTmp05 = 0&
25330     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
25340     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

25350     Select Case blnShow
          Case True
25360       If lngWidth <> lngFrm_Width Then
              'If .Width <> lngForm_Width Then
              ' ** If it's currently not opening size, go ahead.

              ' ** Couldn't quite get this to stretch right to eliminate horizontal scroll bar.
              'lngTmp06 = lngFrm_Width + (3& * lngTpp)
25370         If lngMonitorNum = 1& Then lngTmp05 = lngFrm_Top
25380         DoCmd.MoveSize lngFrm_Left, lngTmp05, lngFrm_Width, lngFrm_Height  'lngFrm_Top
25390         If lngMonitorNum > 1& Then
25400           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
25410         End If
25420         .Width = lngForm_Width

25430         .cmdClose.Left = (.cmdClose.Left + lngTmp04)
25440         .cmdUpdate.Left = (.cmdUpdate.Left + lngTmp04)
25450         .cmdHistory.Left = (.cmdHistory.Left + lngTmp04)
25460         If .GoToReport_arw_hist_img.Visible = True Then
25470           .GoToReport_arw_hist_img.Left = ((.cmdHistory.Left - .GoToReport_arw_hist_img.Width) + (2& * lngTpp))
25480         End If
25490         .cmdReset.Left = (.cmdReset.Left + lngTmp04)
              'lngTmp06 = (.frmAssetPricing_Sub.Width + lngTmp04) + (3& * lngTpp)
25500         .frmAssetPricing_Sub.Width = (.frmAssetPricing_Sub.Width + lngTmp04)
25510         .frmAssetPricing_Sub_box.Width = (.frmAssetPricing_Sub_box.Width + lngTmp04)
25520         .Nav_box01.Width = .frmAssetPricing_Sub.Width
25530         .Nav_hline03.Width = .frmAssetPricing_Sub.Width
25540         .Header_vline01.Left = (.Header_vline01.Left + lngTmp04)
25550         .Header_vline02.Left = (.Header_vline02.Left + lngTmp04)
25560         .Footer_vline01.Left = (.Footer_vline01.Left + lngTmp04)
25570         .Footer_vline02.Left = (.Footer_vline02.Left + lngTmp04)
25580         .Header_hline01.Width = (.Header_hline01.Width + lngTmp04)
25590         .Header_hline02.Width = (.Header_hline02.Width + lngTmp04)
25600         .Footer_hline01.Width = (.Footer_hline01.Width + lngTmp04)
25610         .Footer_hline02.Width = (.Footer_hline02.Width + lngTmp04)
25620         lngTmp02 = ((.cmdClose.Left + .cmdClose.Width) - chkIncludeCurrency_lbl2.Width)
25630         .chkIncludeCurrency_lbl2.Left = lngTmp02

25640         For lngX = 1& To 24&
25650           strTmp01 = Right("00" & CStr(lngX), 2)
25660           .Controls("GoToReport_Emblem_" & strTmp01 & "_img").Left = (.Controls("GoToReport_Emblem_" & strTmp01 & "_img").Left + lngTmp04)
25670         Next

              ' ** This is centered on the form.
25680         lngTmp02 = ((.Width - .cmdPrice_box.Width) / 2)
25690         lngTmp03 = (.cmdPrice.Left - .cmdPrice_box.Left)
25700         lngTmp04 = (.cmdPrice_box.Left - .GoToReport_arw_price_img.Left)
25710         .cmdPrice_box.Left = lngTmp02
25720         .cmdPrice.Left = (.cmdPrice_box.Left + lngTmp03)
25730         .GoToReport_arw_price_img.Left = (.cmdPrice_box.Left - lngTmp04)

25740       End If
25750     Case False
25760       If lngWidth = lngFrm_Width Then
              'If .Width = lngForm_Width Then
              ' ** If it's currently opening size, go ahead.

25770         .cmdClose.Left = (.cmdClose.Left - lngTmp04)
25780         .cmdUpdate.Left = (.cmdUpdate.Left - lngTmp04)
25790         .cmdHistory.Left = (.cmdHistory.Left - lngTmp04)
25800         If .GoToReport_arw_hist_img.Visible = True Then
25810           .GoToReport_arw_hist_img.Left = ((.cmdHistory.Left - .GoToReport_arw_hist_img.Width) + (2& * lngTpp))
25820         End If
25830         .cmdReset.Left = (.cmdReset.Left - lngTmp04)
              'lngTmp06 = (.frmAssetPricing_Sub.Width - lngTmp04) + (3& * lngTpp)
25840         .frmAssetPricing_Sub.Width = (.frmAssetPricing_Sub.Width - lngTmp04)
25850         .frmAssetPricing_Sub_box.Width = (.frmAssetPricing_Sub_box.Width - lngTmp04)
25860         .Nav_box01.Width = .frmAssetPricing_Sub.Width
25870         .Nav_hline03.Width = .frmAssetPricing_Sub.Width
25880         .Header_vline01.Left = (.Header_vline01.Left - lngTmp04)
25890         .Header_vline02.Left = (.Header_vline02.Left - lngTmp04)
25900         .Footer_vline01.Left = (.Footer_vline01.Left - lngTmp04)
25910         .Footer_vline02.Left = (.Footer_vline02.Left - lngTmp04)
25920         .Header_hline01.Width = (.Header_hline01.Width - lngTmp04)
25930         .Header_hline02.Width = (.Header_hline02.Width - lngTmp04)
25940         .Footer_hline01.Width = (.Footer_hline01.Width - lngTmp04)
25950         .Footer_hline02.Width = (.Footer_hline02.Width - lngTmp04)
25960         lngTmp02 = ((.cmdClose.Left + .cmdClose.Width) - chkIncludeCurrency_lbl2.Width)
25970         .chkIncludeCurrency_lbl2.Left = lngTmp02

25980         For lngX = 1& To 24&
25990           strTmp01 = Right("00" & CStr(lngX), 2)
26000           .Controls("GoToReport_Emblem_" & strTmp01 & "_img").Left = (.Controls("GoToReport_Emblem_" & strTmp01 & "_img").Left - lngTmp04)
26010         Next

26020         .Width = (lngForm_Width - lngTmp04)

              'lngTmp06 = (lngWidth - lngTmp04) + (3& * lngTpp)
26030         If lngMonitorNum = 1& Then lngTmp05 = lngTop
26040         DoCmd.MoveSize (lngLeft + (lngTmp04 / 2)), lngTmp05, (lngWidth - lngTmp04), lngHeight  'lngTop
26050         If lngMonitorNum > 1& Then
26060           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
26070         End If

              ' ** This is centered between controls.
26080         lngTmp02 = (.cmdClose.Left - (.Date_New.Left + .Date_New.Width))
26090         lngTmp02 = ((lngTmp02 - .cmdPrice_box.Width) / 2)
26100         lngTmp02 = ((.Date_New.Left + .Date_New.Width) + lngTmp02)
26110         lngTmp03 = (.cmdPrice.Left - .cmdPrice_box.Left)
26120         lngTmp04 = (.cmdPrice_box.Left - .GoToReport_arw_price_img.Left)
26130         .cmdPrice_box.Left = lngTmp02
26140         .cmdPrice.Left = (.cmdPrice_box.Left + lngTmp03)
26150         .GoToReport_arw_price_img.Left = (.cmdPrice_box.Left - lngTmp04)

26160       End If
26170     End Select

26180   End With

EXITP:
26190   Set frm = Nothing
26200   Exit Sub

ERRH:
26210   Select Case ERR.Number
        Case Else
26220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26230   End Select
26240   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

26310   With Me
26320     blnGTR_Emblem = False
26330     For lngX = 1& To 24&
26340       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
26350     Next
26360   End With

EXITP:
26370   Exit Sub

ERRH:
26380   Select Case ERR.Number
        Case Else
26390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26400   End Select
26410   Resume EXITP

End Sub

Private Sub HistoryChk()

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "HistoryChk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim datCurrentDate As Date
        Dim lngRecs1 As Long, lngRecs2 As Long
        Dim blnAdd As Boolean, blnAddAll As Boolean

26510   With Me

26520     datCurrentDate = .Date_Current

26530     Set dbs = CurrentDb
26540     With dbs
            ' ** tblPricing_MasterAsset_History, by specified [curdat].
26550       Set qdf = .QueryDefs("qryAssetPricing_History_05")
26560       With qdf.Parameters
26570         ![curdat] = datCurrentDate
26580       End With
26590       Set rst = qdf.OpenRecordset
26600       blnAdd = False: blnAddAll = False
26610       If rst.BOF = True And rst.EOF = True Then
              ' ** No history for this date.
26620         blnAddAll = True
26630         lngRecs1 = 0&
26640         rst.Close
26650         Set rst = Nothing
26660         Set qdf = Nothing
26670       Else
26680         rst.MoveLast
26690         lngRecs1 = rst.RecordCount
26700         rst.Close
26710         Set rst = Nothing
26720         Set qdf = Nothing
26730         Set rst = .OpenRecordset("masterasset", dbOpenDynaset, dbReadOnly)
26740         rst.MoveLast
26750         lngRecs2 = rst.RecordCount
26760         If lngRecs1 < lngRecs2 Then
26770           blnAdd = True
26780         End If
26790         rst.Close
26800         Set rst = Nothing
26810         Set qdf = Nothing
26820       End If
26830       Select Case blnAddAll
            Case True
              ' ** Append qryAssetPricing_History_06_01 (MasterAsset, linked to tblCurrency, as new
              ' ** tblPricing_MasterAsset_History records, by specified [curdat]) to tblPricing_MasterAsset_History.
26840         Set qdf = .QueryDefs("qryAssetPricing_History_06_02")
26850         With qdf.Parameters
26860           ![curdat] = datCurrentDate
26870         End With
26880         qdf.Execute
26890         Set qdf = Nothing
26900       Case False
26910         If blnAdd = True Then
                ' ** Append qryAssetPricing_History_06_03 (qryAssetPricing_History_06_01 (MasterAsset,
                ' ** linked to tblCurrency, as new tblPricing_MasterAsset_History records, by specified
                ' ** [curdat]), not in tblPricing_MasterAsset_History) to tblPricing_MasterAsset_History.
26920           Set qdf = .QueryDefs("qryAssetPricing_History_06_04")
26930           With qdf.Parameters
26940             ![curdat] = datCurrentDate
26950           End With
26960           qdf.Execute
26970           Set qdf = Nothing
26980         End If
26990       End Select
27000       .Close
27010     End With

27020   End With

EXITP:
27030   Set rst = Nothing
27040   Set qdf = Nothing
27050   Set dbs = Nothing
27060   Exit Sub

ERRH:
27070   DoCmd.Hourglass False
27080   Select Case ERR.Number
        Case Else
27090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27100   End Select
27110   Resume EXITP

End Sub

Private Sub OpeningGambit()

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "OpeningGambit"

        Dim lngScreenHeight As Long, lngScreenWidth As Long, lngXBorders As Long, lngYBorders As Long
        Dim lngTmpTop As Long, lngTmpHeight As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

27210   With Me

27220     DoCmd.Hourglass True
27230     DoEvents

          ' ** Variables are fed empty, then populated ByRef.
27240     GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

27250     lngTmp01 = .Section(0).Height  ' ** Detail, in twips.
27260     lngTmp02 = .Section(1).Height  ' ** Header.
27270     lngTmp03 = .Section(2).Height  ' ** Footer.

27280     lngScreenWidth = GetScreenRes(GSR_HORZRES)  ' ** Module Function: modWindowFunctions.
27290     lngScreenHeight = GetScreenRes(GSR_VERTRES)  ' ** Module Function: modWindowFunctions.
27300     ConvertPIXELSToTWIPS lngScreenWidth, lngScreenHeight  ' ** Module Function: modWindowFunctions.

27310     GetWindowBorders Me, lngXBorders, lngYBorders  ' ** Module Function: modWindowFunctions.

27320     lngTmpTop = lngFrm_Top  ' ** Current form top.
27330     lngTmpHeight = ((lngTmp02 + lngTmp03 + (lngTmp01 * (RZ_WINROWS + 5&))) + lngYBorders) ' ** Ideal form height.
27340     lngTmp04 = ((lngScreenHeight - lngTmpHeight) / 2)  ' ** Top of centered ideal form.
          ' ** If the (current top + lngTmpHeight) is within lngScreenHeight then...
27350     If (lngTmpTop + lngTmpHeight) < lngScreenHeight Then
            ' ** If the bottom gutter is smaller than the top, move it up.
27360       If (lngScreenHeight - (lngTmpTop + lngTmpHeight)) < lngTmpTop Then
27370         lngTmpTop = lngTmp04
27380       End If
27390     Else
            ' ** If the ideal height is showable, then...
27400       If lngTmpHeight < lngScreenHeight Then
              ' ** If the (current top + lngTmpHeight) puts the form bottom out of sight, then...
27410         If (lngTmpTop + lngTmpHeight) > lngScreenHeight Then
27420           lngTmpTop = ((lngScreenHeight - lngTmpHeight) / 2&)
27430         End If
27440       Else
27450         lngTmpTop = 300&
27460         lngTmpHeight = (lngScreenHeight - (lngTmpTop * 2&))
27470       End If
27480     End If
27490     lngTmpHeight = lngTmpHeight - 30&  ' ** Cosmological Constant.
27500     DoEvents

          ' ** These all happen before the Prefs are loaded!
27510     chkShowAssetNo_AfterUpdate  ' ** Procedure: Above.
27520     DoEvents
27530     chkIncludeCurrency_AfterUpdate  ' ** Procedure: Above.
27540     DoEvents

27550     DoCmd.Hourglass True
27560     DoEvents

27570   End With

EXITP:
27580   Exit Sub

ERRH:
27590   DoCmd.Hourglass False
27600   Select Case ERR.Number
        Case Else
27610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27620   End Select
27630   Resume EXITP

End Sub

Public Sub ForExTots_Set(blnRecalc As Boolean)

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "ForExTots_Set"

27710   With Me
27720     Select Case blnRecalc
          Case True
27730       blnForExTots = True
27740       .TimerInterval = 400&
27750     Case False
27760       blnForExTots = False
27770     End Select
27780   End With

EXITP:
27790   Exit Sub

ERRH:
27800   Select Case ERR.Number
        Case Else
27810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27820   End Select
27830   Resume EXITP

End Sub
