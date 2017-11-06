VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssetPricing_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssetPricing_History"

'VGC 03/23/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' ##   ShowAssetTypeDesc()
' ##   IncludeCurrency()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Date:             D {lbxHistoryDate}
' **   One:              E {ckgDeleteDates_chkOne}
' **   Show Hidden:      H {chkShowHidden}
' **   Include Currency: L {chkIncludeCurrency}
' **   Market Value:     M {marketvaluecurrent on frmAssetPricing_History_Sub}
' **   Show AssetNo:     N {chkShowAssetNo}
' **   Oldest:           O {ckgDeleteDates_chkOldest}
' **   Print Report:     P {cmdPrintReport}
' **   Restore:          R {cmdRestore}
' **   Shorten:          S {cmdShorten on frmAssetPricing_History_Sub}
' **   Delete:           T {cmdHistoryDeleteOldest, cmdHistoryDeleteOne}
' **   Widen:            W {cmdWiden on frmAssetPricing_History_Sub}
' **   Exit:             X {cmdClose}
' **   Show Description: Y {chkShowAssetTypeDesc}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean
Private blnHistoryDeleteOldest_Focus As Boolean, blnHistoryDeleteOldest_MouseDown As Boolean
Private blnHistoryDeleteOne_Focus As Boolean, blnHistoryDeleteOne_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngFrmLeft_Adj As Long, lngFrmTop_Adj As Long
Private lngForm_Width As Long, lngClose_Left As Long, lngRestore_Left As Long, lngSub_Width As Long
Private lngCurrID_Width As Long, lngTypeDesc_Width As Long, lngPrintReportLbl_Left As Long
Private lngAssetNoLbl_Width As Long, lngCurrencyLbl_Width As Long, lngAssetTypeDescLbl_Width As Long
Private lngChkBoxLbl_Top As Long, lngChkBoxLbl_Height As Long, lngChkBoxLbl_Offset As Long

Private blnGoingToReport As Boolean, blnUpdateTotals As Boolean
Private lngDateCnt As Long
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim usr As DAO.User, grp As DAO.Group, frm As Access.Form
        Dim blnFound As Boolean
        Dim lngTmp01 As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           lngDateCnt = .lbxHistoryDate.ListCount
170           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

180           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
190           blnHistoryDeleteOldest_Focus = False: blnHistoryDeleteOldest_MouseDown = False
200           blnHistoryDeleteOne_Focus = False: blnHistoryDeleteOne_MouseDown = False
210           gblnHasForExThis = False: blnUpdateTotals = False

220           If gblnGoToReport = True Then
230             .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
240             .GoToReport_arw_pr_img.Visible = True
250             DoEvents
260           End If

270           blnFound = False
280           For Each grp In DBEngine.Workspaces(0).Groups
290             If grp.Name = "ViewOnly" Then
300               For Each usr In grp.Users
310                 If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
320                   blnFound = True
330                   Exit For
340                 End If
350               Next
360             End If
370           Next

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
380           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

390           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

              ' ** Variables are fed empty, then populated ByRef.
400           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.
410           lngFrm_Width = (lngFrm_Width + (3& * lngTpp))

420           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

430           DoCmd.SelectObject acForm, THIS_NAME, False

440           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
450           lngMonitorNum = 1&: lngTmp01 = 0&
460           EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
470           If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

480           If lngMonitorNum = 1& Then lngTmp01 = (lngFrm_Top + 120&)
490           DoCmd.MoveSize lngFrm_Left + 120&, lngTmp01, lngFrm_Width, lngFrm_Height  'lngFrm_Top + 120&
500           If lngMonitorNum > 1& Then
510             LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
520           End If

530           lngFrm_Top = (lngFrm_Top - (8& * lngTpp))
540           lngFrm_Left = (lngFrm_Left - (8& * lngTpp))

550           Set frm = .frmAssetPricing_History_Sub.Form

560           .form_width_line.Left = 0&  ' ** Just get this out of the way.

570           lngForm_Width = .Width
580           lngRestore_Left = .cmdRestore.Left

590           lngCurrID_Width = (((frm.curr_code.Left + frm.curr_code.Width) - _
                (frm.assettype_description.Left + frm.assettype_description.Width)) + _
                ((frm.MarketValueCurrent_usd.Left + frm.MarketValueCurrent_usd.Width) - _
                (frm.marketvaluecurrent_str.Left + frm.marketvaluecurrent_str.Width)))

600           lngTypeDesc_Width = ((frm.assettype_description.Left + frm.assettype_description.Width) - (frm.assettype.Left + frm.assettype.Width))
610           lngClose_Left = .cmdClose.Left
620           lngRestore_Left = .cmdRestore.Left
630           lngSub_Width = (.frmAssetPricing_History_Sub.Width + (3& * lngTpp))
640           lngAssetNoLbl_Width = .chkShowAssetNo_lbl.Width
650           lngCurrencyLbl_Width = .chkIncludeCurrency_lbl.Width
660           lngAssetTypeDescLbl_Width = .chkShowAssetTypeDesc_lbl.Width
670           lngChkBoxLbl_Top = .chkShowAssetNo_lbl.Top
680           lngChkBoxLbl_Height = .chkShowAssetNo_lbl.Height
690           lngChkBoxLbl_Offset = (.chkShowAssetNo_lbl.Left - .chkShowAssetNo.Left)
700           lngFrmLeft_Adj = (8& * lngTpp)
710           lngFrmTop_Adj = (8& * lngTpp)

720           gblnHasForExThis = gblnHasForEx
730           Select Case gblnHasForEx
              Case True
740             Select Case gblnHasForExThis
                Case True
750               If Pref_HasPref(THIS_NAME, "chkIncludeCurrency") = False Then  ' ** Module Procedure: modPreferenceFuncs.
760                 .chkIncludeCurrency = True
770               End If
780               .chkIncludeCurrency.Enabled = True
790             Case False
800               .chkIncludeCurrency.Enabled = True
810             End Select
820           Case False
830             .chkIncludeCurrency = False
840             .chkIncludeCurrency.Enabled = False
850           End Select

860           If blnFound = True Then
870             .cmdRestore.Enabled = False
880           End If

              ' ** Form_Load() triggers here, before getting to the next line.
890           chkShowAssetTypeDesc_AfterUpdate  ' ** Procedure: Below.
900           chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
910           chkShowAssetNo_AfterUpdate  ' ** Procedure: Below.
920           chkShowHidden_AfterUpdate  ' ** Procedure: Below.
930           DoEvents

940           .HistoryDeleteOldest = 0&
950           .cmdHistoryDeleteOldest.Enabled = False
960           .cmdHistoryDeleteOldest_raised_img_dis.Visible = True
970           .cmdHistoryDeleteOldest_raised_img.Visible = False
980           .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
990           .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
1000          .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
1010          .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False

1020          .HistoryDeleteOne = Null
1030          .cmdHistoryDeleteOne.Enabled = False
1040          .cmdHistoryDeleteOne_raised_img.Visible = False
1050          .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = False
1060          .cmdHistoryDeleteOne_raised_focus_img.Visible = False
1070          .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = False
1080          .cmdHistoryDeleteOne_sunken_focus_dots_img.Visible = False
1090          .cmdHistoryDeleteOne_raised_img_dis.Visible = False

1100          ckgDeleteDates_chkOldest_AfterUpdate  ' ** Procedure: Below.
1110          ckgDeleteDates_chkOne_AfterUpdate  ' ** Procedure: Below.

1120          blnUpdateTotals = True
1130          .TimerInterval = 500&

1140          DoCmd.Hourglass False

1150        Else
1160          Cancel = -1
1170        End If
1180      Else
1190        Cancel = -1
1200      End If
1210    End With

1220    If Cancel = -1 Then
1230      DoCmd.Hourglass False
1240      Beep
1250      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1260    Else
1270      If lngDateCnt = 0 Then
1280        Cancel = -1
1290        DoCmd.Hourglass False
1300        MsgBox "There are no previous asset prices to view.", vbInformation + vbOKOnly, "Nothing To Do"
1310      End If
1320    End If

EXITP:
1330    Set frm = Nothing
1340    Set usr = Nothing
1350    Set grp = Nothing
1360    Exit Sub

ERRH:
1370    DoCmd.Hourglass False
1380    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
1390    Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Private Sub Form_Load()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim strFilter As String

1510    With Me

          ' ** By default, give it first date in list box.
1520      .lbxHistoryDate = .lbxHistoryDate.Column(0, 0)

          ' ** lbxHistoryDate: qryPricing_History_06.
          ' ** Column 0: currentDate
          ' ** Column 1: cnt

1530      If .lbxHistoryDate <> .lbxHistoryDate.Column(0, 0) Then
            ' ** This means it got one from Pref_Load().
1540        If IsNull(.lbxHistoryDate.Column(2)) = True Then
              ' ** The last saved date is no longer there.
1550          .lbxHistoryDate = .lbxHistoryDate.Column(0, 0)
1560        End If
1570      End If

1580      strFilter = "[currentDate] = #" & .lbxHistoryDate & "#"
1590      If .chkShowHidden = False Then
1600        strFilter = strFilter & " And [IsHid] = False"
1610      End If
1620      .frmAssetPricing_History_Sub.Form.Filter = strFilter
1630      .frmAssetPricing_History_Sub.Form.FilterOn = True
1640      .frmAssetPricing_History_Sub.Form.Form_Load  ' ** Form Procedure: frmAssetPricing_History_Sub.

1650      .HistoryDateCnt = lngDateCnt

1660    End With

EXITP:
1670    Exit Sub

ERRH:
1680    DoCmd.Hourglass False
1690    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
1700    Case Else
1710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1720    End Select
1730    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1810    intShiftDown = (Shift And acShiftMask) > 0
1820    intAltDown = (Shift And acAltMask) > 0
1830    intCtrlDown = (Shift And acCtrlMask) > 0

1840    intRetVal = KeyCode

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Market Value:     M {marketvaluecurrent on frmAssetPricing_History_Sub}
        ' **   Delete:           T {cmdHistoryDeleteOldest, cmdHistoryDeleteOne}

        ' ** Plain keys.
1850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1860      Select Case intRetVal
          Case vbKeyEscape
1870        intRetVal = 0
1880        cmdClose_Click  ' ** Procedure: Below.
1890      End Select
1900    End If

        ' ** Alt keys.
1910    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1920      Select Case intRetVal
          Case vbKeyM
1930        With Me
1940          intRetVal = 0
1950          lngRecsCur = .frmAssetPricing_History_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_History_Sub.
1960          If lngRecsCur > 0& Then
1970            .frmAssetPricing_History_Sub.SetFocus
1980            If .frmAssetPricing_History_Sub.Form.marketvaluecurrent.Enabled = True Then
1990              .frmAssetPricing_History_Sub.Form.marketvaluecurrent.SetFocus
2000            End If
2010          Else
2020            Beep
2030          End If
2040        End With
2050      Case vbKeyT
2060        With Me
2070          intRetVal = 0
2080          If .cmdHistoryDeleteOldest.Visible = True And .cmdHistoryDeleteOldest.Enabled = True Then
2090            .cmdHistoryDeleteOldest.SetFocus
2100            cmdHistoryDeleteOldest_Click  ' ** Procedure: Below
2110          ElseIf .cmdHistoryDeleteOne.Visible = True And .cmdHistoryDeleteOne.Enabled = True Then
2120            .cmdHistoryDeleteOne.SetFocus
2130            cmdHistoryDeleteOne_Click  ' ** Procedure: Below
2140          Else
2150            Beep
2160          End If
2170        End With
2180      Case vbKeyX
2190        intRetVal = 0
2200        cmdClose_Click  ' ** Procedure: Below.
2210      End Select
2220    End If

        ' ** Ctrl-Shift keys.
2230    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyF
2250        With Me
2260          intRetVal = 0
2270          .FocusHolder.SetFocus
2280        End With
2290      End Select
2300    End If

EXITP:
2310    KeyCode = intRetVal
2320    Exit Sub

ERRH:
2330    intRetVal = 0
2340    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
2350    Case Else
2360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2370    End Select
2380    Resume EXITP

End Sub

Private Sub Form_Timer()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2410    With Me
2420      .TimerInterval = 0&
2430      If gblnSetFocus = True Then
2440        gblnSetFocus = False
2450        strFormName = vbNullString
2460  On Error Resume Next
2470        Do Until Screen.ActiveForm.Name = strFormName
2480          DoCmd.SelectObject acForm, THIS_NAME, False
2490          strFormName = Screen.ActiveForm.Name
2500        Loop
2510  On Error GoTo ERRH
2520      End If  ' ** gblnSetFocus.
2530      If gblnGoToReport = True Then
2540        DoCmd.Hourglass True  ' ** Make sure it's still running.
2550        DoEvents
2560        Select Case blnGoingToReport
            Case True
2570          gblnGoToReport = False
2580          blnGoingToReport = False
2590          .GoToReport_arw_pr_img.Visible = False
2600          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
2610          .cmdClose.SetFocus
2620          DoCmd.Hourglass False  ' ** Should already be off.
2630          DoEvents
2640        Case False
2650          blnGoingToReport = True
2660          .GoToReport_arw_pr_img.Visible = True  ' ** Should already be on.
2670          If .cmdPrintReport.Enabled = True Then
2680            .cmdPrintReport.SetFocus
2690          End If
2700          Beep
2710          DoEvents
2720          DoCmd.Hourglass False
              '.TimerInterval = (GTR_WAIT * 2&)  ' ** .TimerInterval set below.
2730        End Select
2740      End If
2750      If blnUpdateTotals = True Then
2760        blnUpdateTotals = False
2770        .IsHidSum = Abs(.frmAssetPricing_History_Sub.Form.IsHidSum)
2780        DoEvents
2790        ShowHiddenAssets  ' ** Procedure: Below.
2800        If blnIsOpen = True Then
2810          blnIsOpen = False
2820          lbxHistoryDate_AfterUpdate  ' ** Procedure: Below.
2830          DoEvents
2840          If blnGoingToReport = True Then
2850            .TimerInterval = (GTR_WAIT * 2&)
2860          End If
2870        End If
2880      End If
2890    End With

EXITP:
2900    Exit Sub

ERRH:
2910    DoCmd.Hourglass False
2920    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
2930    Case Else
2940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950    End Select
2960    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** If Close is clicked before timer out, these remain on.
3010    gblnGoToReport = False
3020    blnGoingToReport = False

3030    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Assure all reports are closed.
3040    Do While Reports.Count > 0
3050      DoCmd.Close acReport, Reports(0).Name
3060    Loop

3070    Select Case strCallingForm
        Case "frmAssetPricing"
3080      DoCmd.SelectObject acForm, strCallingForm, False
3090    End Select

EXITP:
3100    Exit Sub

ERRH:
3110    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
3120    Case Else
3130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3140    End Select
3150    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3210    With Me
3220      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3230        Select Case blnPrintReport_Focus
            Case True
3240          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3250          .cmdPrintReport_raised_img.Visible = False
3260        Case False
3270          .cmdPrintReport_raised_img.Visible = True
3280          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3290        End Select
3300        .cmdPrintReport_raised_focus_img.Visible = False
3310        .cmdPrintReport_raised_focus_dots_img.Visible = False
3320        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3330        .cmdPrintReport_raised_img_dis.Visible = False
3340      End If
3350      If .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = True Or .cmdHistoryDeleteOldest_raised_focus_img.Visible = True Then
3360        Select Case blnHistoryDeleteOldest_Focus
            Case True
3370          .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = True
3380          .cmdHistoryDeleteOldest_raised_img.Visible = False
3390        Case False
3400          .cmdHistoryDeleteOldest_raised_img.Visible = True
3410          .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
3420        End Select
3430        .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
3440        .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
3450        .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False
3460        .cmdHistoryDeleteOldest_raised_img_dis.Visible = False
3470      End If
3480      If .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = True Or .cmdHistoryDeleteOne_raised_focus_img.Visible = True Then
3490        Select Case blnHistoryDeleteOne_Focus
            Case True
3500          .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = True
3510          .cmdHistoryDeleteOne_raised_img.Visible = False
3520        Case False
3530          .cmdHistoryDeleteOne_raised_img.Visible = True
3540          .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = False
3550        End Select
3560        .cmdHistoryDeleteOne_raised_focus_img.Visible = False
3570        .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = False
3580        .cmdHistoryDeleteOne_sunken_focus_dots_img.Visible = False
3590        .cmdHistoryDeleteOne_raised_img_dis.Visible = False
3600      End If
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

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

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

Public Sub cmdClose_Click()

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3810    DoCmd.Close acForm, THIS_NAME

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
3840    Case Else
3850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3860    End Select
3870    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
3910    intShiftDown = (Shift And acShiftMask) > 0
3920    intAltDown = (Shift And acAltMask) > 0
3930    intCtrlDown = (Shift And acCtrlMask) > 0

3940    intRetVal = KeyCode

        ' ** Plain keys.
3950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3960      Select Case intRetVal
          Case vbKeyTab
3970        With Me
3980          intRetVal = 0
3990          If .lbxHistoryDate.Enabled = True Then
4000            .lbxHistoryDate.SetFocus
4010          Else
4020            Beep
4030          End If
4040        End With
4050      End Select
4060    End If

        ' ** Shift keys.
4070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4080      Select Case intRetVal
          Case vbKeyTab
4090        With Me
4100          intRetVal = 0
4110          If .cmdRestore.Enabled = True Then
4120            .cmdRestore.SetFocus
4130          Else
4140            Beep
4150          End If
4160        End With
4170      End Select
4180    End If

EXITP:
4190    KeyCode = intRetVal
4200    Exit Sub

ERRH:
4210    intRetVal = 0
4220    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
4230    Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub lbxHistoryDate_AfterUpdate()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryDate_AfterUpdate"

        Dim strFilter As String

4310    With Me
4320      strFilter = "[currentDate] = #" & .lbxHistoryDate & "#"
4330      gdatMarketDate = .lbxHistoryDate
4340      If .chkShowHidden = False Then
4350        strFilter = strFilter & " And [IsHid] = False"
4360      End If
4370      .frmAssetPricing_History_Sub.Form.Filter = strFilter
4380      .frmAssetPricing_History_Sub.Form.FilterOn = True
4390      .frmAssetPricing_History_Sub.Form.Form_Load  ' ** Form Procedure: frmAssetPricing_History_Sub.
4400      DoEvents
4410      If .HistoryDeleteOne.Visible = True Then
4420        .HistoryDeleteOne = .lbxHistoryDate
4430      End If
4440      blnUpdateTotals = True
4450      .TimerInterval = 500&
4460    End With

EXITP:
4470    Exit Sub

ERRH:
4480    Select Case ERR.Number
        Case Else
4490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4500    End Select
4510    Resume EXITP

End Sub

Private Sub lbxHistoryDate_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "lbxHistoryDate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
4610    intShiftDown = (Shift And acShiftMask) > 0
4620    intAltDown = (Shift And acAltMask) > 0
4630    intCtrlDown = (Shift And acCtrlMask) > 0

4640    intRetVal = KeyCode

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4670        With Me
4680          intRetVal = 0
4690          lngRecsCur = .frmAssetPricing_History_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_History_Sub.
4700          If lngRecsCur > 0& Then
4710            .frmAssetPricing_History_Sub.SetFocus
4720            .frmAssetPricing_History_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAssetPricing_History_Sub.
4730            .frmAssetPricing_History_Sub.Form.cusip.SetFocus
4740          Else
4750            .cmdClose.SetFocus
4760          End If
4770        End With
4780      End Select
4790    End If

        ' ** Shift keys.
4800    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4810      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4820        With Me
4830          intRetVal = 0
4840          .cmdClose.SetFocus
4850        End With
4860      End Select
4870    End If

        ' ** Ctrl keys.
4880    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4890      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4900        With Me
4910          intRetVal = 0
4920          If .ckgDeleteDates_chkOldest.Enabled = True Then
4930            .ckgDeleteDates_chkOldest.SetFocus
4940          Else
4950            Beep
4960          End If
4970        End With
4980      End Select
4990    End If

EXITP:
5000    KeyCode = intRetVal
5010    Exit Sub

ERRH:
5020    intRetVal = 0
5030    Select Case ERR.Number
        Case Else
5040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5050    End Select
5060    Resume EXITP

End Sub

Private Sub ckgDeleteDates_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDeleteDates_box_MouseMove"

5110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub ckgDeleteDates_chkOldest_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDeleteDates_chkOldest_AfterUpdate"

5210    With Me
5220      Select Case .ckgDeleteDates_chkOldest
          Case True
5230        .ckgDeleteDates_chkOldest_lbl.FontBold = True
5240        .ckgDeleteDates_chkOne = False
5250        .ckgDeleteDates_chkOne_lbl.FontBold = False
5260        .HistoryDelete_box.Visible = True
5270        .HistoryDelete_vline01.Visible = True
5280        .HistoryDelete_vline02.Visible = True
5290        .HistoryDelete_vline03.Visible = True
5300        .HistoryDelete_vline04.Visible = True
5310        .HistoryDelete_hline01.Visible = True
5320        .HistoryDelete_hline02.Visible = True
5330        .HistoryDeleteOne.Visible = False
5340        .cmdHistoryDeleteOne.Visible = False
5350        .cmdHistoryDeleteOne.Enabled = False
5360        .cmdHistoryDeleteOne_raised_img.Visible = False
5370        .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = False
5380        .cmdHistoryDeleteOne_raised_focus_img.Visible = False
5390        .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = False
5400        .cmdHistoryDeleteOne_sunken_focus_dots_img.Visible = False
5410        .cmdHistoryDeleteOne_raised_img_dis.Visible = False
5420        .HistoryDeleteOldest = 0
5430        .HistoryDeleteOldest.Visible = True
5440        .HistoryDeleteOldest_lbl2.Visible = True
5450        .cmdHistoryDeleteOldest.Visible = True
5460        .cmdHistoryDeleteOldest.Enabled = False
5470        .cmdHistoryDeleteOldest_raised_img_dis.Visible = True
5480      Case False
5490        .ckgDeleteDates_chkOldest_lbl.FontBold = False
5500        .HistoryDeleteOldest.Visible = False
5510        .HistoryDeleteOldest = 0
5520        .HistoryDeleteOldest_lbl2.Visible = False
5530        .cmdHistoryDeleteOldest.Visible = False
5540        .cmdHistoryDeleteOldest.Enabled = False
5550        .cmdHistoryDeleteOldest_raised_img.Visible = False
5560        .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
5570        .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
5580        .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
5590        .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False
5600        .cmdHistoryDeleteOldest_raised_img_dis.Visible = False
5610        .HistoryDelete_vline01.Visible = False
5620        .HistoryDelete_vline02.Visible = False
5630        .HistoryDelete_vline03.Visible = False
5640        .HistoryDelete_vline04.Visible = False
5650        .HistoryDelete_hline01.Visible = False
5660        .HistoryDelete_hline02.Visible = False
5670        .HistoryDelete_box.Visible = False
5680      End Select
5690    End With

EXITP:
5700    Exit Sub

ERRH:
5710    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
5720    Case Else
5730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5740    End Select
5750    Resume EXITP

End Sub

Private Sub ckgDeleteDates_chkOldest_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDeleteDates_chkOldest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
5810    intShiftDown = (Shift And acShiftMask) > 0
5820    intAltDown = (Shift And acAltMask) > 0
5830    intCtrlDown = (Shift And acCtrlMask) > 0

5840    intRetVal = KeyCode

        ' ** Plain keys.
5850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5870        With Me
5880          intRetVal = 0
5890          If .ckgDeleteDates_chkOldest = True And .HistoryDeleteOldest.Visible = True Then
5900            .HistoryDeleteOldest.SetFocus
5910          Else
5920            .ckgDeleteDates_chkOne.SetFocus
5930          End If
5940        End With
5950      End Select
5960    End If

        ' ** Shift keys.
5970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5980      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5990        With Me
6000          intRetVal = 0
6010          .lbxHistoryDate.SetFocus
6020        End With
6030      End Select
6040    End If

EXITP:
6050    KeyCode = intRetVal
6060    Exit Sub

ERRH:
6070    intRetVal = 0
6080    Select Case ERR.Number
        Case Else
6090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6100    End Select
6110    Resume EXITP

End Sub

Private Sub HistoryDelete_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "HistoryDelete_box_MouseMove"

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

Private Sub HistoryDeleteOldest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "HistoryDeleteOldest_MouseMove"

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

Private Sub HistoryDeleteOldest_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "HistoryDeleteOldest_lbl2_MouseMove"

6410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub HistoryDeleteOldest_AfterUpdate()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "HistoryDeleteOldest_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim datCurDate As Date
        Dim lngDelDates As Long
        Dim blnDeleteDates As Boolean
        Dim lngX As Long, lngY As Long

6510    With Me

6520      DoCmd.Hourglass True
6530      DoEvents

6540      blnDeleteDates = False
6550      datCurDate = 0

6560      If IsNull(.lbxHistoryDate) = False Then
6570        datCurDate = .lbxHistoryDate
6580      End If

6590      Select Case IsNull(.HistoryDeleteOldest)
          Case True
6600        .HistoryDeleteOldest = 0&
6610        .cmdHistoryDeleteOldest.Enabled = False
6620      Case False
6630        If .HistoryDeleteOldest < 0 Then
6640          Beep
6650          .HistoryDeleteOldest = 0&
6660          .cmdHistoryDeleteOldest.Enabled = False
6670        Else
6680          blnDeleteDates = True
6690          If .HistoryDeleteOldest > .lbxHistoryDate.ListCount Then
6700            .HistoryDeleteOldest = .lbxHistoryDate.ListCount
6710          End If
6720          If .HistoryDeleteOldest = 1 Then
6730            .HistoryDeleteOldest_lbl2.Caption = "Date"
6740          Else
6750            .HistoryDeleteOldest_lbl2.Caption = "Dates"
6760          End If
6770          .cmdHistoryDeleteOldest.Enabled = True
6780        End If
6790      End Select

6800      Select Case .cmdHistoryDeleteOldest.Enabled
          Case True
6810        .cmdHistoryDeleteOldest_raised_img.Visible = True
6820        .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
6830        .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
6840        .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
6850        .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False
6860        .cmdHistoryDeleteOldest_raised_img_dis.Visible = False
6870      Case False
6880        .cmdHistoryDeleteOldest_raised_img_dis.Visible = True
6890        .cmdHistoryDeleteOldest_raised_img.Visible = False
6900        .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
6910        .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
6920        .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
6930        .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False
6940      End Select
6950      DoEvents

6960      Select Case blnDeleteDates
          Case True

6970        lngDelDates = .HistoryDeleteOldest

6980        Set dbs = CurrentDb
6990        With dbs

              ' ** Update tblMark, for mark = False.
7000          Set qdf = .QueryDefs("qryPricing_History_06g")
7010          qdf.Execute
7020          Set qdf = Nothing

7030          lngY = 0&
7040          For lngX = (lngDateCnt - 1&) To 0& Step -1&
                ' ** Update qryPricing_History_06c (tblPricing_MasterAsset_History,
                ' ** linked to tblMark, with Date_Del), by specified [curdat], [deldat].
7050            Set qdf = .QueryDefs("qryPricing_History_06d")
7060            With qdf.Parameters
7070              ![curdat] = Me.lbxHistoryDate.Column(0, lngX)
7080              ![deldat] = True
7090            End With
7100            qdf.Execute
7110            Set qdf = Nothing
7120            lngY = lngY + 1&
7130            If lngY = lngDelDates Then
7140              Exit For
7150            End If
7160          Next

7170          .Close
7180        End With

7190        If .lbxHistoryDate.RowSource <> "qryPricing_History_06f" Then
7200          .lbxHistoryDate.RowSource = "qryPricing_History_06f"
7210          .lbxHistoryDate.ColumnWidths = "1304;180;0"
7220        End If
7230        .lbxHistoryDate.Requery

7240      Case False
7250        If .lbxHistoryDate.RowSource <> "qryPricing_History_06" Then
7260          .lbxHistoryDate.RowSource = "qryPricing_History_06"
7270          .lbxHistoryDate.ColumnWidths = "1304;0;0"
7280          .lbxHistoryDate.Requery
7290        End If
7300      End Select

7310      If datCurDate > 0 Then
7320        .lbxHistoryDate = datCurDate
7330      End If

7340      .lbxHistoryDate.SetFocus

7350    End With

7360    DoCmd.Hourglass False

EXITP:
7370    Set qdf = Nothing
7380    Set dbs = Nothing
7390    Exit Sub

ERRH:
7400    DoCmd.Hourglass False
7410    Select Case ERR.Number
        Case Else
7420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7430    End Select
7440    Resume EXITP

End Sub

Private Sub HistoryDeleteOldest_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "HistoryDeleteOldest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
7510    intShiftDown = (Shift And acShiftMask) > 0
7520    intAltDown = (Shift And acAltMask) > 0
7530    intCtrlDown = (Shift And acCtrlMask) > 0

7540    intRetVal = KeyCode

        ' ** Plain keys.
7550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7570        With Me
7580          intRetVal = 0
7590  On Error Resume Next
7600          If .HistoryDeleteOldest > 0 And .cmdHistoryDeleteOldest.Enabled = True Then
7610            .cmdHistoryDeleteOldest.SetFocus
7620          Else
7630            .ckgDeleteDates_chkOne.SetFocus
7640          End If
7650  On Error GoTo ERRH
7660        End With
7670      End Select
7680    End If

        ' ** Shift keys.
7690    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7700      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7710        With Me
7720          intRetVal = 0
7730          .ckgDeleteDates_chkOldest.SetFocus
7740        End With
7750      End Select
7760    End If

EXITP:
7770    KeyCode = intRetVal
7780    Exit Sub

ERRH:
7790    intRetVal = 0
7800    Select Case ERR.Number
        Case Else
7810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7820    End Select
7830    Resume EXITP

End Sub

Private Sub HistoryDeleteOne_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "HistoryDeleteOne_MouseMove"

7910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7920    Exit Sub

ERRH:
7930    Select Case ERR.Number
        Case Else
7940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7950    End Select
7960    Resume EXITP

End Sub

Private Sub cmdHistoryDeleteOldest_GotFocus()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOldest_GotFocus"

8010    With Me
8020      blnHistoryDeleteOldest_Focus = True
8030      .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = True
8040      .cmdHistoryDeleteOldest_raised_img.Visible = False
8050      .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
8060      .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
8070      .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False
8080      .cmdHistoryDeleteOldest_raised_img_dis.Visible = False
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

Private Sub cmdHistoryDeleteOldest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOldest_MouseDown"

8210    With Me
8220      blnHistoryDeleteOldest_MouseDown = True
8230      .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = True
8240      .cmdHistoryDeleteOldest_raised_img.Visible = False
8250      .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
8260      .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
8270      .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
8280      .cmdHistoryDeleteOldest_raised_img_dis.Visible = False
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

Public Sub cmdHistoryDeleteOldest_Click()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOldest_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strFilter As String
        Dim lngTotalDates As Long, lngTotalRecs As Long
        Dim lngDels As Long, arr_varDel() As Variant
        Dim lngRecs As Long
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant
        Dim lngX As Long, lngE As Long
        Dim blnContinue As Boolean, blnDelAll As Boolean

        ' ** Array: arr_varDel().
        Const D_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const D_DAT As Integer = 0
        Const D_CNT As Integer = 1

8410    blnContinue = True
8420    blnDelAll = False

8430    With Me
8440      If .lbxHistoryDate.ListCount > 0 Then
8450        If IsNull(.HistoryDeleteOldest) = False Then
8460          If .HistoryDeleteOldest > 0 Then
8470            If .HistoryDeleteOldest > .lbxHistoryDate.ListCount Then
8480              .HistoryDeleteOldest = .lbxHistoryDate.ListCount
8490            End If
8500            If .HistoryDeleteOldest = .lbxHistoryDate.ListCount Then
8510              varTmp00 = DCount("*", "tblPricing_MasterAsset_History")
8520              lngTotalDates = .lbxHistoryDate.ListCount
8530              lngTotalRecs = varTmp00
8540              Beep
8550              msgResponse = MsgBox("Are you sure you want to delete all the pricing history?" & vbCrLf & vbCrLf & _
                    "Dates To Be Deleted: " & CStr(lngTotalDates) & vbCrLf & _
                    "Price Records: " & CStr(lngTotalRecs), vbQuestion + vbOKCancel + vbDefaultButton2, "Delete All Pricing History")
8560              If msgResponse <> vbOK Then
8570                blnContinue = False
8580              Else
8590                blnDelAll = True
8600              End If
8610            Else

8620              lngTotalDates = .HistoryDeleteOldest
8630              lngTotalRecs = 0&

8640              lngDels = 0&
8650              ReDim arr_varDel(D_ELEMS, 0)

8660              Set dbs = CurrentDb
8670              With dbs
                    ' ** tblPricing_MasterAsset_History, grouped by CurrentDate, with cnt.
8680                Set qdf = .QueryDefs("qryPricing_History_09")
8690                Set rst = qdf.OpenRecordset
8700                With rst
8710                  .MoveLast
8720                  lngRecs = .RecordCount
8730                  .MoveFirst
8740                  For lngX = 1& To lngTotalDates
8750                    lngDels = lngDels + 1&
8760                    lngE = lngDels - 1&
8770                    ReDim Preserve arr_varDel(D_ELEMS, lngE)
8780                    arr_varDel(D_DAT, lngE) = ![currentDate]
8790                    arr_varDel(D_CNT, lngE) = ![cnt]
8800                    lngTotalRecs = lngTotalRecs + ![cnt]
8810                    If lngX = lngRecs Then
8820                      Exit For
8830                    Else
8840                      If lngX < lngTotalDates Then .MoveNext
8850                    End If
8860                  Next
8870                  .Close
8880                End With
8890                Set rst = Nothing
8900                Set qdf = Nothing
8910                .Close
8920              End With
8930              Set dbs = Nothing

8940              Beep
8950              msgResponse = MsgBox("Are you sure you want to delete the oldest " & CStr(lngTotalDates) & _
                    " date" & IIf(lngTotalDates > 1&, "s", vbNullString) & "?" & vbCrLf & vbCrLf & _
                    "Dates To Be Deleted: " & CStr(lngTotalDates) & vbCrLf & _
                    "Price Records: " & CStr(lngTotalRecs), vbQuestion + vbOKCancel + vbDefaultButton2, "Delete Previous Pricing History")
8960              If msgResponse <> vbOK Then
8970                blnContinue = False
8980              End If

8990            End If  ' ** ListCount.

9000            If blnContinue = True Then

9010              DoCmd.Hourglass True
9020              DoEvents

9030              Select Case blnDelAll
                  Case True
                    ' ** Set it to no date.
9040                .lbxHistoryDate = 0
9050                .lbxHistoryDate.Enabled = False
9060                .lbxHistoryDate_lbl.BackStyle = acBackStyleTransparent
9070                .lbxHistoryDate_lbl_box2.Visible = True
9080                .lbxHistoryDate.BorderColor = WIN_CLR_DISR
9090                strFilter = "[currentDate] = #01/01/1900#"
9100                If .chkShowHidden = False Then
9110                  strFilter = strFilter & " And [IsHid] = False"
9120                End If
9130                .frmAssetPricing_History_Sub.Form.Filter = strFilter
9140                .frmAssetPricing_History_Sub.Form.FilterOn = True
9150                .frmAssetPricing_History_Sub.Form.Form_Load  ' ** Form Procedure: frmAssetPricing_History_Sub.
9160              Case False
                    ' ** Set it to the newest date.
9170                If .lbxHistoryDate <> .lbxHistoryDate.Column(0, 0) Then
9180                  .lbxHistoryDate = .lbxHistoryDate.Column(0, 0)
9190                  strFilter = "[currentDate] = #" & .lbxHistoryDate & "#"
9200                  If .chkShowHidden = False Then
9210                    strFilter = strFilter & " And [IsHid] = False"
9220                  End If
9230                  .frmAssetPricing_History_Sub.Form.Filter = strFilter
9240                  .frmAssetPricing_History_Sub.Form.FilterOn = True
9250                  .frmAssetPricing_History_Sub.Form.Form_Load  ' ** Form Procedure: frmAssetPricing_History_Sub.
9260                End If
9270              End Select
9280              DoEvents

9290              Set dbs = CurrentDb
9300              With dbs
9310                Select Case blnDelAll
                    Case True
                      ' ** Empty tblPricing_MasterAsset_History.
9320                  Set qdf = .QueryDefs("qryPricing_History_10")
9330                  qdf.Execute dbFailOnError
9340                  Set qdf = Nothing
9350                Case False
9360                  For lngX = 0& To (lngTotalDates - 1&)
                        ' ** Delete tblPricing_MasterAsset_History, by specified [curdat].
9370                    Set qdf = .QueryDefs("qryPricing_History_11")
9380                    With qdf.Parameters
9390                      ![curdat] = arr_varDel(D_DAT, lngX)
9400                    End With
9410                    qdf.Execute dbFailOnError
9420                    Set qdf = Nothing
9430                  Next
9440                End Select
9450                .Close
9460              End With
9470              Set dbs = Nothing

9480              .lbxHistoryDate.Requery
9490              .frmAssetPricing_History_Sub.Form.Requery
9500              DoEvents
9510              Select Case blnDelAll
                  Case True
9520                .HistoryDateCnt = 0&
9530                .cmdRestore.Enabled = False
9540                .cmdPrintReport.Enabled = False
9550                .chkShowAssetTypeDesc.Enabled = False
                    '.chkShowAssetTypeDesc_lbl2a.ForeColor = WIN_CLR_DISF
                    '.chkShowAssetTypeDesc_lbl2b.ForeColor = WIN_CLR_DISF
                    'If .chkShowAssetTypeDesc_lbl2a.Visible = True Then
                    '  .chkShowAssetTypeDesc_lbl2a_dim_hi.Visible = True
                    '  .chkShowAssetTypeDesc_lbl2b_dim_hi.Visible = True
                    'End If
9560                .cmdPrintReport_raised_img_dis.Visible = True
9570                .cmdPrintReport_raised_img.Visible = False
9580                .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9590                .cmdPrintReport_raised_focus_img.Visible = False
9600                .cmdPrintReport_raised_focus_dots_img.Visible = False
9610                .cmdPrintReport_sunken_focus_dots_img.Visible = False
9620              Case False
9630                .lbxHistoryDate = .lbxHistoryDate.Column(0, 0)
9640                .HistoryDateCnt = .lbxHistoryDate.ListCount
9650                strFilter = "[currentDate] = #" & .lbxHistoryDate & "#"
9660                If .chkShowHidden = False Then
9670                  strFilter = strFilter & " And [IsHid] = False"
9680                End If
9690                .frmAssetPricing_History_Sub.Form.Filter = strFilter
9700                .frmAssetPricing_History_Sub.Form.FilterOn = True
9710                .frmAssetPricing_History_Sub.Form.Form_Load  ' ** Form Procedure: frmAssetPricing_History_Sub.
9720              End Select

9730              DoCmd.Hourglass False
9740              Beep

9750            End If  ' ** blnContinue.

9760          Else
9770            Beep
9780          End If
9790        Else
9800          Beep
9810        End If
9820      Else
            ' ** Should be disabled anyway!
9830        Beep
9840      End If
9850    End With

EXITP:
9860    Set rst = Nothing
9870    Set qdf = Nothing
9880    Set dbs = Nothing
9890    Exit Sub

ERRH:
9900    DoCmd.Hourglass False
9910    Select Case ERR.Number
        Case Else
9920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9930    End Select
9940    Resume EXITP

End Sub

Private Sub cmdHistoryDeleteOldest_KeyDown(KeyCode As Integer, Shift As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOldest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
10010   intShiftDown = (Shift And acShiftMask) > 0
10020   intAltDown = (Shift And acAltMask) > 0
10030   intCtrlDown = (Shift And acCtrlMask) > 0

10040   intRetVal = KeyCode

        ' ** Plain keys.
10050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10060     Select Case intRetVal
          Case vbKeyTab
10070       With Me
10080         intRetVal = 0
10090         .ckgDeleteDates_chkOne.SetFocus
10100       End With
10110     End Select
10120   End If

        ' ** Shift keys.
10130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10140     Select Case intRetVal
          Case vbKeyTab
10150       With Me
10160         intRetVal = 0
10170         .HistoryDeleteOldest.SetFocus
10180       End With
10190     End Select
10200   End If

EXITP:
10210   KeyCode = intRetVal
10220   Exit Sub

ERRH:
10230   intRetVal = 0
10240   Select Case ERR.Number
        Case Else
10250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10260   End Select
10270   Resume EXITP

End Sub

Private Sub cmdHistoryDeleteOldest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOldest_MouseMove"

10310   With Me
10320     If blnHistoryDeleteOldest_MouseDown = False Then
10330       Select Case blnHistoryDeleteOldest_Focus
            Case True
10340         .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = True
10350         .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
10360       Case False
10370         .cmdHistoryDeleteOldest_raised_focus_img.Visible = True
10380         .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
10390       End Select
10400       .cmdHistoryDeleteOldest_raised_img.Visible = False
10410       .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
10420       .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False
10430       .cmdHistoryDeleteOldest_raised_img_dis.Visible = False
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

Private Sub cmdHistoryDeleteOldest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOldest_MouseUp"

10610   With Me
10620     .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = True
10630     .cmdHistoryDeleteOldest_raised_img.Visible = False
10640     .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
10650     .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
10660     .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False
10670     .cmdHistoryDeleteOldest_raised_img_dis.Visible = False
10680     blnHistoryDeleteOldest_MouseDown = False
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

Private Sub cmdHistoryDeleteOldest_LostFocus()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOldest_LostFocus"

10810   With Me
10820     .cmdHistoryDeleteOldest_raised_img.Visible = True
10830     .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
10840     .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
10850     .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
10860     .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False
10870     .cmdHistoryDeleteOldest_raised_img_dis.Visible = False
10880     blnHistoryDeleteOldest_Focus = False
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

Private Sub ckgDeleteDates_chkOne_AfterUpdate()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDeleteDates_chkOne_AfterUpdate"

11010   With Me
11020     Select Case .ckgDeleteDates_chkOne
          Case True
11030       .ckgDeleteDates_chkOne_lbl.FontBold = True
11040       .ckgDeleteDates_chkOldest = False
11050       .ckgDeleteDates_chkOldest_lbl.FontBold = False
11060       .HistoryDelete_box.Visible = True
11070       .HistoryDelete_vline01.Visible = True
11080       .HistoryDelete_vline02.Visible = True
11090       .HistoryDelete_vline03.Visible = True
11100       .HistoryDelete_vline04.Visible = True
11110       .HistoryDelete_hline01.Visible = True
11120       .HistoryDelete_hline02.Visible = True
11130       .HistoryDeleteOldest.Visible = False
11140       .HistoryDeleteOldest_lbl2.Visible = False
11150       .cmdHistoryDeleteOldest_raised_img.Visible = False
11160       .cmdHistoryDeleteOldest_raised_semifocus_dots_img.Visible = False
11170       .cmdHistoryDeleteOldest_raised_focus_img.Visible = False
11180       .cmdHistoryDeleteOldest_raised_focus_dots_img.Visible = False
11190       .cmdHistoryDeleteOldest_sunken_focus_dots_img.Visible = False
11200       .cmdHistoryDeleteOldest_raised_img_dis.Visible = False
11210       .HistoryDeleteOne.Top = .HistoryDeleteOldest.Top
11220       .HistoryDeleteOne_lbl.Top = .HistoryDeleteOldest_lbl.Top
11230       .HistoryDeleteOne = .lbxHistoryDate
11240       .HistoryDeleteOne.Visible = True
11250       .cmdHistoryDeleteOne.Top = .cmdHistoryDeleteOldest.Top
11260       .cmdHistoryDeleteOne_raised_img.Top = .cmdHistoryDeleteOne.Top
11270       .cmdHistoryDeleteOne_raised_semifocus_dots_img.Top = .cmdHistoryDeleteOne.Top
11280       .cmdHistoryDeleteOne_raised_focus_img.Top = .cmdHistoryDeleteOne.Top
11290       .cmdHistoryDeleteOne_raised_focus_dots_img.Top = .cmdHistoryDeleteOne.Top
11300       .cmdHistoryDeleteOne_sunken_focus_dots_img.Top = .cmdHistoryDeleteOne.Top
11310       .cmdHistoryDeleteOne_raised_img_dis.Top = .cmdHistoryDeleteOne.Top
11320       .cmdHistoryDeleteOne.Visible = True
11330       .cmdHistoryDeleteOne.Enabled = True
11340       .cmdHistoryDeleteOne_raised_img.Visible = True
11350     Case False
11360       .ckgDeleteDates_chkOne_lbl.FontBold = False
11370       .HistoryDeleteOne.Visible = False
11380       .HistoryDeleteOne = Null
11390       .cmdHistoryDeleteOne.Visible = False
11400       .cmdHistoryDeleteOne.Enabled = False
11410       .cmdHistoryDeleteOne_raised_img.Visible = False
11420       .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = False
11430       .cmdHistoryDeleteOne_raised_focus_img.Visible = False
11440       .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = False
11450       .cmdHistoryDeleteOne_sunken_focus_dots_img.Visible = False
11460       .cmdHistoryDeleteOne_raised_img_dis.Visible = False
11470       .HistoryDelete_vline01.Visible = False
11480       .HistoryDelete_vline02.Visible = False
11490       .HistoryDelete_vline03.Visible = False
11500       .HistoryDelete_vline04.Visible = False
11510       .HistoryDelete_hline01.Visible = False
11520       .HistoryDelete_hline02.Visible = False
11530       .HistoryDelete_box.Visible = False
11540     End Select
11550   End With

EXITP:
11560   Exit Sub

ERRH:
11570   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
11580   Case Else
11590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11600   End Select
11610   Resume EXITP

End Sub

Private Sub ckgDeleteDates_chkOne_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "ckgDeleteDates_chkOne_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
11710   intShiftDown = (Shift And acShiftMask) > 0
11720   intAltDown = (Shift And acAltMask) > 0
11730   intCtrlDown = (Shift And acCtrlMask) > 0

11740   intRetVal = KeyCode

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11770       With Me
11780         intRetVal = 0
11790         If .ckgDeleteDates_chkOne = True And .cmdHistoryDeleteOne.Visible = True And .cmdHistoryDeleteOne.Enabled = True Then
11800           .cmdHistoryDeleteOne.SetFocus
11810         ElseIf .chkIncludeCurrency.Enabled = True Then
11820           .chkIncludeCurrency.SetFocus
11830         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
11840           .chkShowAssetNo.SetFocus
11850         Else
11860           .chkShowAssetTypeDesc.SetFocus
11870         End If
11880       End With
11890     End Select
11900   End If

        ' ** Shift keys.
11910   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11920     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11930       With Me
11940         intRetVal = 0
11950         If .ckgDeleteDates_chkOldest = True And .cmdHistoryDeleteOldest.Visible = True And .cmdHistoryDeleteOldest.Enabled = True Then
11960           .cmdHistoryDeleteOldest.SetFocus
11970         Else
11980           .ckgDeleteDates_chkOldest.SetFocus
11990         End If
12000       End With
12010     End Select
12020   End If

EXITP:
12030   KeyCode = intRetVal
12040   Exit Sub

ERRH:
12050   intRetVal = 0
12060   Select Case ERR.Number
        Case Else
12070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12080   End Select
12090   Resume EXITP

End Sub

Private Sub cmdHistoryDeleteOne_GotFocus()

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOne_GotFocus"

12110   With Me
12120     blnHistoryDeleteOne_Focus = True
12130     .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = True
12140     .cmdHistoryDeleteOne_raised_img.Visible = False
12150     .cmdHistoryDeleteOne_raised_focus_img.Visible = False
12160     .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = False
12170     .cmdHistoryDeleteOne_sunken_focus_dots_img.Visible = False
12180     .cmdHistoryDeleteOne_raised_img_dis.Visible = False
12190   End With

EXITP:
12200   Exit Sub

ERRH:
12210   Select Case ERR.Number
        Case Else
12220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12230   End Select
12240   Resume EXITP

End Sub

Private Sub cmdHistoryDeleteOne_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOne_MouseDown"

12310   With Me
12320     blnHistoryDeleteOne_MouseDown = True
12330     .cmdHistoryDeleteOne_sunken_focus_dots_img.Visible = True
12340     .cmdHistoryDeleteOne_raised_img.Visible = False
12350     .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = False
12360     .cmdHistoryDeleteOne_raised_focus_img.Visible = False
12370     .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = False
12380     .cmdHistoryDeleteOne_raised_img_dis.Visible = False
12390   End With

EXITP:
12400   Exit Sub

ERRH:
12410   Select Case ERR.Number
        Case Else
12420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12430   End Select
12440   Resume EXITP

End Sub

Public Sub cmdHistoryDeleteOne_Click()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOne_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim datCurDate As Date
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

12510   blnContinue = True

12520   With Me
12530     If .lbxHistoryDate.ListCount > 0 Then
12540       If IsNull(.lbxHistoryDate) = False Then
12550         datCurDate = .lbxHistoryDate
12560         If IsNull(.HistoryDeleteOne) = True Then
12570           .HistoryDeleteOne = datCurDate
12580         ElseIf .HistoryDeleteOne <> datCurDate Then
12590           .HistoryDeleteOne = datCurDate
12600         End If
12610         If .HistoryDeleteOne > 0 Then
12620           lngRecsCur = .frmAssetPricing_History_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_History_Sub.
12630           Beep
12640           msgResponse = MsgBox("Are you sure you want to delete this date's " & _
                  "pricing history?" & vbCrLf & vbCrLf & _
                  "Date To Be Deleted: " & Format(datCurDate, "mm/dd/yyyy") & vbCrLf & _
                  "Price Records: " & CStr(lngRecsCur), vbQuestion + vbOKCancel + vbDefaultButton2, "Delete Previous Pricing History")
12650           If msgResponse <> vbOK Then
12660             blnContinue = False
12670           End If
12680           If blnContinue = True Then
12690             DoCmd.Hourglass True
12700             DoEvents
12710             Set dbs = CurrentDb
12720             With dbs
                    ' ** Delete tblPricing_MasterAsset_History, by specified [curdat].
12730               Set qdf = .QueryDefs("qryPricing_History_11")
12740               With qdf.Parameters
12750                 ![curdat] = datCurDate
12760               End With
12770               qdf.Execute
12780               .Close
12790             End With
12800           End If  ' ** blnContinue.
12810         Else
12820           Beep
12830         End If
12840       Else
12850         Beep
12860       End If
12870     Else
            ' ** Should be disabled anyway!
12880       Beep
12890     End If
12900   End With

12910   DoCmd.Hourglass False

EXITP:
12920   Set qdf = Nothing
12930   Set dbs = Nothing
12940   Exit Sub

ERRH:
12950   DoCmd.Hourglass False
12960   Select Case ERR.Number
        Case Else
12970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12980   End Select
12990   Resume EXITP

End Sub

Private Sub cmdHistoryDeleteOne_KeyDown(KeyCode As Integer, Shift As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOne_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
13010   intShiftDown = (Shift And acShiftMask) > 0
13020   intAltDown = (Shift And acAltMask) > 0
13030   intCtrlDown = (Shift And acCtrlMask) > 0

13040   intRetVal = KeyCode

        ' ** Plain keys.
13050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13060     Select Case intRetVal
          Case vbKeyTab
13070       With Me
13080         intRetVal = 0
13090         If .chkIncludeCurrency.Enabled = True Then
13100           .chkIncludeCurrency.SetFocus
13110         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
13120           .chkShowAssetNo.SetFocus
13130         Else
13140           .chkShowAssetTypeDesc.SetFocus
13150         End If
13160       End With
13170     End Select
13180   End If

        ' ** Shift keys.
13190   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13200     Select Case intRetVal
          Case vbKeyTab
13210       With Me
13220         intRetVal = 0
13230         .ckgDeleteDates_chkOne.SetFocus
13240       End With
13250     End Select
13260   End If

EXITP:
13270   KeyCode = intRetVal
13280   Exit Sub

ERRH:
13290   intRetVal = 0
13300   Select Case ERR.Number
        Case Else
13310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13320   End Select
13330   Resume EXITP

End Sub

Private Sub cmdHistoryDeleteOne_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOne_MouseMove"

13410   With Me
13420     If blnHistoryDeleteOne_MouseDown = False Then
13430       Select Case blnHistoryDeleteOne_Focus
            Case True
13440         .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = True
13450         .cmdHistoryDeleteOne_raised_focus_img.Visible = False
13460       Case False
13470         .cmdHistoryDeleteOne_raised_focus_img.Visible = True
13480         .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = False
13490       End Select
13500       .cmdHistoryDeleteOne_raised_img.Visible = False
13510       .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = False
13520       .cmdHistoryDeleteOne_sunken_focus_dots_img.Visible = False
13530       .cmdHistoryDeleteOne_raised_img_dis.Visible = False
13540     End If
13550   End With

EXITP:
13560   Exit Sub

ERRH:
13570   Select Case ERR.Number
        Case Else
13580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13590   End Select
13600   Resume EXITP

End Sub

Private Sub cmdHistoryDeleteOne_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOne_MouseUp"

13710   With Me
13720     .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = True
13730     .cmdHistoryDeleteOne_raised_img.Visible = False
13740     .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = False
13750     .cmdHistoryDeleteOne_raised_focus_img.Visible = False
13760     .cmdHistoryDeleteOne_sunken_focus_dots_img.Visible = False
13770     .cmdHistoryDeleteOne_raised_img_dis.Visible = False
13780     blnHistoryDeleteOne_MouseDown = False
13790   End With

EXITP:
13800   Exit Sub

ERRH:
13810   Select Case ERR.Number
        Case Else
13820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13830   End Select
13840   Resume EXITP

End Sub

Private Sub cmdHistoryDeleteOne_LostFocus()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdHistoryDeleteOne_LostFocus"

13910   With Me
13920     .cmdHistoryDeleteOne_raised_img.Visible = True
13930     .cmdHistoryDeleteOne_raised_semifocus_dots_img.Visible = False
13940     .cmdHistoryDeleteOne_raised_focus_img.Visible = False
13950     .cmdHistoryDeleteOne_raised_focus_dots_img.Visible = False
13960     .cmdHistoryDeleteOne_sunken_focus_dots_img.Visible = False
13970     .cmdHistoryDeleteOne_raised_img_dis.Visible = False
13980     blnHistoryDeleteOne_Focus = False
13990   End With

EXITP:
14000   Exit Sub

ERRH:
14010   Select Case ERR.Number
        Case Else
14020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14030   End Select
14040   Resume EXITP

End Sub

Public Sub chkIncludeCurrency_AfterUpdate()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

14110   With Me
14120     Select Case .chkIncludeCurrency
          Case True
14130       .chkIncludeCurrency_lbl.FontBold = True
14140       If blnIsOpen = True And .chkShowAssetTypeDesc = True Then
              ' ** All present and accounted for.
14150       Else
14160         .frmAssetPricing_History_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAssetPricing_History_Sub.
14170         DoEvents
14180         IncludeCurrency True  ' ** Procedure: Below
14190       End If
14200     Case False
14210       .chkIncludeCurrency_lbl.FontBold = False
14220       .frmAssetPricing_History_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAssetPricing_History_Sub.
14230       DoEvents
14240       IncludeCurrency False  ' ** Procedure: Below.
14250     End Select
14260     .FocusHolder.SetFocus
14270     DoEvents
14280     .GoToPage 1, 0, 0
14290     DoEvents
14300     .frmAssetPricing_History_Sub.SetFocus
14310     .frmAssetPricing_History_Sub.Form.FocusHolder.SetFocus
14320     DoEvents
14330   End With

EXITP:
14340   Exit Sub

ERRH:
14350   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
14360   Case Else
14370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14380   End Select
14390   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
14410   intShiftDown = (Shift And acShiftMask) > 0
14420   intAltDown = (Shift And acAltMask) > 0
14430   intCtrlDown = (Shift And acCtrlMask) > 0

14440   intRetVal = KeyCode

        ' ** Plain keys.
14450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14460     Select Case intRetVal
          Case vbKeyTab
14470       With Me
14480         intRetVal = 0
14490         If .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
14500           .chkShowAssetNo.SetFocus
14510         Else
14520           .chkShowAssetTypeDesc.SetFocus
14530         End If
14540       End With
14550     End Select
14560   End If

        ' ** Shift keys.
14570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14580     Select Case intRetVal
          Case vbKeyTab
14590       With Me
14600         intRetVal = 0
14610         lngRecsCur = .frmAssetPricing_History_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_History_Sub.
14620         If lngRecsCur > 0& Then
14630           .frmAssetPricing_History_Sub.SetFocus
14640           .frmAssetPricing_History_Sub.Form.FocusHolder.SetFocus
14650         Else
14660           .cmdClose.SetFocus
14670         End If
14680       End With
14690     End Select
14700   End If

EXITP:
14710   KeyCode = intRetVal
14720   Exit Sub

ERRH:
14730   intRetVal = 0
14740   Select Case ERR.Number
        Case Else
14750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14760   End Select
14770   Resume EXITP

End Sub

Public Sub chkShowAssetNo_AfterUpdate()

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetNo_AfterUpdate"

14810   With Me
14820     Select Case .chkShowAssetNo
          Case True
14830       .chkShowAssetNo_lbl.FontBold = True
14840       .frmAssetPricing_History_Sub.Form.ShowAssetNo_Sub True  ' ** Form Procedure: frmAssetPricing_History_Sub.
14850       DoEvents
14860     Case False
14870       .chkShowAssetNo_lbl.FontBold = False
14880       .frmAssetPricing_History_Sub.Form.ShowAssetNo_Sub False  ' ** Form Procedure: frmAssetPricing_History_Sub.
14890       DoEvents
14900     End Select
14910     .FocusHolder.SetFocus
14920     DoEvents
14930     .frmAssetPricing_History_Sub.SetFocus
14940     .frmAssetPricing_History_Sub.Form.FocusHolder.SetFocus
14950     DoEvents
14960   End With

EXITP:
14970   Exit Sub

ERRH:
14980   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
14990   Case Else
15000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15010   End Select
15020   Resume EXITP

End Sub

Private Sub chkShowAssetNo_KeyDown(KeyCode As Integer, Shift As Integer)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
15110   intShiftDown = (Shift And acShiftMask) > 0
15120   intAltDown = (Shift And acAltMask) > 0
15130   intCtrlDown = (Shift And acCtrlMask) > 0

15140   intRetVal = KeyCode

        ' ** Plain keys.
15150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15160     Select Case intRetVal
          Case vbKeyTab
15170       With Me
15180         intRetVal = 0
15190         .chkShowAssetTypeDesc.SetFocus
15200       End With
15210     End Select
15220   End If

        ' ** Shift keys.
15230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15240     Select Case intRetVal
          Case vbKeyTab
15250       With Me
15260         intRetVal = 0
15270         If .chkIncludeCurrency.Enabled = True Then
15280           .chkIncludeCurrency.SetFocus
15290         Else
15300           lngRecsCur = .frmAssetPricing_History_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_History_Sub.
15310           If lngRecsCur > 0& Then
15320             .frmAssetPricing_History_Sub.SetFocus
15330             .frmAssetPricing_History_Sub.Form.FocusHolder.SetFocus
15340           Else
15350             .cmdClose.SetFocus
15360           End If
15370         End If
15380       End With
15390     End Select
15400   End If

EXITP:
15410   KeyCode = intRetVal
15420   Exit Sub

ERRH:
15430   intRetVal = 0
15440   Select Case ERR.Number
        Case Else
15450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15460   End Select
15470   Resume EXITP

End Sub

Private Sub chkShowHidden_AfterUpdate()

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowHidden_AfterUpdate"

        Dim strFilter As String
        Dim intPos01 As Integer
        Dim strTmp01 As String

15510   With Me
15520     strFilter = .frmAssetPricing_History_Sub.Form.Filter
15530     strTmp01 = vbNullString
15540     Select Case .chkShowHidden
          Case True
15550       .chkShowHidden_lbl.FontBold = True
15560       .chkShowHidden_lbl2_dim.FontBold = True
15570       .chkShowHidden_lbl2_dim_hi.FontBold = True
15580       If strFilter = vbNullString Then
              ' ** This shouldn't ever be the case, since the sub should only show 1 date at a time.
15590       Else
15600         If Left(strFilter, 7) = "[IsHid]" Then
15610           intPos01 = InStr(strFilter, " And ")
15620           If intPos01 > 0 Then
15630             strFilter = Trim(Mid(strFilter, (intPos01 + 4)))
15640           Else
                  ' ** Shouldn't happen!
15650           End If
15660         Else
15670           intPos01 = InStr(strFilter, "And [IsHid]")
15680           If intPos01 > 0 Then
15690             strFilter = Trim(Left(strFilter, (intPos01 - 1)))
15700           Else
15710             intPos01 = InStr(strFilter, "[IsHid]")
15720             If intPos01 > 0 Then
                    ' ** This is getting tiresome.
15730               strTmp01 = Trim(Left(strFilter, (intPos01 - 1)))
15740               If Right(strTmp01, 4) = " And" Then
15750                 strTmp01 = Trim(Left(strTmp01, intPos01))
15760               End If
15770               intPos01 = InStr(intPos01, strFilter, " And ")
15780               If intPos01 > 0 Then
15790                 strTmp01 = Trim(strTmp01 & Mid(strFilter, (intPos01 + 4)))
15800               Else
                      ' ** I give up!
15810               End If
15820             End If
15830             strFilter = strTmp01
15840           End If
15850         End If
15860       End If
15870       .frmAssetPricing_History_Sub.Form.Filter = strFilter
15880       .frmAssetPricing_History_Sub.Form.FilterOn = True
15890     Case False
15900       .chkShowHidden_lbl.FontBold = False
15910       .chkShowHidden_lbl2_dim.FontBold = False
15920       .chkShowHidden_lbl2_dim_hi.FontBold = False
15930       intPos01 = InStr(strFilter, "[IsHid]")
15940       If intPos01 = 0 Then
15950         strFilter = strFilter & " And [IsHid] = False"
15960       Else
15970         If InStr(strFilter, "[IsHid] = False") > 0 Then
                ' ** Leave it.
15980         Else
                ' ** Not sure what else it could be.
15990         End If
16000       End If
16010       .frmAssetPricing_History_Sub.Form.Filter = strFilter
16020       .frmAssetPricing_History_Sub.Form.FilterOn = True
16030     End Select
16040   End With

EXITP:
16050   Exit Sub

ERRH:
16060   Select Case ERR.Number
        Case Else
16070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16080   End Select
16090   Resume EXITP

End Sub

Private Sub chkShowHidden_KeyDown(KeyCode As Integer, Shift As Integer)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowHidden_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
16110   intShiftDown = (Shift And acShiftMask) > 0
16120   intAltDown = (Shift And acAltMask) > 0
16130   intCtrlDown = (Shift And acCtrlMask) > 0

16140   intRetVal = KeyCode

        ' ** Plain keys.
16150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16160     Select Case intRetVal
          Case vbKeyTab
16170       With Me
16180         intRetVal = 0
16190         If .cmdPrintReport.Enabled = True Then
16200           .cmdPrintReport.SetFocus
16210         ElseIf .cmdRestore.Enabled = True Then
16220           .cmdRestore.SetFocus
16230         Else
16240           .cmdClose.SetFocus
16250         End If
16260       End With
16270     End Select
16280   End If

        ' ** Shift keys.
16290   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16300     Select Case intRetVal
          Case vbKeyTab
16310       With Me
16320         intRetVal = 0
16330         .chkShowAssetTypeDesc.SetFocus
16340       End With
16350     End Select
16360   End If

EXITP:
16370   KeyCode = intRetVal
16380   Exit Sub

ERRH:
16390   intRetVal = 0
16400   Select Case ERR.Number
        Case Else
16410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16420   End Select
16430   Resume EXITP

End Sub

Public Sub chkShowAssetTypeDesc_AfterUpdate()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetTypeDesc_AfterUpdate"

16510   With Me
16520     Select Case .chkShowAssetTypeDesc
          Case True
16530       .chkShowAssetTypeDesc_lbl.FontBold = True
16540       If blnIsOpen = True And .chkIncludeCurrency = True Then
              ' ** All's right with the world.
16550       Else
16560         .frmAssetPricing_History_Sub.Form.ShowAssetTypeDesc_Sub True  ' ** Form Procedure: frmAssetPricing_History_Sub.
16570         DoEvents
16580         ShowAssetTypeDesc True  ' ** Procedure: Below.
16590       End If
16600     Case False
16610       .chkShowAssetTypeDesc_lbl.FontBold = False
16620       .frmAssetPricing_History_Sub.Form.ShowAssetTypeDesc_Sub False  ' ** Form Procedure: frmAssetPricing_History_Sub.
16630       DoEvents
16640       ShowAssetTypeDesc False  ' ** Procedure: Below.
16650     End Select
16660     .FocusHolder.SetFocus
16670     DoEvents
16680     .GoToPage 1, 0, 0
16690     DoEvents
16700     .frmAssetPricing_History_Sub.SetFocus
16710     .frmAssetPricing_History_Sub.Form.FocusHolder.SetFocus
16720     DoEvents
16730   End With

EXITP:
16740   Exit Sub

ERRH:
16750   Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form loading, ignore.
16760   Case Else
16770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16780   End Select
16790   Resume EXITP

End Sub

Private Sub chkShowAssetTypeDesc_KeyDown(KeyCode As Integer, Shift As Integer)

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowAssetTypeDesc_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
16810   intShiftDown = (Shift And acShiftMask) > 0
16820   intAltDown = (Shift And acAltMask) > 0
16830   intCtrlDown = (Shift And acCtrlMask) > 0

16840   intRetVal = KeyCode

        ' ** Plain keys.
16850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16860     Select Case intRetVal
          Case vbKeyTab
16870       With Me
16880         intRetVal = 0
16890         If .cmdPrintReport.Enabled = True Then
16900           .cmdPrintReport.SetFocus
16910         ElseIf .cmdRestore.Enabled = True Then
16920           .cmdRestore.SetFocus
16930         Else
16940           .cmdClose.SetFocus
16950         End If
16960       End With
16970     End Select
16980   End If

        ' ** Shift keys.
16990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17000     Select Case intRetVal
          Case vbKeyTab
17010       With Me
17020         intRetVal = 0
17030         If .chkShowHidden.Enabled = True Then
17040           .chkShowHidden.SetFocus
17050         ElseIf .chkShowAssetNo.Visible = True And .chkShowAssetNo.Enabled = True Then
17060           .chkShowAssetNo.SetFocus
17070         ElseIf .chkIncludeCurrency.Enabled = True Then
17080           .chkIncludeCurrency.SetFocus
17090         Else
17100           lngRecsCur = .frmAssetPricing_History_Sub.Form.RecCnt  ' ** Form Function: frmAssetPricing_History_Sub.
17110           If lngRecsCur > 0& Then
17120             .frmAssetPricing_History_Sub.SetFocus
17130             .frmAssetPricing_History_Sub.Form.FocusHolder.SetFocus
17140           Else
17150             .cmdClose.SetFocus
17160           End If
17170         End If
17180       End With
17190     End Select
17200   End If

EXITP:
17210   KeyCode = intRetVal
17220   Exit Sub

ERRH:
17230   intRetVal = 0
17240   Select Case ERR.Number
        Case Else
17250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17260   End Select
17270   Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

17310   With Me
17320     blnPrintReport_Focus = True
17330     .cmdPrintReport_raised_semifocus_dots_img.Visible = True
17340     .cmdPrintReport_raised_img.Visible = False
17350     .cmdPrintReport_raised_focus_img.Visible = False
17360     .cmdPrintReport_raised_focus_dots_img.Visible = False
17370     .cmdPrintReport_sunken_focus_dots_img.Visible = False
17380     .cmdPrintReport_raised_img_dis.Visible = False
17390   End With

EXITP:
17400   Exit Sub

ERRH:
17410   Select Case ERR.Number
        Case Else
17420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17430   End Select
17440   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

17510   With Me
17520     blnPrintReport_MouseDown = True
17530     .cmdPrintReport_sunken_focus_dots_img.Visible = True
17540     .cmdPrintReport_raised_img.Visible = False
17550     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
17560     .cmdPrintReport_raised_focus_img.Visible = False
17570     .cmdPrintReport_raised_focus_dots_img.Visible = False
17580     .cmdPrintReport_raised_img_dis.Visible = False
17590   End With

EXITP:
17600   Exit Sub

ERRH:
17610   Select Case ERR.Number
        Case Else
17620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17630   End Select
17640   Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim datCurrentDate As Date
        Dim strDocName As String, strFilter As String

17710   With Me
17720     If IsNull(.lbxHistoryDate) = False Then
17730       datCurrentDate = CDate(.lbxHistoryDate)
17740       strDocName = "rptAssetPricing_History"
17750       strFilter = "[currentDate] = #" & Format(datCurrentDate, "mm/dd/yyyy") & "#"
17760       If .chkShowHidden = False Then
17770         strFilter = strFilter & " And [IsHid] = False"
17780       End If
17790       If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
17800         DoCmd.OpenReport strDocName, acViewPreview, , strFilter, , _
                THIS_NAME & "~" & Format(datCurrentDate, "mm/dd/yyyy")
17810         .Visible = False
17820         Forms("frmAssetPricing").Visible = False
17830         DoCmd.SelectObject acReport, strDocName, False
17840         DoCmd.Maximize
17850         DoCmd.RunCommand acCmdFitToWindow
17860       Else
              '##GTR_Ref: rptAssetPricing_History
17870         DoCmd.OpenReport strDocName, acViewNormal, , strFilter, , _
                THIS_NAME & "~" & Format(datCurrentDate, "mm/dd/yyyy")
17880       End If
17890     End If
17900   End With

EXITP:
17910   Exit Sub

ERRH:
17920   Forms("frmAssetPricing").Visible = True
17930   Me.Visible = True
17940   DoCmd.Restore
17950   Select Case ERR.Number
        Case Else
17960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17970   End Select
17980   Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
18010   intShiftDown = (Shift And acShiftMask) > 0
18020   intAltDown = (Shift And acAltMask) > 0
18030   intCtrlDown = (Shift And acCtrlMask) > 0

18040   intRetVal = KeyCode

        ' ** Plain keys.
18050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18060     Select Case intRetVal
          Case vbKeyTab
18070       With Me
18080         intRetVal = 0
18090         If .cmdRestore.Enabled = True Then
18100           .cmdRestore.SetFocus
18110         Else
18120           .cmdClose.SetFocus
18130         End If
18140       End With
18150     End Select
18160   End If

        ' ** Shift keys.
18170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18180     Select Case intRetVal
          Case vbKeyTab
18190       With Me
18200         intRetVal = 0
18210         .chkShowAssetTypeDesc.SetFocus
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

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

18410   With Me
18420     If blnPrintReport_MouseDown = False Then
18430       Select Case blnPrintReport_Focus
            Case True
18440         .cmdPrintReport_raised_focus_dots_img.Visible = True
18450         .cmdPrintReport_raised_focus_img.Visible = False
18460       Case False
18470         .cmdPrintReport_raised_focus_img.Visible = True
18480         .cmdPrintReport_raised_focus_dots_img.Visible = False
18490       End Select
18500       .cmdPrintReport_raised_img.Visible = False
18510       .cmdPrintReport_raised_semifocus_dots_img.Visible = False
18520       .cmdPrintReport_sunken_focus_dots_img.Visible = False
18530       .cmdPrintReport_raised_img_dis.Visible = False
18540     End If
18550   End With

EXITP:
18560   Exit Sub

ERRH:
18570   Select Case ERR.Number
        Case Else
18580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18590   End Select
18600   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

18710   With Me
18720     .cmdPrintReport_raised_focus_dots_img.Visible = True
18730     .cmdPrintReport_raised_img.Visible = False
18740     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
18750     .cmdPrintReport_raised_focus_img.Visible = False
18760     .cmdPrintReport_sunken_focus_dots_img.Visible = False
18770     .cmdPrintReport_raised_img_dis.Visible = False
18780     blnPrintReport_MouseDown = False
18790   End With

EXITP:
18800   Exit Sub

ERRH:
18810   Select Case ERR.Number
        Case Else
18820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18830   End Select
18840   Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

18910   With Me
18920     .cmdPrintReport_raised_img.Visible = True
18930     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
18940     .cmdPrintReport_raised_focus_img.Visible = False
18950     .cmdPrintReport_raised_focus_dots_img.Visible = False
18960     .cmdPrintReport_sunken_focus_dots_img.Visible = False
18970     .cmdPrintReport_raised_img_dis.Visible = False
18980     blnPrintReport_Focus = False
18990   End With

EXITP:
19000   Exit Sub

ERRH:
19010   Select Case ERR.Number
        Case Else
19020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19030   End Select
19040   Resume EXITP

End Sub

Private Sub cmdRestore_Click()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRestore_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim datCurDate As Date
        Dim Cancel As Integer

19110   With Me
19120     If IsNull(.lbxHistoryDate) = False Then
19130       If IsDate(.lbxHistoryDate) = True Then
19140         If MsgBox("Are you sure you want to replace the" & vbCrLf & _
                  "current Asset Pricing values with these?", vbQuestion + vbYesNo + vbDefaultButton2, "Restore Asset Pricing") = vbYes Then
19150           datCurDate = CDate(.lbxHistoryDate)
19160           Set dbs = CurrentDb
19170           With dbs
                  ' ** Update tblPricing_MasterAsset from tblPricing_MasterAsset_History, by specified [curdat].
19180             Set qdf = .QueryDefs("qryPricing_History_08")  ' ** Includes AssetPricing_Changed = True.
19190             With qdf.Parameters
19200               ![curdat] = datCurDate
19210             End With
19220             qdf.Execute dbFailOnError
                  ' ** Update tblPricing_MasterAsset, by specified [curdat].
19230             Set qdf = .QueryDefs("qryPricing_History_08a")
19240             With qdf.Parameters
19250               ![curdat] = datCurDate
19260             End With
19270             qdf.Execute  ' ** If they're all updated, this won't hit any new ones.
19280             .Close
19290           End With
19300           With Forms("frmAssetPricing")
19310             .frmAssetPricing_Sub.Form.Requery
19320             .chkUpdated = True
19330             .chkRestored = True
19340             .Date_New = CDate(Me.lbxHistoryDate)
19350             .opgShow = .opgShow_optAll.OptionValue
19360             .opgShow_AfterUpdate  ' ** Form Procedure: frmAssetPricing.
19370             Cancel = 0
19380             .Date_New_Exit Cancel  ' ** Form Procedure: frmAssetPricing.
19390             DoEvents
19400             If .cmdUpdate.Enabled = False Then
19410               .cmdUpdate.Enabled = True
19420             End If
                  ' ** If they change anything, it will replace the original!
19430             .frmAssetPricing_Sub.Form.yield_entry.Locked = True
19440             .frmAssetPricing_Sub.Form.marketvalue.Locked = True
19450           End With
19460           DoEvents
19470           MsgBox "Values restored.", vbInformation + vbOKOnly, ("Values Restored" & Space(40))
19480           cmdClose_Click  ' ** Procedure: Above.
19490         End If
19500       Else
19510         MsgBox "Please select a date from the " & Left(.lbxHistoryDate_lbl.Caption, (Len(.lbxHistoryDate_lbl.Caption) - 1)) & " List.", _
                vbInformation + vbOKOnly, "Entry Required"
19520       End If
19530     Else
19540       MsgBox "Please select a date from the " & Left(.lbxHistoryDate_lbl.Caption, (Len(.lbxHistoryDate_lbl.Caption) - 1)) & " List.", _
              vbInformation + vbOKOnly, "Entry Required"
19550     End If
19560   End With

EXITP:
19570   Set qdf = Nothing
19580   Set dbs = Nothing
19590   Exit Sub

ERRH:
19600   Select Case ERR.Number
        Case Else
19610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19620   End Select
19630   Resume EXITP

End Sub

Private Sub cmdRestore_KeyDown(KeyCode As Integer, Shift As Integer)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdRestore_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
19710   intShiftDown = (Shift And acShiftMask) > 0
19720   intAltDown = (Shift And acAltMask) > 0
19730   intCtrlDown = (Shift And acCtrlMask) > 0

19740   intRetVal = KeyCode

        ' ** Plain keys.
19750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19760     Select Case intRetVal
          Case vbKeyTab
19770       With Me
19780         intRetVal = 0
19790         .cmdClose.SetFocus
19800       End With
19810     End Select
19820   End If

        ' ** Shift keys.
19830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19840     Select Case intRetVal
          Case vbKeyTab
19850       With Me
19860         intRetVal = 0
19870         If .cmdPrintReport.Enabled = True Then
19880           .cmdPrintReport.SetFocus
19890         Else
19900           .chkShowAssetTypeDesc.SetFocus
19910         End If
19920       End With
19930     End Select
19940   End If

EXITP:
19950   KeyCode = intRetVal
19960   Exit Sub

ERRH:
19970   intRetVal = 0
19980   Select Case ERR.Number
        Case Else
19990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20000   End Select
20010   Resume EXITP

End Sub

Private Sub ShowAssetTypeDesc(blnShow As Boolean)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowAssetTypeDesc"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngAddedWidth As Long, lngNewFrmWidth As Long, lngNewFrmLeft As Long
        Dim lngNewForm_Width As Long
        Dim lngTmp01 As Long

20110   With Me

20120     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
20130       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
20140     End If

          ' ** Variables are fed empty, then populated ByRef.
20150     GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

20160     lngAddedWidth = lngTypeDesc_Width  ' ** Or minused width.

20170     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
20180     lngMonitorNum = 1&: lngTmp01 = 0&
20190     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
20200     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

20210     Select Case .chkIncludeCurrency
          Case True
20220       Select Case blnShow
            Case True
              ' ** If both are showing, it's the opening width.
20230         lngNewFrmWidth = lngFrm_Width  ' ** The original width.
20240       Case False
              ' ** If type desc isn't showing, it's just minus that.
20250         lngNewFrmWidth = (lngFrm_Width - lngAddedWidth)
20260       End Select
20270     Case False
20280       Select Case blnShow
            Case True
              ' ** If currency isn't showing, it's just minus that.
20290         lngNewFrmWidth = (lngFrm_Width - lngCurrID_Width)
20300       Case False
              ' ** If neither are showing, minus them both.
20310         lngNewFrmWidth = ((lngFrm_Width - lngTypeDesc_Width) - lngCurrID_Width)
20320       End Select
20330     End Select

20340     Select Case blnShow
          Case True
20350       If (blnIsOpen = False And lngWidth <> lngNewFrmWidth) Or (blnIsOpen = True) Then

20360         Select Case .chkIncludeCurrency
              Case True
20370           lngNewFrmLeft = lngFrm_Left  ' ** The original left.
20380         Case False
20390           lngNewFrmLeft = (lngFrm_Left + (lngCurrID_Width / 2))
20400         End Select
20410         If lngLeft <> lngNewFrmLeft And lngWidth <> lngNewFrmWidth Then
20420           DoCmd.SelectObject acForm, THIS_NAME, False
20430           If lngMonitorNum = 1& Then lngTmp01 = (lngFrm_Top + lngFrmTop_Adj)
20440           DoCmd.MoveSize (lngNewFrmLeft + lngFrmLeft_Adj), lngTmp01, lngNewFrmWidth, lngFrm_Height  '(lngFrm_Top + lngFrmTop_Adj)
20450           If lngMonitorNum > 1& Then
20460             LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
20470           End If
20480         End If

20490         Select Case .chkIncludeCurrency
              Case True
20500           lngNewForm_Width = lngForm_Width
20510           .cmdClose.Left = lngClose_Left
20520           .cmdRestore.Left = lngRestore_Left
20530         Case False
20540           lngNewForm_Width = (lngForm_Width - lngCurrID_Width)
20550           .cmdClose.Left = (lngClose_Left - lngCurrID_Width)
20560           .cmdRestore.Left = (lngRestore_Left - lngCurrID_Width)
20570         End Select
20580         .Width = lngNewForm_Width

20590         .Header_vline01.Left = lngNewForm_Width
20600         .Header_vline02.Left = lngNewForm_Width
20610         .Footer_vline01.Left = lngNewForm_Width
20620         .Footer_vline02.Left = lngNewForm_Width
20630         .Header_hline01.Width = lngNewForm_Width
20640         .Header_hline02.Width = lngNewForm_Width
20650         .Footer_hline01.Width = lngNewForm_Width
20660         .Footer_hline02.Width = lngNewForm_Width

20670         Select Case .chkIncludeCurrency
              Case True
20680           .frmAssetPricing_History_Sub.Width = lngSub_Width
20690           .frmAssetPricing_History_Sub_box.Width = (.frmAssetPricing_History_Sub.Width + (2& * lngTpp))
20700         Case False
20710           .frmAssetPricing_History_Sub.Width = (lngSub_Width - lngCurrID_Width)
20720           .frmAssetPricing_History_Sub_box.Width = (.frmAssetPricing_History_Sub.Width + (2& * lngTpp))
20730         End Select
20740         .Nav_box01.Width = .frmAssetPricing_History_Sub.Width
20750         .Nav_hline03.Width = .frmAssetPricing_History_Sub.Width

20760       End If
20770     Case False
20780       If (blnIsOpen = False And lngWidth <> lngNewFrmWidth) Or (blnIsOpen = True) Then

20790         Select Case .chkIncludeCurrency
              Case True
20800           lngNewForm_Width = (lngForm_Width - lngTypeDesc_Width)
20810           .cmdClose.Left = (lngClose_Left - lngTypeDesc_Width)
20820           .cmdRestore.Left = (lngRestore_Left - lngTypeDesc_Width)
20830         Case False
20840           lngNewForm_Width = ((lngForm_Width - lngTypeDesc_Width) - lngCurrID_Width)
20850           .cmdClose.Left = ((lngClose_Left - lngTypeDesc_Width) - lngCurrID_Width)
20860           .cmdRestore.Left = ((lngRestore_Left - lngTypeDesc_Width) - lngCurrID_Width)
20870         End Select

20880         .Header_vline01.Left = lngNewForm_Width
20890         .Header_vline02.Left = lngNewForm_Width
20900         .Footer_vline01.Left = lngNewForm_Width
20910         .Footer_vline02.Left = lngNewForm_Width
20920         .Header_hline01.Width = lngNewForm_Width
20930         .Header_hline02.Width = lngNewForm_Width
20940         .Footer_hline01.Width = lngNewForm_Width
20950         .Footer_hline02.Width = lngNewForm_Width

20960         Select Case .chkIncludeCurrency
              Case True
20970           .frmAssetPricing_History_Sub.Width = (lngSub_Width - lngTypeDesc_Width)
20980           .frmAssetPricing_History_Sub_box.Width = (.frmAssetPricing_History_Sub.Width + (2& * lngTpp))
20990         Case False
21000           .frmAssetPricing_History_Sub.Width = ((lngSub_Width - lngTypeDesc_Width) - lngCurrID_Width)
21010           .frmAssetPricing_History_Sub_box.Width = (.frmAssetPricing_History_Sub.Width + (2& * lngTpp))
21020         End Select
21030         .Nav_box01.Width = .frmAssetPricing_History_Sub.Width
21040         .Nav_hline03.Width = .frmAssetPricing_History_Sub.Width

21050         .Width = lngNewForm_Width
21060         Select Case .chkIncludeCurrency
              Case True
21070           lngNewFrmLeft = (lngFrm_Left + (lngAddedWidth / 2))
21080         Case False
21090           lngNewFrmLeft = (lngFrm_Left + ((lngTypeDesc_Width + lngCurrID_Width) / 2))
21100         End Select
21110         If lngLeft <> lngNewFrmLeft And lngWidth <> lngNewFrmWidth Then
21120           DoCmd.SelectObject acForm, THIS_NAME, False
21130           If lngMonitorNum = 1& Then lngTmp01 = (lngFrm_Top + lngFrmTop_Adj)
21140           DoCmd.MoveSize (lngNewFrmLeft + lngFrmLeft_Adj), lngTmp01, lngNewFrmWidth, lngFrm_Height  '(lngFrm_Top + lngFrmTop_Adj)
21150           If lngMonitorNum > 1& Then
21160             LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
21170           End If
21180         End If

21190       End If
21200     End Select

21210   End With

EXITP:
21220   Exit Sub

ERRH:
21230   Select Case ERR.Number
        Case Else
21240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21250   End Select
21260   Resume EXITP

End Sub

Private Sub IncludeCurrency(blnShow As Boolean)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngAddedWidth As Long, lngNewFrmWidth As Long, lngNewFrmLeft As Long
        Dim lngNewForm_Width As Long
        Dim lngTmp01 As Long

21310   With Me

21320     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
21330       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
21340     End If

          ' ** Variables are fed empty, then populated ByRef.
21350     GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

21360     lngAddedWidth = lngCurrID_Width  ' ** Or minused width. (As in minus/plus.)

21370     Select Case .chkShowAssetTypeDesc
          Case True
21380       Select Case blnShow
            Case True
              ' ** If both are showing, it's the opening width.
21390         lngNewFrmWidth = lngFrm_Width  ' ** The original width.
21400       Case False
              ' ** If currency isn't showing, it's just minus curr_code.
21410         lngNewFrmWidth = (lngFrm_Width - lngAddedWidth)
21420       End Select
21430     Case False
21440       Select Case blnShow
            Case True
              ' ** If type desc isn't showing, it's just minus that.
21450         lngNewFrmWidth = (lngFrm_Width - lngTypeDesc_Width)
21460       Case False
              ' ** If neither are showing, minus them both.
21470         lngNewFrmWidth = ((lngFrm_Width - lngCurrID_Width) - lngTypeDesc_Width)
21480       End Select
21490     End Select

21500     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
21510     lngMonitorNum = 1&: lngTmp01 = 0&
21520     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
21530     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

21540     Select Case blnShow
          Case True
21550       If (blnIsOpen = False And lngWidth <> lngNewFrmWidth) Or (blnIsOpen = True) Then

21560         Select Case .chkShowAssetTypeDesc
              Case True
21570           lngNewFrmLeft = lngFrm_Left  ' ** The original left.
21580         Case False
21590           lngNewFrmLeft = (lngFrm_Left + (lngTypeDesc_Width / 2))
21600         End Select
21610         If lngLeft <> lngNewFrmLeft And lngWidth <> lngNewFrmWidth Then
21620           DoCmd.SelectObject acForm, THIS_NAME, False
21630           If lngMonitorNum = 1& Then lngTmp01 = (lngFrm_Top + lngFrmTop_Adj)
21640           DoCmd.MoveSize (lngNewFrmLeft + lngFrmLeft_Adj), lngTmp01, lngNewFrmWidth, lngFrm_Height  '(lngFrm_Top + lngFrmTop_Adj)
21650           If lngMonitorNum > 1& Then
21660             LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
21670           End If
21680         End If

21690         Select Case .chkShowAssetTypeDesc
              Case True
21700           lngNewForm_Width = lngForm_Width
21710           .Width = lngNewForm_Width
21720           .cmdClose.Left = lngClose_Left
21730           .cmdRestore.Left = lngRestore_Left
21740         Case False
21750           lngNewForm_Width = (lngForm_Width - lngTypeDesc_Width)
21760           .Width = lngNewForm_Width
21770           .cmdClose.Left = (lngClose_Left - lngTypeDesc_Width)
21780           .cmdRestore.Left = (lngRestore_Left - lngTypeDesc_Width)
21790         End Select

21800         .Header_vline01.Left = lngNewForm_Width
21810         .Header_vline02.Left = lngNewForm_Width
21820         .Footer_vline01.Left = lngNewForm_Width
21830         .Footer_vline02.Left = lngNewForm_Width
21840         .Header_hline01.Width = lngNewForm_Width
21850         .Header_hline02.Width = lngNewForm_Width
21860         .Footer_hline01.Width = lngNewForm_Width
21870         .Footer_hline02.Width = lngNewForm_Width

21880         Select Case .chkShowAssetTypeDesc
              Case True
21890           .frmAssetPricing_History_Sub.Width = lngSub_Width
21900           .frmAssetPricing_History_Sub_box.Width = (.frmAssetPricing_History_Sub.Width + (2& * lngTpp))
21910         Case False
21920           .frmAssetPricing_History_Sub.Width = (lngSub_Width - lngTypeDesc_Width)
21930           .frmAssetPricing_History_Sub_box.Width = (.frmAssetPricing_History_Sub.Width + (2& * lngTpp))
21940         End Select
21950         .Nav_box01.Width = .frmAssetPricing_History_Sub.Width
21960         .Nav_hline03.Width = .frmAssetPricing_History_Sub.Width

21970       End If
21980     Case False
21990       If (blnIsOpen = False And lngWidth <> lngNewFrmWidth) Or (blnIsOpen = True) Then

22000         Select Case .chkShowAssetTypeDesc
              Case True
22010           lngNewForm_Width = (lngForm_Width - lngCurrID_Width)
22020           .cmdClose.Left = (lngClose_Left - lngCurrID_Width)
22030           .cmdRestore.Left = (lngRestore_Left - lngCurrID_Width)
22040         Case False
22050           lngNewForm_Width = ((lngForm_Width - lngTypeDesc_Width) - lngCurrID_Width)
22060           .cmdClose.Left = ((lngClose_Left - lngTypeDesc_Width) - lngCurrID_Width)
22070           .cmdRestore.Left = ((lngRestore_Left - lngTypeDesc_Width) - lngCurrID_Width)
22080         End Select

22090         .Header_vline01.Left = lngNewForm_Width
22100         .Header_vline02.Left = lngNewForm_Width
22110         .Footer_vline01.Left = lngNewForm_Width
22120         .Footer_vline02.Left = lngNewForm_Width
22130         .Header_hline01.Width = lngNewForm_Width
22140         .Header_hline02.Width = lngNewForm_Width
22150         .Footer_hline01.Width = lngNewForm_Width
22160         .Footer_hline02.Width = lngNewForm_Width

22170         Select Case .chkShowAssetTypeDesc
              Case True
22180           .frmAssetPricing_History_Sub.Width = (lngSub_Width - lngCurrID_Width)
22190           .frmAssetPricing_History_Sub_box.Width = (.frmAssetPricing_History_Sub.Width + (2& * lngTpp))
22200         Case False
22210           .frmAssetPricing_History_Sub.Width = ((lngSub_Width - lngTypeDesc_Width) - lngCurrID_Width)
22220           .frmAssetPricing_History_Sub_box.Width = (.frmAssetPricing_History_Sub.Width + (2& * lngTpp))
22230         End Select
22240         .Nav_box01.Width = .frmAssetPricing_History_Sub.Width
22250         .Nav_hline03.Width = .frmAssetPricing_History_Sub.Width

22260         .Width = lngNewForm_Width
22270         Select Case .chkShowAssetTypeDesc
              Case True
22280           lngNewFrmLeft = (lngFrm_Left + (lngAddedWidth / 2))
22290         Case False
22300           lngNewFrmLeft = (lngFrm_Left + ((lngTypeDesc_Width + lngCurrID_Width) / 2))
22310         End Select
22320         If lngLeft <> lngNewFrmLeft And lngWidth <> lngNewFrmWidth Then
22330           DoCmd.SelectObject acForm, THIS_NAME, False
22340           If lngMonitorNum = 1& Then lngTmp01 = (lngFrm_Top + lngFrmTop_Adj)
22350           DoCmd.MoveSize (lngNewFrmLeft + lngFrmLeft_Adj), lngTmp01, lngNewFrmWidth, lngFrm_Height  '(lngFrm_Top + lngFrmTop_Adj)
22360           If lngMonitorNum > 1& Then
22370             LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
22380           End If
22390         End If

22400       End If
22410     End Select

22420   End With

EXITP:
22430   Exit Sub

ERRH:
22440   Select Case ERR.Number
        Case Else
22450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22460   End Select
22470   Resume EXITP

End Sub

Private Sub ShowHiddenAssets()

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowHiddenAssets"

        Dim lngHids As Long
        Dim blnShow As Boolean
        Dim varTmp00 As Variant, varTmp01 As Variant

22510   With Me

22520     DoCmd.Hourglass True
22530     DoEvents

          ' ** qryAssetPricing_History_07_01 (MasterAsset, with IsHid), grouped, just IsHid = True, with cnt_ast.
22540     varTmp00 = DLookup("[cnt_ast]", "qryAssetPricing_History_07_02")
22550     varTmp00 = ZeroIfNull(varTmp00)  ' ** Module Function: modStringFuncs
          ' ** qryAssetPricing_History_07_05 (tblPricing_MasterAsset_History, with IsHid, by
          ' ** specified GlobalVarGet('gdatMarketDate')), grouped, just IsHid = True, with cnt_ast.
22560     varTmp01 = DLookup("[cnt_ast]", "qryAssetPricing_History_07_06")
22570     varTmp01 = ZeroIfNull(varTmp01)  ' ** Module Function: modStringFuncs

22580     If varTmp00 = 0 And varTmp01 = 0 Then
22590       blnShow = False
22600     Else
22610 On Error Resume Next
22620       varTmp00 = .frmAssetPricing_History_Sub.Form.IsHidSum
22630       If ERR.Number = 0 Then
22640 On Error GoTo ERRH
22650         Select Case IsNull(varTmp00)
              Case True
22660           blnShow = False
22670           lngHids = 0&
22680         Case False
22690           lngHids = varTmp00
22700           If lngHids > 0& Then
22710             blnShow = True
22720           Else
22730             blnShow = False
22740           End If
22750         End Select
22760       Else
22770 On Error GoTo ERRH
              ' ** 2427  ' ** The expression you entered has no value.
              ' ** 2455  ' ** You entered an expression that has an invalid reference to the property Form/Report.
22780         blnShow = False
22790         lngHids = 0&
22800       End If
22810       If blnShow = False And varTmp01 > 0 Then
22820         blnShow = True
22830       End If
22840     End If

22850     Select Case blnShow
          Case True
22860       .chkShowHidden.Enabled = True
22870       .chkShowHidden_lbl.Visible = True
22880       .chkShowHidden_lbl2_dim.Visible = False
22890       .chkShowHidden_lbl2_dim_hi.Visible = False
22900     Case False
22910       .chkShowHidden.Enabled = False
22920       .chkShowHidden_lbl2_dim.Visible = True
22930       .chkShowHidden_lbl2_dim_hi.Visible = True
22940       .chkShowHidden_lbl.Visible = False
22950     End Select

22960     DoCmd.Hourglass False

22970   End With

EXITP:
22980   Exit Sub

ERRH:
22990   DoCmd.Hourglass False
23000   Select Case ERR.Number
        Case Else
23010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23020   End Select
23030   Resume EXITP

End Sub
