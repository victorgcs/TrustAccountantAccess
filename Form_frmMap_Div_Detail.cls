VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Div_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Div_Detail"

'VGC 08/30/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   IncludeCurrency_Sub()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Dividend}
' **   OK:               O {cmdOK}
' **   Print Report:     P {cmdPrintReport}
' **   Reinvest:         R {cmdReinvest}
' **   Tax Codes:        T {cmbTaxCodes_Dividend}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots}

' ** MapType = "Dividend"

'rptMap_Dividend
'qryMapReport_01

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Combo box column constants: curr_id.
Private Const CBX_C_CURRID As Integer = 0  'curr_id
'Private Const CBX_C_CODE   As Integer = 1  'curr_code
'Private Const CBX_C_NAME   As Integer = 2  'curr_name
Private Const CBX_C_SYM    As Integer = 3  'currsym_symbol
Private Const CBX_C_DEC    As Integer = 4  'curr_decimal
'Private Const CBX_C_RATE1  As Integer = 5  'curr_rate1
Private Const CBX_C_RATE2  As Integer = 6  'curr_rate2
Private Const CBX_C_DATE   As Integer = 7  'curr_date

' ** Array: arr_varNoChar().
Private lngNoChars As Long, arr_varNoChar As Variant
Private Const N_CID   As Integer = 0
'Private Const N_CNAM  As Integer = 1
'Private Const N_FID   As Integer = 2
'Private Const N_FNAM  As Integer = 3
'Private Const N_NOCHR As Integer = 4

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngPrintReportLbl_Left As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private strFrmt1 As String, strFrmt2 As String, intDec1 As Integer, intDec2 As Integer, lngBkClr As Long
Private blnGoingToReport As Boolean, blnGTR_Emblem As Boolean, lngGTR_Stat As Long
Private blnCurrID As Boolean, lngCurrID As Long
Private blnRecalcTots As Boolean
Private dblShareface As Double, dblAmount As Double, dblPerShare As Double
Private dblOverUnder As Double, lngTaxcode As Long, lngAssetNo As Long
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = Trim(.OpenArgs)
150           If CharCnt(strTmp01, "~") = 7& Then  ' ** Module Function: modStringFuncs.

160             blnIsOpen = True

                ' ** strCallingForm & "~" & CStr(dblOverUnder) & "~" & CStr(dblPerShare) & "~" & CStr(lngTaxCode) & "~" & CStr(lngAssetNo)
170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             dblOverUnder = Val(Left(strTmp01, (intPos01 - 1)))
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             dblPerShare = Val(Left(strTmp01, (intPos01 - 1)))
250             strTmp01 = Mid(strTmp01, (intPos01 + 1))
260             intPos01 = InStr(strTmp01, "~")
270             lngTaxcode = Val(Left(strTmp01, (intPos01 - 1)))
280             strTmp01 = Mid(strTmp01, (intPos01 + 1))
290             intPos01 = InStr(strTmp01, "~")
300             lngAssetNo = Val(Left(strTmp01, (intPos01 - 1)))
310             strTmp01 = Mid(strTmp01, (intPos01 + 1))
320             intPos01 = InStr(strTmp01, "~")
330             dblShareface = Val(Left(strTmp01, (intPos01 - 1)))
340             strTmp01 = Mid(strTmp01, (intPos01 + 1))
350             intPos01 = InStr(strTmp01, "~")
360             dblAmount = Val(Mid(strTmp01, (intPos01 - 1)))
370             lngCurrID = Val(Mid(strTmp01, (intPos01 + 1)))
380             lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

390             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
400             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

                ' ** Variables are fed empty, then populated ByRef.
410             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

420             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

430             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
440             If lngMonitorCnt > 1& Then
450               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
460               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
470               If lngMonitorNum > 1& Then
480                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
490               End If
500             End If

510             CLR_DISABLED_FG = CLR_DKGRY
520             CLR_DISABLED_BG = CLR_LTTEAL

530             blnCurrID = Pref_CurrID  ' ** Module Function: modPreferenceFuncs.
540             .curr_id = lngCurrID
550             IncludeCurrency_Sub  ' ** Procedure: Below.

560             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
570             blnGTR_Emblem = False: lngGTR_Stat = 0&

580             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

590             Set dbs = CurrentDb
600             With dbs
610               Select Case strCallingForm
                  Case "frmJournal"
                    ' ** Journal Map, grouped and summed, by accountno.
620                 Set qdf = .QueryDefs("qryMap_Div_02_04")
630               Case "frmJournal_Columns"
                    ' ** Journal Map, grouped and summed, by accountno.
640                 Set qdf = .QueryDefs("qryJournal_Columns_30_05_04")
650               End Select
660               Set rst = qdf.OpenRecordset
670               With rst
680                 If .BOF = True And .EOF = True Then
690                   lngRecs = 0&
700                 Else
710                   .MoveLast
720                   lngRecs = .RecordCount
730                 End If
740                 .Close
750               End With
760               Set rst = Nothing
770               .Close
780             End With
790             Set dbs = Nothing
800             DoEvents

810             Select Case strCallingForm
                Case "frmJournal"
820               If .frmMap_Div_Detail_Sub.Form.RecordSource <> "qryMap_Div_02_04" Then .frmMap_Div_Detail_Sub.Form.RecordSource = "qryMap_Div_02_04"
830             Case "frmJournal_Columns"
840               If .frmMap_Div_Detail_Sub.Form.RecordSource <> "qryJournal_Columns_30_05_04" Then .frmMap_Div_Detail_Sub.Form.RecordSource = "qryJournal_Columns_30_05_04"
850             End Select

860             .maptype = "Dividend"
870             .txtNumberRecords = "Number of Records = " & CStr(lngRecs)
880             .overshort = dblOverUnder
890             .pershare = dblPerShare  '5 dec.

900             If .cmdReinvest.Enabled = True Then
910               .cmdReinvest.SetFocus
920             ElseIf .cmdOK.Enabled = True Then
930               .cmdOK.SetFocus
940             Else
950               .description.SetFocus
960             End If

970             DoEvents

980             blnRecalcTots = True
990             .TimerInterval = 500&

1000            DoEvents
1010            DoCmd.SelectObject acForm, THIS_NAME, False
1020            DoEvents

1030            blnIsOpen = False

1040          Else
1050            Cancel = -1
1060          End If
1070        Else
1080          Cancel = -1
1090        End If
1100      Else
1110        Cancel = -1
1120      End If
1130    End With

1140    If Cancel = -1 Then
1150      DoCmd.Hourglass False
1160      Beep
1170      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1180    End If

EXITP:
1190    Set rst = Nothing
1200    Set qdf = Nothing
1210    Set dbs = Nothing
1220    Exit Sub

ERRH:
1230    DoCmd.Hourglass False
1240    Select Case ERR.Number
        Case Else
1250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1260    End Select
1270    Resume EXITP

End Sub

Private Sub Form_Load()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1310    With Me

1320      .total_shareface_sum.ForeColor = CLR_DISABLED_FG
1330      .total_shareface_sum.BackColor = CLR_DISABLED_BG
1340      .total_cash_sum.ForeColor = CLR_DISABLED_FG
1350      .total_cash_sum.BackColor = CLR_DISABLED_BG
1360      .overshort.ForeColor = CLR_DISABLED_FG
1370      .overshort.BackColor = CLR_DISABLED_BG
1380      .pershare.ForeColor = CLR_DISABLED_FG
1390      .pershare.BackColor = CLR_DISABLED_BG
1400      DoEvents

          ' ** Borrowing these variables from the Court Reports.
1410      If gstrCrtRpt_Version <> vbNullString Then
1420        .cusip = gstrCrtRpt_Ordinal
1430        .cusip.ForeColor = CLR_DISABLED_FG
1440        .cusip.BackColor = CLR_DISABLED_BG
1450        .cusip.Visible = True
1460        .totdesc = gstrCrtRpt_Version
1470        .totdesc.ForeColor = CLR_DISABLED_FG
1480        .totdesc.BackColor = CLR_DISABLED_BG
1490        .totdesc.Visible = True
1500      End If
1510      DoEvents

1520      .transdate = .frmMap_Div_Detail_Sub.Form.transdate
1530      .cmbTaxCodes_Dividend = lngTaxcode
1540      DoEvents

1550      Select Case IsNull(.cmbRevenueCodes_Dividend)
          Case True
1560        .cmbRevenueCodes_Dividend = REVID_INC
1570        .cmbRevenueCodes_Dividend_display = Null
1580      Case False
1590        If .cmbRevenueCodes_Dividend = REVID_INC Or .cmbRevenueCodes_Dividend = REVID_EXP Then
1600          .cmbRevenueCodes_Dividend_display = Null
1610        Else
1620          .cmbRevenueCodes_Dividend_display = .cmbRevenueCodes_Dividend.Column(1)
1630        End If
1640      End Select
1650      DoEvents

1660      Select Case IsNull(.cmbTaxCodes_Dividend)
          Case True
1670        .cmbTaxCodes_Dividend = TAXID_INC
1680        .cmbTaxCodes_Dividend_display = Null
1690      Case False
1700        If .cmbTaxCodes_Dividend = TAXID_INC Or .cmbTaxCodes_Dividend = TAXID_DED Then
1710          .cmbTaxCodes_Dividend_display = Null
1720        Else
1730          .cmbTaxCodes_Dividend_display = .cmbTaxCodes_Dividend.Column(1)
1740        End If
1750      End Select
1760      DoEvents

1770      Select Case gblnRevenueExpenseTracking
          Case True
1780        .cmbRevenueCodes_Dividend.Visible = True
1790        .cmbRevenueCodes_Dividend_display.Visible = True
1800        .cmbRevenueCodes_Dividend_lbl2.Visible = False  ' ** Option off msg.
1810      Case False
1820        .cmbRevenueCodes_Dividend.Visible = False
1830        .cmbRevenueCodes_Dividend_display.Visible = False
1840        .cmbRevenueCodes_Dividend_lbl_box.Visible = False
1850        .cmbRevenueCodes_Dividend_lbl2.Visible = True  ' ** Option off msg.
1860      End Select
1870      DoEvents

1880      Select Case gblnIncomeTaxCoding
          Case True
1890        .cmbTaxCodes_Dividend.Visible = True
1900        .cmbTaxCodes_Dividend_display.Visible = True
1910        .cmbTaxCodes_Dividend_lbl2.Visible = False  ' ** Option off msg.
1920      Case False
1930        .cmbTaxCodes_Dividend.Visible = False
1940        .cmbTaxCodes_Dividend_display.Visible = False
1950        .cmbTaxCodes_Dividend_lbl_box.Visible = False  ' ** Option off msg.
1960        .cmbTaxCodes_Dividend_lbl2.Visible = True
1970      End Select
1980      DoEvents

1990      If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
2000        .cmbRevenueCodes_Dividend_box_lbl.ForeColor = WIN_CLR_DISF
2010        .cmbRevenueCodes_Dividend_box_lbl_dim_hi.Visible = True
2020      End If
2030      DoEvents

2040      curr_id_AfterUpdate  ' ** Procedure: Below.
2050      IncludeCurrency False  ' ** Procedure: Below
2060      DoEvents

2070      DoCmd.Hourglass False
2080      .frmMap_Div_Detail_Sub.Form.TimerInterval = 500&

2090    End With

EXITP:
2100    Exit Sub

ERRH:
2110    DoCmd.Hourglass False
2120    Select Case ERR.Number
        Case Else
2130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2140    End Select
2150    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Comments:         M {description}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc Totals:    F5 {RecalcTots}

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyEscape
2270        intRetVal = 0
2280        cmdCancel_Click  ' ** Procedure: Below.
2290      Case vbKeyF5
2300        intRetVal = 0
2310        RecalcTots  ' ** Function: Below.
2320      End Select
2330    End If

        ' ** Ctrl keys.
2340    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2350      Select Case intRetVal
          Case vbKeyM
2360        With Me
2370          intRetVal = 0
2380          .description.SetFocus
2390        End With
2400      End Select
2410    End If

        ' ** Ctrl-Shift keys.
2420    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2430      Select Case intRetVal
          Case vbKeyF
2440        With Me
2450          intRetVal = 0
2460          .FocusHolder.SetFocus
2470        End With
2480      End Select
2490    End If

EXITP:
2500    KeyCode = intRetVal
2510    Exit Sub

ERRH:
2520    intRetVal = 0
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub Form_Timer()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim lngTmp03 As Long
        Dim blnRetVal As Boolean

        Static lngTmp01 As Long, lngTmp02 As Long

2610    With Me
2620      .TimerInterval = 0&
2630      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
2640        lngGTR_Stat = lngGTR_Stat + 1&
2650        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
2660        lngTmp01 = lngTmp01 + 1&
2670        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
2680        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
2690        If lngTmp01 > 1& Then
2700          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
2710        End If
2720        If lngTmp01 = 24& Then
2730          lngTmp01 = 0&
2740          lngTmp02 = lngTmp02 + 1&
2750          If lngTmp02 < 8& Then
2760            .TimerInterval = 50&
2770          Else
2780            blnGTR_Emblem = False
2790          End If
2800        Else
2810          .TimerInterval = 50&
2820        End If
2830      End If  ' ** blnGTR_Emblem.
2840      If gblnGoToReport = True Then
2850        DoCmd.Hourglass True  ' ** Make sure it's still running.
2860        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or _
                (lngGTR_Stat > 96& And lngGTR_Stat < 120&) Or (lngGTR_Stat > 120&)) Then
              ' ** Just let the emblem go.
2870        Else
2880          If blnGTR_Emblem = False Then
2890            DoCmd.Hourglass True  ' ** Make sure it's still running.
2900            blnGTR_Emblem = True
2910            If IsNull(garr_varGoToReport(GTR_CTL4)) = False Then
2920              lngTmp01 = 0&: lngTmp02 = 0&
2930              lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
2940              .GoToReport_arw_map_okr_img.Left = (.GoToReport_arw_map_okr_img.Left - lngTmp03)
2950              .cmdReinvest.Left = (.cmdReinvest.Left - lngTmp03)
2960              .cmdOK.Left = (.cmdOK.Left - lngTmp03)
2970              .cmdCancel.Left = (.cmdCancel.Left - lngTmp03)
2980            End If
2990          End If
3000          DoEvents
3010          Select Case blnGoingToReport
              Case True
3020            Select Case IsNull(garr_varGoToReport(GTR_CTL4))
                Case True
3030              DoCmd.Hourglass False
3040              gblnGoToReport = False
3050              blnGoingToReport = False
3060              .GoToReport_arw_pr_img.Visible = False
3070              .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
3080              GTREmblem_Off  ' ** Procedure: Below.
3090              .cmdCancel.SetFocus
3100              .TimerInterval = 0&
3110            Case False
3120              .TimerInterval = 50&
3130              cmdReinvest_Click  ' ** Procedure: Below.
3140            End Select
3150          Case False
3160            DoCmd.Hourglass True  ' ** Make sure it's still running.
3170            DoEvents
3180            blnGoingToReport = True
3190            Select Case IsNull(garr_varGoToReport(GTR_CTL4))
                Case True
3200              .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3210              .GoToReport_arw_pr_img.Visible = True
3220              If .cmdPrintReport.Enabled = True Then
3230                .cmdPrintReport.SetFocus
3240              End If
3250              DoEvents
                  ' ** Make sure this remains off!
3260              GTREmblem_Off  ' ** Procedure: Below.
3270              Beep
3280              DoCmd.Hourglass False
3290              DoEvents
3300              .TimerInterval = (GTR_WAIT * 2&)
3310            Case False
3320              lngBtn_Top = (.cmdReinvest.Top + (lngTpp))
3330              lngBtn_Left = ((.cmdReinvest.Left - .GoToReport_arw_map_okr_img.Width) + (2& * lngTpp))
3340              .GoToReport_arw_map_okr_img.Top = lngBtn_Top
3350              .GoToReport_arw_map_okr_img.Left = lngBtn_Left
3360              .GoToReport_arw_map_okr_img.Visible = True
3370              .cmdReinvest.SetFocus
3380              DoEvents
3390              .TimerInterval = 50&
3400            End Select
3410          End Select
3420        End If
3430      End If
3440      If blnRecalcTots = True Then
3450        blnRecalcTots = False
3460        blnRetVal = RecalcTots  ' ** Function: Below.
3470        If blnRetVal = False Then
3480          blnRecalcTots = True
3490          .TimerInterval = 500&
3500        End If
3510      End If
3520    End With

EXITP:
3530    Exit Sub

ERRH:
3540    DoCmd.Hourglass False
3550    Select Case ERR.Number
        Case Else
3560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3570    End Select
3580    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3610    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Assure all reports are closed.
3620    Do While Reports.Count > 0
3630      DoCmd.Close acReport, Reports(0).Name
3640    Loop

3650    gstrReportCallingForm = vbNullString

EXITP:
3660    Exit Sub

ERRH:
3670    Select Case ERR.Number
        Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3810    With Me
3820      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3830        Select Case blnPrintReport_Focus
            Case True
3840          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3850          .cmdPrintReport_raised_img.Visible = False
3860        Case False
3870          .cmdPrintReport_raised_img.Visible = True
3880          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3890        End Select
3900        .cmdPrintReport_raised_focus_img.Visible = False
3910        .cmdPrintReport_raised_focus_dots_img.Visible = False
3920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3930        .cmdPrintReport_raised_img_dis.Visible = False
3940      End If
3950    End With

EXITP:
3960    Exit Sub

ERRH:
3970    Select Case ERR.Number
        Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub cmdOK_Click()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRevCode As Long, lngRevType As Long, lngTaxType As Long
        Dim strRevDesc As String, strTaxDesc As String
        Dim varComments As Variant

4210    DoCmd.Hourglass True
4220    DoEvents

4230    With Me
4240      lngRevCode = .cmbRevenueCodes_Dividend
4250      strRevDesc = .cmbRevenueCodes_Dividend.Column(1)
4260      lngRevType = .cmbRevenueCodes_Dividend.Column(2)
4270      lngTaxcode = .cmbTaxCodes_Dividend
4280      strTaxDesc = .cmbTaxCodes_Dividend.Column(1)
4290      lngTaxType = .cmbTaxCodes_Dividend.Column(2)
4300      varComments = .description
4310    End With

4320    Set dbs = CurrentDb
4330    With dbs

4340      Select Case strCallingForm
          Case "frmJournal"
            ' ** Append Journal Map to Journal, by specified [desc], [revcod], [taxcod],
            ' ** [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id.
4350        Set qdf = .QueryDefs("qryMap_Div_03")
4360      Case "frmJournal_Columns"
            ' ** Append Journal Map to tblJournal_Column, by specified
            ' ** [desc], [revcod], [taxcod], [revdesc], [taxdesc], [revtyp], [taxtyp],
            ' ** [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id.
4370        Set qdf = .QueryDefs("qryJournal_Columns_30_06")
4380        With qdf.Parameters
4390          ![revdesc] = strRevDesc
4400          ![taxdesc] = strTaxDesc
4410          ![revtyp] = lngRevType
4420          ![taxtyp] = lngTaxType
4430        End With
4440      End Select
4450      With qdf.Parameters
4460        ![Desc] = varComments
4470        ![revcod] = lngRevCode
4480        ![taxcod] = lngTaxcode
4490        ![reinvstd] = False
4500      End With
4510      qdf.Execute
4520      Set qdf = Nothing
4530      DoEvents

4540      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
4550        Set qdf = .QueryDefs("qryMap_01")
4560      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
4570        Set qdf = .QueryDefs("qryJournal_Columns_30_04")
4580      End Select
4590      qdf.Execute
4600      Set qdf = Nothing
4610      DoEvents
          ' ** Empty tmpAccount.
4620      Set qdf = .QueryDefs("qryMap_Div_08_01")
4630      qdf.Execute
4640      Set qdf = Nothing
4650      DoEvents

4660      .Close
4670    End With
4680    Set dbs = Nothing
4690    DoEvents

4700    Select Case strCallingForm
        Case "frmJournal"
4710      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4720        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
4730      Else
4740        Forms(strCallingForm).frmJournal_Sub1_Dividend.Form.Requery
4750        DoCmd.SelectObject acForm, strCallingForm, False
4760      End If
4770      DoEvents
4780      Forms(strCallingForm).frmJournal_Sub1_Dividend.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub1_Dividend.
4790    Case "frmJournal_Columns"
4800      gblnSetFocus = True
4810      gstrReturningForm = "frmMap_Return"
4820      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4830        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
4840      Else
4850        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
4860        Forms(strCallingForm).frmJournal_Columns_Sub.Form.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
4870        DoCmd.SelectObject acForm, strCallingForm, False
4880        Forms(strCallingForm).TimerInterval = 250&
4890      End If
4900    End Select

4910    DoCmd.Close acForm, THIS_NAME

EXITP:
4920    DoCmd.Hourglass False
4930    Set qdf = Nothing
4940    Set dbs = Nothing
4950    Exit Sub

ERRH:
4960    DoCmd.Hourglass False
4970    Select Case ERR.Number
        Case Else
4980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4990    End Select
5000    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
5190          .cmdReinvest.SetFocus
5200        End With
5210      End Select
5220    End If

        ' ** Shift keys.
5230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5240      Select Case intRetVal
          Case vbKeyTab
5250        With Me
5260          intRetVal = 0
5270          .cmdPrintReport.SetFocus
5280        End With
5290      End Select
5300    End If

EXITP:
5310    KeyCode = intRetVal
5320    Exit Sub

ERRH:
5330    intRetVal = 0
5340    Select Case ERR.Number
        Case Else
5350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5360    End Select
5370    Resume EXITP

End Sub

Private Sub cmdReinvest_Click()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReinvest_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset
        Dim lngRevCode As Long, lngRevType As Long, lngTaxType As Long
        Dim strRevDesc As String, strTaxDesc As String
        Dim varComments As Variant
        Dim strDocName As String

5410    DoCmd.Hourglass True
5420    DoEvents

5430    With Me
5440      lngRevCode = .cmbRevenueCodes_Dividend
5450      strRevDesc = .cmbRevenueCodes_Dividend.Column(1)
5460      lngRevType = .cmbRevenueCodes_Dividend.Column(2)
5470      lngTaxcode = .cmbTaxCodes_Dividend
5480      strTaxDesc = .cmbTaxCodes_Dividend.Column(1)
5490      lngTaxType = .cmbTaxCodes_Dividend.Column(2)
5500      varComments = .description
5510    End With

5520    Set dbs = CurrentDb
5530    With dbs

5540      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty tblMark.
5550        Set qdf = .QueryDefs("qryMap_Div_03m")
5560        qdf.Execute
5570        Set qdf = Nothing
5580        DoEvents
            ' ** Append qryMap_Div_03n (Journal, grouped, with Max(ID)) to tblMark.
5590        Set qdf = .QueryDefs("qryMap_Div_03o")
5600        qdf.Execute
5610        Set qdf = Nothing
5620        DoEvents
            ' ** Append Journal Map to Journal, by specified [desc], [revcod], [taxcod],
            ' ** [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id
5630        Set qdf = .QueryDefs("qryMap_Div_03")
5640      Case "frmJournal_Columns"
            ' ** Empty tblMark.
5650        Set qdf = .QueryDefs("qryJournal_Columns_30_06m")
5660        qdf.Execute
5670        Set qdf = Nothing
5680        DoEvents
            ' ** Append qryJournal_Columns_30_06b (tblJournal_Column, grouped, with Max(JrnlCol_ID))to tblMark.
5690        Set qdf = .QueryDefs("qryJournal_Columns_30_06o")
5700        qdf.Execute
5710        Set qdf = Nothing
5720        DoEvents
            ' ** Append Journal Map to tblJournal_Column, by specified [desc], [revcod], [taxcod],
            ' ** [revdesc], [taxdesc], [revtyp], [taxtyp], [reinvstd], GlobalVarGet('gblnGoToReport').
5730        Set qdf = .QueryDefs("qryJournal_Columns_30_06")
5740        With qdf.Parameters
5750          ![revdesc] = strRevDesc
5760          ![taxdesc] = strTaxDesc
5770          ![revtyp] = lngRevType
5780          ![taxtyp] = lngTaxType
5790        End With
5800      End Select
5810      With qdf.Parameters
5820        ![Desc] = varComments
5830        ![revcod] = lngRevCode
5840        ![taxcod] = lngTaxcode
5850        ![reinvstd] = True  ' ** It will only update the appropriate side.
5860      End With
5870      qdf.Execute    ' ** Journal Map ID is in CheckNum.
5880      Set qdf = Nothing
5890      DoEvents

          ' ** Empty tblJournal_Map_Staging3.
5900      Set qdf = .QueryDefs("qryMap_Reinvest_05_07")
5910      qdf.Execute
5920      Set qdf = Nothing
5930      DoEvents

          ' ** Append Journal Map to tblJournal_Map_Staging3.
5940      Set qdf = .QueryDefs("qryMap_Reinvest_05_04")
5950      qdf.Execute
5960      Set qdf = Nothing
5970      DoEvents

5980      Select Case strCallingForm
          Case "frmJournal"
            ' ** Update qryMap_Reinvest_05_05 (tblJournal_Map_Staging3,
            ' ** linked to Journal, with Journal_ID_new).
5990        Set qdf = .QueryDefs("qryMap_Reinvest_05_06")
6000      Case "frmJournal_Columns"
            ' ** Update qryJournal_Columns_30_14 (tblJournal_Map_Staging3,
            ' ** linked to tblJournal_Column, with JrnlCol_ID_new).
6010        Set qdf = .QueryDefs("qryJournal_Columns_30_15")
6020      End Select
6030      qdf.Execute
6040      Set qdf = Nothing
6050      DoEvents

6060      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
6070        Set qdf = .QueryDefs("qryMap_01")
6080      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
6090        Set qdf = .QueryDefs("qryJournal_Columns_30_04")
6100      End Select
6110      qdf.Execute
6120      Set qdf = Nothing

6130      .Close
6140    End With
6150    Set dbs = Nothing
6160    DoEvents

6170    If strCallingForm = "frmJournal_Columns" Then
6180      Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
6190    End If

6200    strDocName = "frmMap_Reinvest_DivInt_Price"
6210    DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & "Dividend" & "~" & CStr(lngAssetNo) & "~" & _
          CStr(dblPerShare) & "~0" & "~" & CStr(lngCurrID)  ' ** 2nd to last parameter not needed for Dividend.
6220    If gblnGoToReport = True Then
6230      Forms(strDocName).TimerInterval = 100&
6240    End If

6250    DoCmd.Close acForm, THIS_NAME

EXITP:
6260    DoCmd.Hourglass False
6270    Set rst1 = Nothing
6280    Set rst2 = Nothing
6290    Set rst3 = Nothing
6300    Set qdf = Nothing
6310    Set dbs = Nothing
6320    Exit Sub

ERRH:
6330    DoCmd.Hourglass False
6340    Select Case ERR.Number
        Case Else
6350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6360    End Select
6370    Resume EXITP

End Sub

Private Sub cmdReinvest_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReinvest_KeyDown"

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
6570          .cmdOK.SetFocus
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

Public Sub cmdCancel_Click()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intTmp01 As Integer
        Dim intX As Integer

6710    DoCmd.Hourglass True
6720    DoEvents

6730    intTmp01 = 0
6740    Do While Forms.Count > 1
6750      intTmp01 = intTmp01 + 1
6760      For intX = (Forms.Count - 1) To 0 Step -1
6770        Select Case Forms(intX).Name
            Case THIS_NAME
              ' ** Skip.
6780        Case "frmJournal", "frmJournal_Columns"
              ' ** Skip.
6790        Case Else
6800          DoCmd.Close acForm, Forms(intX).Name
6810          Exit For  ' ** Do them one-at-a-time.
6820        End Select
6830      Next
6840      If intTmp01 >= 10 Then Exit Do  ' ** Don't get caught in a loop.
6850    Loop

6860    Set dbs = CurrentDb
6870    With dbs
6880      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
6890        Set qdf = .QueryDefs("qryMap_01")
6900      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
6910        Set qdf = .QueryDefs("qryJournal_Columns_30_04")
6920      End Select
6930      qdf.Execute
6940      Set qdf = Nothing
6950      DoEvents
          ' ** Empty tmpAccount.
6960      Set qdf = .QueryDefs("qryMap_Div_08_01")
6970      qdf.Execute
6980      Set qdf = Nothing
6990      DoEvents
7000      .Close
7010    End With
7020    Set dbs = Nothing
7030    DoEvents

7040    gblnSetFocus = True
7050    Select Case strCallingForm
        Case "frmJournal"
7060      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
7070        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
7080      Else
7090        gstrReturningForm = THIS_NAME
7100        Forms(strCallingForm).frmJournal_Sub1_Dividend.Form.Requery
7110        DoCmd.SelectObject acForm, strCallingForm, False
7120        Forms(strCallingForm).TimerInterval = 250&
7130      End If
7140    Case "frmJournal_Columns"
7150      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
7160        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
7170      Else
7180        gstrReturningForm = THIS_NAME
7190        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
7200        DoCmd.SelectObject acForm, strCallingForm, False
7210        Forms(strCallingForm).TimerInterval = 250&
7220      End If
7230    End Select

7240    DoCmd.Close acForm, THIS_NAME

EXITP:
7250    Set qdf = Nothing
7260    Set dbs = Nothing
7270    Exit Sub

ERRH:
7280    DoCmd.Hourglass False
7290    Select Case ERR.Number
        Case Else
7300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7310    End Select
7320    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7420    intShiftDown = (Shift And acShiftMask) > 0
7430    intAltDown = (Shift And acAltMask) > 0
7440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7460      Select Case intRetVal
          Case vbKeyTab
7470        With Me
7480          intRetVal = 0
7490          lngRecsCur = .frmMap_Div_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Div_Detail_Sub.
7500          If lngRecsCur > 0 Then
7510            .frmMap_Div_Detail_Sub.SetFocus
7520            .frmMap_Div_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMap_Div_Detail_Sub.
7530            .frmMap_Div_Detail_Sub.Form.FocusHolder.SetFocus
7540          Else
7550            .description.SetFocus
7560          End If
7570        End With
7580      End Select
7590    End If

        ' ** Shift keys.
7600    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7610      Select Case intRetVal
          Case vbKeyTab
7620        With Me
7630          intRetVal = 0
7640          .cmdReinvest.SetFocus
7650        End With
7660      End Select
7670    End If

EXITP:
7680    KeyCode = intRetVal
7690    Exit Sub

ERRH:
7700    intRetVal = 0
7710    Select Case ERR.Number
        Case Else
7720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7730    End Select
7740    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

7810    With Me
7820      blnPrintReport_Focus = True
7830      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
7840      .cmdPrintReport_raised_img.Visible = False
7850      .cmdPrintReport_raised_focus_img.Visible = False
7860      .cmdPrintReport_raised_focus_dots_img.Visible = False
7870      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7880      .cmdPrintReport_raised_img_dis.Visible = False
7890    End With

EXITP:
7900    Exit Sub

ERRH:
7910    Select Case ERR.Number
        Case Else
7920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7930    End Select
7940    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

8010    With Me
8020      blnPrintReport_MouseDown = True
8030      .cmdPrintReport_sunken_focus_dots_img.Visible = True
8040      .cmdPrintReport_raised_img.Visible = False
8050      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8060      .cmdPrintReport_raised_focus_img.Visible = False
8070      .cmdPrintReport_raised_focus_dots_img.Visible = False
8080      .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_Click()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

        ' ** qryMapReport_01.
        'strSQL = "SELECT DISTINCTROW account.accountno, account.shortname, Sum([journal map].shareface) AS total_shareface, " & _
        '  "Sum([journal map].icash) AS total_icash, " & CoInfo & ", " & _
        '  "CStr([masterasset].[Description]) & " & _
        '  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%'),'') & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'),'') AS totdesc " & _
        '  "FROM ([journal map] INNER JOIN account ON [journal map].accountno = account.accountno) " & _
        '  "INNER JOIN masterasset ON [journal map].assetno = masterasset.assetno " & _
        '  "GROUP BY account.accountno, account.shortname, CStr([masterasset].[Description]) & " & _
        '  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%'),'') & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'),'');"

8210    With Me
8220      strDocName = "rptMap_Dividend"
8230      gstrReportCallingForm = THIS_NAME
8240      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
8250        .Modal = False
8260        DoCmd.OpenReport strDocName, acViewPreview, , , , acViewPreview
8270        DoCmd.Maximize
8280        DoCmd.RunCommand acCmdFitToWindow
8290        .Visible = False
8300      Else
            '##GTR_Ref: rptMap_Dividend
8310        DoCmd.OpenReport strDocName, acViewNormal, , , , acViewNormal
8320      End If
8330    End With

EXITP:
8340    Exit Sub

ERRH:
8350    Me.Visible = True
8360    DoCmd.Restore
8370    Select Case ERR.Number
        Case Else
8380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8390    End Select
8400    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8520    intShiftDown = (Shift And acShiftMask) > 0
8530    intAltDown = (Shift And acAltMask) > 0
8540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8560      Select Case intRetVal
          Case vbKeyTab
8570        With Me
8580          intRetVal = 0
8590          .cmdOK.SetFocus
8600        End With
8610      End Select
8620    End If

        ' ** Shift keys.
8630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8640      Select Case intRetVal
          Case vbKeyTab
8650        With Me
8660          intRetVal = 0
8670          If .cmbTaxCodes_Dividend.Visible = True And .cmbTaxCodes_Dividend.Enabled = True Then
8680            .cmbTaxCodes_Dividend.SetFocus
8690          ElseIf .cmbRevenueCodes_Dividend.Visible = True And .cmbRevenueCodes_Dividend.Enabled = True Then
8700            .cmbRevenueCodes_Dividend.SetFocus
8710          Else
8720            .description.SetFocus
8730          End If
8740        End With
8750      End Select
8760    End If

EXITP:
8770    KeyCode = intRetVal
8780    Exit Sub

ERRH:
8790    intRetVal = 0
8800    Select Case ERR.Number
        Case Else
8810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8820    End Select
8830    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

8910    With Me
8920      If blnPrintReport_MouseDown = False Then
8930        Select Case blnPrintReport_Focus
            Case True
8940          .cmdPrintReport_raised_focus_dots_img.Visible = True
8950          .cmdPrintReport_raised_focus_img.Visible = False
8960        Case False
8970          .cmdPrintReport_raised_focus_img.Visible = True
8980          .cmdPrintReport_raised_focus_dots_img.Visible = False
8990        End Select
9000        .cmdPrintReport_raised_img.Visible = False
9010        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9020        .cmdPrintReport_sunken_focus_dots_img.Visible = False
9030        .cmdPrintReport_raised_img_dis.Visible = False
9040      End If
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

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

9210    With Me
9220      .cmdPrintReport_raised_focus_dots_img.Visible = True
9230      .cmdPrintReport_raised_img.Visible = False
9240      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9250      .cmdPrintReport_raised_focus_img.Visible = False
9260      .cmdPrintReport_sunken_focus_dots_img.Visible = False
9270      .cmdPrintReport_raised_img_dis.Visible = False
9280      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

9410    With Me
9420      .cmdPrintReport_raised_img.Visible = True
9430      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9440      .cmdPrintReport_raised_focus_img.Visible = False
9450      .cmdPrintReport_raised_focus_dots_img.Visible = False
9460      .cmdPrintReport_sunken_focus_dots_img.Visible = False
9470      .cmdPrintReport_raised_img_dis.Visible = False
9480      blnPrintReport_Focus = False
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

Private Sub total_shareface_sum_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_sum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9670        With Me
9680          intRetVal = 0
9690          .description.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9750        With Me
9760          intRetVal = 0
9770          lngRecsCur = .frmMap_Div_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Div_Detail_Sub.
9780          If lngRecsCur > 0& Then
9790            .frmMap_Div_Detail_Sub.SetFocus
9800            .frmMap_Div_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Div_Detail_Sub.
9810            .frmMap_Div_Detail_Sub.Form.FocusHolder.SetFocus
9820          Else
9830            .cmdCancel.SetFocus
9840          End If
9850        End With
9860      End Select
9870    End If

EXITP:
9880    KeyCode = intRetVal
9890    Exit Sub

ERRH:
9900    intRetVal = 0
9910    Select Case ERR.Number
        Case Else
9920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9930    End Select
9940    Resume EXITP

End Sub

Private Sub total_cash_sum_KeyDown(KeyCode As Integer, Shift As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "total_cash_sum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10020   intShiftDown = (Shift And acShiftMask) > 0
10030   intAltDown = (Shift And acAltMask) > 0
10040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10070       With Me
10080         intRetVal = 0
10090         .description.SetFocus
10100       End With
10110     End Select
10120   End If

        ' ** Shift keys.
10130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10150       With Me
10160         intRetVal = 0
10170         lngRecsCur = .frmMap_Div_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Div_Detail_Sub.
10180         If lngRecsCur > 0& Then
10190           .frmMap_Div_Detail_Sub.SetFocus
10200           .frmMap_Div_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Div_Detail_Sub.
10210           .frmMap_Div_Detail_Sub.Form.FocusHolder.SetFocus
10220         Else
10230           .cmdCancel.SetFocus
10240         End If
10250       End With
10260     End Select
10270   End If

EXITP:
10280   KeyCode = intRetVal
10290   Exit Sub

ERRH:
10300   intRetVal = 0
10310   Select Case ERR.Number
        Case Else
10320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10330   End Select
10340   Resume EXITP

End Sub

Private Sub Overshort_KeyDown(KeyCode As Integer, Shift As Integer)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "Overshort_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10420   intShiftDown = (Shift And acShiftMask) > 0
10430   intAltDown = (Shift And acAltMask) > 0
10440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10470       With Me
10480         intRetVal = 0
10490         .description.SetFocus
10500       End With
10510     End Select
10520   End If

        ' ** Shift keys.
10530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10550       With Me
10560         intRetVal = 0
10570         lngRecsCur = .frmMap_Div_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Div_Detail_Sub.
10580         If lngRecsCur > 0& Then
10590           .frmMap_Div_Detail_Sub.SetFocus
10600           .frmMap_Div_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Div_Detail_Sub.
10610           .frmMap_Div_Detail_Sub.Form.FocusHolder.SetFocus
10620         Else
10630           .cmdCancel.SetFocus
10640         End If
10650       End With
10660     End Select
10670   End If

EXITP:
10680   KeyCode = intRetVal
10690   Exit Sub

ERRH:
10700   intRetVal = 0
10710   Select Case ERR.Number
        Case Else
10720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10730   End Select
10740   Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

10810   With Me
10820     Select Case .curr_date.Visible
          Case True
10830       .curr_date.Visible = False
10840     Case False
10850       .curr_date = .curr_id.Column(CBX_C_DATE)
10860       .curr_date.Visible = True
10870     End Select
10880   End With

EXITP:
10890   Exit Sub

ERRH:
10900   Select Case ERR.Number
        Case Else
10910     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
10920   End Select
10930   Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String
        Dim lngX As Long

11010   With Me

11020     If IsNull(.curr_id) = False Then

11030       lngCurrID = .curr_id.Column(CBX_C_CURRID)

11040       blnFound = False
11050       If lngCurrID <> 150& Then
11060         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
11070           NoChar_Load  ' ** Procedure: Below.
11080         End If
11090         For lngX = 0& To (lngNoChars - 1&)
11100           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
11110             blnFound = True
11120           End If
11130         Next
11140       End If

11150       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
11160         .total_cash_sum.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
11170         .total_cash_sum.FontSize = 10
11180         .total_cash_sum.FontBold = True
11190         .overshort.FontName = "Courier New"
11200         .overshort.FontSize = 10
11210         .overshort.FontBold = True
11220         .pershare.FontName = "Courier New"
11230         .pershare.FontSize = 10
11240         .pershare.FontBold = True
11250       Case False
11260         If .total_cash_sum.FontName <> "Arial" Then
11270           .total_cash_sum.FontName = "Arial"
11280           .total_cash_sum.FontSize = 10
11290           .total_cash_sum.FontBold = False
11300           .overshort.FontName = "Arial"
11310           .overshort.FontSize = 10
11320           .overshort.FontBold = False
11330           .pershare.FontName = "Arial"
11340           .pershare.FontSize = 10
11350           .pershare.FontBold = False
11360         End If
11370       End Select

11380       If lngCurrID = 150& Then  ' ** USD.
11390         strFrmt1 = "Currency"
11400         intDec1 = 2
11410         strFrmt2 = "Currency"
11420         intDec2 = 5
11430         lngBkClr = CLR_DISABLED_BG
11440       Else
11450         strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
11460         If .curr_id.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
11470           intDec1 = 2
11480         Else
11490           intDec1 = .curr_id.Column(CBX_C_DEC)  ' ** Decimal places.
11500         End If
11510         intDec2 = 5
11520         intLen = Len(strTmp01)
11530         Select Case intLen
              Case 1
                'strTmp01 = strTmp01
11540         Case 2
11550           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
11560         Case 3
11570           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
11580         Case 4
11590           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
11600         Case 5
11610           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                  "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
11620         End Select
11630         Select Case intDec1
              Case 0
11640           strFrmt1 = strTmp01 & "#,##0;(" & strTmp01 & "#,##0);" & strTmp01 & "0;" & strTmp01 & "0"
11650           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11660         Case 1  ' ** None have this currently.
11670           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
11680           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11690         Case 2
11700           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
11710           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11720         Case 3
11730           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
11740           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11750         End Select
11760         lngBkClr = CLR_LTGRN
11770       End If
11780       .total_cash_sum.Format = strFrmt1
11790       .total_cash_sum.DecimalPlaces = intDec1
11800       .total_cash_sum.BackColor = lngBkClr
11810       .overshort.Format = strFrmt1
11820       .overshort.DecimalPlaces = intDec1
11830       .overshort.BackColor = lngBkClr
11840       .pershare.Format = strFrmt2
11850       .pershare.DecimalPlaces = intDec2
11860       .pershare.BackColor = lngBkClr

11870       If lngCurrID = 150& Then
11880         .total_cash_sum_usd.Visible = False
11890         .Overshort_usd.Visible = False
11900         .pershare_usd.Visible = False
11910       Else
11920         .total_cash_sum_usd = 0
11930         .total_cash_sum_usd.Visible = True
11940         .Overshort_usd = 0
11950         .Overshort_usd.Visible = True
11960         .pershare_usd = 0
11970         .pershare_usd.Visible = True
11980       End If

11990     End If

12000   End With

EXITP:
12010   Exit Sub

ERRH:
12020   Select Case ERR.Number
        Case Else
12030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12040   End Select
12050   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

12110   With Me
12120     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
12130     .curr_id.SetFocus
12140     Response = acDataErrContinue
12150   End With

EXITP:
12160   Exit Sub

ERRH:
12170   Select Case ERR.Number
        Case Else
12180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12190   End Select
12200   Resume EXITP

End Sub

Private Sub description_GotFocus()

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

12310   With Me
12320     .description.SelLength = 0
12330     .description.SelStart = 999
12340   End With

EXITP:
12350   Exit Sub

ERRH:
12360   Select Case ERR.Number
        Case Else
12370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12380   End Select
12390   Resume EXITP

End Sub

Private Sub description_KeyDown(KeyCode As Integer, Shift As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "description_KeyDown"

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
12490         If .cmbRevenueCodes_Dividend.Visible = True And .cmbRevenueCodes_Dividend.Enabled = True Then
12500           .cmbRevenueCodes_Dividend.SetFocus
12510         ElseIf .cmbTaxCodes_Dividend.Visible = True And .cmbTaxCodes_Dividend.Enabled = True Then
12520           .cmbTaxCodes_Dividend.SetFocus
12530         Else
12540           .cmdReinvest.SetFocus
12550         End If
12560       End With
12570     End Select
12580   End If

        ' ** Shift keys.
12590   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12600     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12610       With Me
12620         intRetVal = 0
12630         lngRecsCur = .frmMap_Div_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Div_Detail_Sub.
12640         If lngRecsCur > 0& Then
12650           .frmMap_Div_Detail_Sub.SetFocus
12660           .frmMap_Div_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Div_Detail_Sub.
12670           .frmMap_Div_Detail_Sub.Form.FocusHolder.SetFocus
12680         Else
12690           .cmdCancel.SetFocus
12700         End If
12710       End With
12720     End Select
12730   End If

EXITP:
12740   KeyCode = intRetVal
12750   Exit Sub

ERRH:
12760   intRetVal = 0
12770   Select Case ERR.Number
        Case Else
12780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12790   End Select
12800   Resume EXITP

End Sub

Private Sub description_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "description_MouseMove"

12910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12920   Exit Sub

ERRH:
12930   Select Case ERR.Number
        Case Else
12940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12950   End Select
12960   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Dividend_display_GotFocus()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Dividend_display_GotFocus"

13010   With Me
13020     .cmbRevenueCodes_Dividend.SetFocus
13030   End With

EXITP:
13040   Exit Sub

ERRH:
13050   Select Case ERR.Number
        Case Else
13060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13070   End Select
13080   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Dividend_display_KeyDown(KeyCode As Integer, Shift As Integer)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Dividend_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13120   intShiftDown = (Shift And acShiftMask) > 0
13130   intAltDown = (Shift And acAltMask) > 0
13140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13170       With Me
13180         intRetVal = 0
13190         .cmbRevenueCodes_Dividend.SetFocus
13200       End With
13210     End Select
13220   End If

        ' ** Shift keys.
13230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13250       With Me
13260         intRetVal = 0
13270         .description.SetFocus
13280       End With
13290     End Select
13300   End If

EXITP:
13310   KeyCode = intRetVal
13320   Exit Sub

ERRH:
13330   intRetVal = 0
13340   Select Case ERR.Number
        Case Else
13350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13360   End Select
13370   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Dividend_Enter()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Dividend_Enter"

13410   With Me
13420     Select Case IsNull(.cmbRevenueCodes_Dividend)
          Case True
13430       ExpandCombo .cmbRevenueCodes_Dividend  ' ** Module Procedure: modUtilities.
13440     Case False
13450       If .cmbRevenueCodes_Dividend = REVID_INC Or .cmbRevenueCodes_Dividend = REVID_EXP Then
13460         ExpandCombo .cmbRevenueCodes_Dividend  ' ** Module Procedure: modUtilities.
13470       End If
13480     End Select
13490   End With

EXITP:
13500   Exit Sub

ERRH:
13510   Select Case ERR.Number
        Case Else
13520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13530   End Select
13540   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Dividend_AfterUpdate()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Dividend_AfterUpdate"

13610   With Me
13620     Select Case IsNull(.cmbRevenueCodes_Dividend)
          Case True
13630       .cmbRevenueCodes_Dividend = REVID_INC
13640     Case False
            ' ** Nothing at the moment.
13650     End Select
13660     If .cmbRevenueCodes_Dividend = REVID_INC Or .cmbRevenueCodes_Dividend = REVID_EXP Then
13670       .cmbRevenueCodes_Dividend_display = Null
13680     Else
13690       .cmbRevenueCodes_Dividend_display = .cmbRevenueCodes_Dividend.Column(1)
13700     End If
13710   End With

EXITP:
13720   Exit Sub

ERRH:
13730   Select Case ERR.Number
        Case Else
13740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13750   End Select
13760   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Dividend_KeyDown(KeyCode As Integer, Shift As Integer)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Dividend_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13820   intShiftDown = (Shift And acShiftMask) > 0
13830   intAltDown = (Shift And acAltMask) > 0
13840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13870       With Me
13880         intRetVal = 0
13890         If .cmbTaxCodes_Dividend.Visible = True And .cmbTaxCodes_Dividend.Enabled = True Then
13900           .cmbTaxCodes_Dividend.SetFocus
13910         Else
13920           .cmdReinvest.SetFocus
13930         End If
13940       End With
13950     End Select
13960   End If

        ' ** Shift keys.
13970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13990       With Me
14000         intRetVal = 0
14010         .description.SetFocus
14020       End With
14030     End Select
14040   End If

EXITP:
14050   KeyCode = intRetVal
14060   Exit Sub

ERRH:
14070   intRetVal = 0
14080   Select Case ERR.Number
        Case Else
14090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14100   End Select
14110   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Dividend_display_GotFocus()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Dividend_display_GotFocus"

14210   With Me
14220     .cmbTaxCodes_Dividend.SetFocus
14230   End With

EXITP:
14240   Exit Sub

ERRH:
14250   Select Case ERR.Number
        Case Else
14260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14270   End Select
14280   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Dividend_display_KeyDown(KeyCode As Integer, Shift As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Dividend_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14320   intShiftDown = (Shift And acShiftMask) > 0
14330   intAltDown = (Shift And acAltMask) > 0
14340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14370       With Me
14380         intRetVal = 0
14390         .cmbTaxCodes_Dividend.SetFocus
14400       End With
14410     End Select
14420   End If

        ' ** Shift keys.
14430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14450       With Me
14460         intRetVal = 0
14470         If .cmbRevenueCodes_Dividend.Visible = True And .cmbRevenueCodes_Dividend.Enabled = True Then
14480           .cmbRevenueCodes_Dividend.SetFocus
14490         Else
14500           .description.SetFocus
14510         End If
14520       End With
14530     End Select
14540   End If

EXITP:
14550   KeyCode = intRetVal
14560   Exit Sub

ERRH:
14570   intRetVal = 0
14580   Select Case ERR.Number
        Case Else
14590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14600   End Select
14610   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Dividend_Enter()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Dividend_Enter"

14710   With Me
14720     Select Case IsNull(.cmbTaxCodes_Dividend)
          Case True
14730       ExpandCombo .cmbTaxCodes_Dividend  ' ** Module Procedure: modUtilities.
14740     Case False
14750       If .cmbTaxCodes_Dividend = TAXID_INC Or .cmbTaxCodes_Dividend = TAXID_DED Then
14760         ExpandCombo .cmbTaxCodes_Dividend  ' ** Module Procedure: modUtilities.
14770       End If
14780     End Select
14790   End With

EXITP:
14800   Exit Sub

ERRH:
14810   Select Case ERR.Number
        Case Else
14820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14830   End Select
14840   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Dividend_AfterUpdate()

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Dividend_AfterUpdate"

14910   With Me
14920     Select Case IsNull(.cmbTaxCodes_Dividend)
          Case True
14930       .cmbTaxCodes_Dividend = TAXID_INC
14940     Case False
            ' ** Nothing at the moment.
14950     End Select
14960     If .cmbTaxCodes_Dividend = TAXID_INC Or .cmbTaxCodes_Dividend = TAXID_DED Then
14970       .cmbTaxCodes_Dividend_display = Null
14980     Else
14990       .cmbTaxCodes_Dividend_display = .cmbTaxCodes_Dividend.Column(1)
15000     End If
15010   End With

EXITP:
15020   Exit Sub

ERRH:
15030   Select Case ERR.Number
        Case Else
15040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15050   End Select
15060   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Dividend_KeyDown(KeyCode As Integer, Shift As Integer)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Dividend_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15120   intShiftDown = (Shift And acShiftMask) > 0
15130   intAltDown = (Shift And acAltMask) > 0
15140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15170       With Me
15180         intRetVal = 0
15190         .cmdReinvest.SetFocus
15200       End With
15210     End Select
15220   End If

        ' ** Shift keys.
15230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15250       With Me
15260         intRetVal = 0
15270         If .cmbRevenueCodes_Dividend.Visible = True And .cmbRevenueCodes_Dividend.Enabled = True Then
15280           .cmbRevenueCodes_Dividend.SetFocus
15290         Else
15300           .description.SetFocus
15310         End If
15320       End With
15330     End Select
15340   End If

EXITP:
15350   KeyCode = intRetVal
15360   Exit Sub

ERRH:
15370   intRetVal = 0
15380   Select Case ERR.Number
        Case Else
15390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15400   End Select
15410   Resume EXITP

End Sub

Public Function RecalcTots() As Boolean
' ** I give up trying to get it from the subform!

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim dblTmp01 As Double, dblTmp02 As Double
        Dim lngX As Long
        Dim blnRetVal As Boolean

15510   With Me
15520     blnRetVal = False
15530     Set rst = .frmMap_Div_Detail_Sub.Form.RecordsetClone
15540     With rst
15550       If .BOF = True And .EOF = True Then
              ' ** Ho-Hum...
15560       Else
15570         .MoveLast
15580         lngRecs = .RecordCount
15590         .MoveFirst
15600         dblTmp01 = 0#: dblTmp02 = 0#
15610         For lngX = 1& To lngRecs
15620           dblTmp01 = dblTmp01 + ![total_shareface]
15630           dblTmp02 = dblTmp02 + ![total_icash]
15640           If lngX < lngRecs Then .MoveNext
15650         Next
15660       End If
15670       .Close
15680     End With
15690     If dblTmp01 <> 0# Then
15700       blnRetVal = True
15710     End If
15720     .total_shareface_sum = dblTmp01
15730     .total_cash_sum = dblTmp02
15740   End With

EXITP:
15750   Set rst = Nothing
15760   RecalcTots = blnRetVal
15770   Exit Function

ERRH:
15780   blnRetVal = False
15790   Select Case ERR.Number
        Case Else
15800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15810   End Select
15820   Resume EXITP

End Function

Public Sub IncludeCurrency(blnCalc As Boolean)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim dblRate2 As Double
        Dim blnFound As Boolean
        Dim dblTmp01 As Double

15910   With Me
15920     Select Case blnCalc
          Case True
15930       If blnCurrID = True Then
15940         blnFound = False
15950         Do While blnFound = False
15960           If IsNull(.total_cash_sum) = False Then
15970             If .total_cash_sum <> 0 Then
15980               blnFound = True
15990               dblRate2 = .curr_id.Column(CBX_C_RATE2)
16000               dblTmp01 = (.total_cash_sum * dblRate2)
16010               dblTmp01 = Round(dblTmp01, 2)
16020               .total_cash_sum_usd = dblTmp01
16030               dblTmp01 = (.overshort * dblRate2)
16040               dblTmp01 = Round(dblTmp01, 2)
16050               .Overshort_usd = dblTmp01
16060               dblTmp01 = (.pershare * dblRate2)
16070               dblTmp01 = Round(dblTmp01, 5)
16080               .pershare_usd = dblTmp01
16090               If lngCurrID <> 150& Then
16100                 .total_cash_sum_usd.Visible = True
16110                 .Overshort_usd.Visible = True
16120                 .pershare_usd.Visible = True
16130               End If
16140             End If
16150           End If
16160         Loop
16170       End If
16180     Case False
16190       .curr_id.Enabled = False
16200       .curr_id.Locked = True
16210       .curr_id.ForeColor = CLR_DISABLED_FG
16220       .curr_id.BackColor = CLR_VLTGRN
16230       If blnCurrID = True And lngCurrID <> 150& Then
16240         .curr_id.Visible = True
16250         .curr_id_cmd.Visible = True
16260         .curr_id_cmd.Enabled = True
16270       Else
16280         .curr_id.Visible = False
16290         .curr_id_cmd.Visible = False
16300         .curr_id_cmd.Enabled = False
16310         .cmbRevenueCodes_Dividend_vline05.Visible = False
16320         .cmbRevenueCodes_Dividend_vline06.Visible = False
16330         .cmbRevenueCodes_Dividend_box.Width = (.cmbRevenueCodes_Dividend_vline06.Left - .cmbRevenueCodes_Dividend_box.Left)
16340         .overshort.Top = .Overshort_alt_line.Top
16350         .Overshort_lbl.Top = (.Overshort_alt_line.Top + (2& * lngTpp))
16360         .pershare.Top = .pershare_alt_line.Top
16370         .pershare_lbl.Top = (.pershare_alt_line.Top + (2& * lngTpp))
16380       End If
16390       .total_cash_sum_usd.Visible = False
16400       .Overshort_usd.Visible = False
16410       .pershare_usd.Visible = False
16420     End Select
16430   End With

EXITP:
16440   Exit Sub

ERRH:
16450   DoCmd.Hourglass False
16460   Select Case ERR.Number
        Case Else
16470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16480   End Select
16490   Resume EXITP

End Sub

Private Sub IncludeCurrency_Sub()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long

16510   With Me
16520     If blnCurrID = True And lngCurrID <> 150& Then

16530       With .frmMap_Div_Detail_Sub.Form
16540         lngTmp01 = .Width - .total_icash_alt_line.Left
16550       End With

16560       lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
16570       lngMonitorNum = 1&: lngTmp02 = 0&
16580       EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
16590       If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

16600       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
16610       DoCmd.MoveSize (lngFrm_Left - (lngTmp01 / 2)), lngTmp02, (lngFrm_Width + lngTmp01), lngFrm_Height  'lngFrm_Top
16620       If lngMonitorNum > 1& Then
16630         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
16640       End If

16650       .Width = (.Width + lngTmp01)
16660       .frmMap_Div_Detail_Sub.Width = (.frmMap_Div_Detail_Sub.Width + lngTmp01)
16670       .frmMap_Div_Detail_Sub_box.Width = (.frmMap_Div_Detail_Sub_box.Width + lngTmp01)

16680       .Header_vline01.Left = .Width
16690       .Header_vline02.Left = .Width
16700       .Footer_vline01.Left = .Width
16710       .Footer_vline02.Left = .Width
16720       .Header_hline01.Width = .Width
16730       .Header_hline02.Width = .Width
16740       .Footer_hline01.Width = .Width
16750       .Footer_hline02.Width = .Width

16760       .total_cash_sum_usd.Left = ((.total_cash_sum.Left + .total_cash_sum.Width) + (4& * lngTpp))
16770       .total_cash_sum_usd.Top = .total_cash_sum.Top
16780       .Overshort_usd.Left = ((.overshort.Left + .overshort.Width) + (4& * lngTpp))
16790       .overshort.Top = .Overshort_alt_line.Top
16800       .Overshort_lbl.Top = (.Overshort_alt_line.Top + lngTpp)
16810       .Overshort_usd.Top = .Overshort_alt_line.Top
16820       .pershare_usd.Left = ((.total_cash_sum_usd.Left + .total_cash_sum_usd.Width) - .pershare_usd.Width)
16830       .pershare.Top = .pershare_alt_line.Top
16840       .pershare_lbl.Top = (.pershare_alt_line.Top + lngTpp)
16850       .pershare_usd.Top = .pershare_alt_line.Top

16860       .total_cash_sum_usd.ForeColor = CLR_DISABLED_FG
16870       .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
16880       .Overshort_usd.ForeColor = CLR_DISABLED_FG
16890       .Overshort_usd.BackColor = CLR_DISABLED_BG
16900       .pershare_usd.ForeColor = CLR_DISABLED_FG
16910       .pershare_usd.BackColor = CLR_DISABLED_BG

16920       .maptype.Left = (.maptype.Left + lngTmp01)
16930       .cmdCancel.Left = (.cmdCancel.Left + lngTmp01)
16940       .cmdReinvest.Left = (.cmdReinvest.Left + lngTmp01)
16950       .cmdOK.Left = (.cmdOK.Left + lngTmp01)

16960       lngTmp01 = (.GoToReport_Emblem_01_img.Left + lngTmp01)
16970       For lngX = 1& To 24&
16980         .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Left = lngTmp01
16990       Next

17000     End If
17010   End With

EXITP:
17020   Exit Sub

ERRH:
17030   DoCmd.Hourglass False
17040   Select Case ERR.Number
        Case Else
17050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17060   End Select
17070   Resume EXITP

End Sub

Private Sub NoChar_Load()

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

17110   Set dbs = CurrentDb
17120   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
17130     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
17140     Set rst = qdf.OpenRecordset
17150     With rst
17160       .MoveLast
17170       lngNoChars = .RecordCount
17180       .MoveFirst
17190       arr_varNoChar = .GetRows(lngNoChars)
            ' *****************************************************
            ' ** Array: arr_varNoChar()
            ' **
            ' **   Field  Element  Name                Constant
            ' **   =====  =======  ==================  ==========
            ' **      1      0     curr_id             N_CID
            ' **      2      1     currsym_name        N_CNAM
            ' **      3      2     font_id             N_FID
            ' **      4      3     font_name           N_FNAM
            ' **      5      4     currfont3_nochar    N_NOCHR
            ' **
            ' *****************************************************
17200       .Close
17210     End With
17220     .Close
17230   End With

EXITP:
17240   Set rst = Nothing
17250   Set qdf = Nothing
17260   Set dbs = Nothing
17270   Exit Sub

ERRH:
17280   DoCmd.Hourglass False
17290   Select Case ERR.Number
        Case Else
17300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17310   End Select
17320   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

17410   With Me
17420     blnGTR_Emblem = False
17430     For lngX = 1& To 24&
17440       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
17450     Next
17460   End With

EXITP:
17470   Exit Sub

ERRH:
17480   Select Case ERR.Number
        Case Else
17490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17500   End Select
17510   Resume EXITP

End Sub
