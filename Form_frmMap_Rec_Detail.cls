VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Rec_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Rec_Detail"

'VGC 08/30/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   IncludeCurrency_Sub()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Received}
' **   OK:               O {cmdOK}
' **   Print Report:     P {cmdPrintReport}
' **   Reinvest:         R {cmdReinvest}
' **   Tax Codes:        T {cmbTaxCodes_Received}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots}

' ** MapType = "Received"

'rptMap_Received
'qryMapReport_03

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
Private strAccountNo_OU As String, lngAssetNo As Long, dblOverUnder As Double, dblPerShare As Double
Private datTransDate As Date, lngTaxcode As Long, datAssetDate As Date
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
150           If CharCnt(strTmp01, "~") = 6& Then  ' ** Module Function: modStringFuncs.

160             blnIsOpen = True

                ' ** strCallingForm & "~" & cstr(lngAssetNo) & "~" & CStr(dblOverUnder) & "~" & CStr(dblPerShare) & "~" & _
                ' **   Format(datTransDate, "mm/dd/yyyy") & "~" & Format(datAssetDate, "mm/dd/yyyy hh:nn:ss")
170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             lngAssetNo = Val(Left(strTmp01, (intPos01 - 1)))
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             dblOverUnder = Val(Left(strTmp01, (intPos01 - 1)))
250             strTmp01 = Mid(strTmp01, (intPos01 + 1))
260             intPos01 = InStr(strTmp01, "~")
270             dblPerShare = Val(Left(strTmp01, (intPos01 - 1)))
280             strTmp01 = Mid(strTmp01, (intPos01 + 1))
290             intPos01 = InStr(strTmp01, "~")
300             datTransDate = CDate(Left(strTmp01, (intPos01 - 1)))
310             strTmp01 = Mid(strTmp01, (intPos01 + 1))
320             intPos01 = InStr(strTmp01, "~")
330             datAssetDate = CDate(Left(strTmp01, (intPos01 - 1)))
340             lngCurrID = Val(Mid(strTmp01, (intPos01 + 1)))
350             lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left
                ' ** gstrAccountNo has specified account, otherwise vbNullString.

360             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
370             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

                ' ** Variables are fed empty, then populated ByRef.
380             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

390             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

400             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
410             If lngMonitorCnt > 1& Then
420               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
430               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
440               If lngMonitorNum > 1& Then
450                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
460               End If
470             End If

480             CLR_DISABLED_FG = CLR_DKGRY
490             CLR_DISABLED_BG = CLR_LTTEAL

500             blnCurrID = Pref_CurrID  ' ** Module Function: modPreferenceFuncs.
510             .curr_id = lngCurrID
520             IncludeCurrency_Sub  ' ** Procedure: Below.

530             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
540             blnGTR_Emblem = False: lngGTR_Stat = 0&

550             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

560             Select Case gblnAccountNoWithType
                Case True
570               strAccountNo_OU = "99-INCOME O/U"
580             Case False
590               strAccountNo_OU = "INCOME O/U"
600             End Select

610             Set dbs = CurrentDb
620             With dbs
630               Select Case strCallingForm
                  Case "frmJournal"
                    ' ** tblJournal_Map_Staging2, linked to Account.
640                 Set qdf = .QueryDefs("qryMap_Rec_02_04")
650               Case "frmJournal_Columns"
                    ' ** Journal Map, grouped and summed, by accountno.
660                 Set qdf = .QueryDefs("qryJournal_Columns_35_05_04")
670               End Select
680               Set rst = qdf.OpenRecordset
690               With rst
700                 If .BOF = True And .EOF = True Then
710                   lngRecs = 0&
720                 Else
730                   .MoveLast
740                   lngRecs = .RecordCount
750                 End If
760                 .Close
770               End With
780               Set rst = Nothing
790               .Close
800             End With
810             Set dbs = Nothing
820             DoEvents

830             Select Case strCallingForm
                Case "frmJournal"
840               If .frmMap_Rec_Detail_Sub.Form.RecordSource <> "qryMap_Rec_02_04" Then .frmMap_Rec_Detail_Sub.Form.RecordSource = "qryMap_Rec_02_04"
850             Case "frmJournal_Columns"
860               If .frmMap_Rec_Detail_Sub.Form.RecordSource <> "qryJournal_Columns_35_05_04" Then .frmMap_Rec_Detail_Sub.Form.RecordSource = "qryJournal_Columns_35_05_04"
870             End Select

880             .maptype = "Received"
890             .txtNumberRecords = "Number of Records = " & CStr(lngRecs)
900             .overshort = dblOverUnder
910             .description = "Long Term Capital Gain"
920             .pershare = dblPerShare  '5 dec.

930             If .cmdReinvest.Enabled = True Then
940               .cmdReinvest.SetFocus
950             ElseIf .cmdOK.Enabled = True Then
960               .cmdOK.SetFocus
970             Else
980               .description.SetFocus
990             End If

1000            blnRecalcTots = True
1010            .TimerInterval = 500&

1020            DoEvents
1030            DoCmd.SelectObject acForm, THIS_NAME, False
1040            DoEvents

1050            blnIsOpen = False

1060          Else
1070            Cancel = -1
1080          End If
1090        Else
1100          Cancel = -1
1110        End If
1120      Else
1130        Cancel = -1
1140      End If
1150    End With

1160    If Cancel = -1 Then
1170      Beep
1180      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1190    End If

EXITP:
1200    Set rst = Nothing
1210    Set qdf = Nothing
1220    Set dbs = Nothing
1230    Exit Sub

ERRH:
1240    DoCmd.Hourglass False
1250    Select Case ERR.Number
        Case Else
1260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1270    End Select
1280    Resume EXITP

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

          ' ** Borrowing these variables from the Court Reports.
1400      If gstrCrtRpt_Version <> vbNullString Then
1410        .cusip = gstrCrtRpt_Ordinal
1420        .cusip.ForeColor = CLR_DISABLED_FG
1430        .cusip.BackColor = CLR_DISABLED_BG
1440        .cusip.Visible = True
1450        .totdesc = gstrCrtRpt_Version
1460        .totdesc.ForeColor = CLR_DISABLED_FG
1470        .totdesc.BackColor = CLR_DISABLED_BG
1480        .totdesc.Visible = True
1490      End If

1500      If gstrAccountNo <> vbNullString Then
            ' ** An account was specified.
1510        If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1520          lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1530        End If
1540        .cusip.Top = .totdesc.Top
            '.totdesc.Width = .shortname.Width
1550        .totdesc.Left = ((.cusip.Left + .cusip.Width) + (4& * lngTpp))
1560        .accountno = gstrAccountNo
1570        .shortname = gstrAccountName
1580        .accountno.Visible = True
1590        .shortname.Visible = True
1600      End If

1610      .transdate = datTransDate
1620      .assetdate = datAssetDate

1630      Select Case IsNull(.cmbRevenueCodes_Received)
          Case True
1640        .cmbRevenueCodes_Received = REVID_INC
1650        .cmbRevenueCodes_Received_display = Null
1660      Case False
1670        If .cmbRevenueCodes_Received = REVID_INC Or .cmbRevenueCodes_Received = REVID_EXP Then
1680          .cmbRevenueCodes_Received_display = Null
1690        Else
1700          .cmbRevenueCodes_Received_display = .cmbRevenueCodes_Received.Column(1)
1710        End If
1720      End Select

1730      Select Case IsNull(.cmbTaxCodes_Received)
          Case True
1740        .cmbTaxCodes_Received = TAXID_INC
1750        .cmbTaxCodes_Received_display = Null
1760      Case False
1770        If .cmbTaxCodes_Received = TAXID_INC Or .cmbTaxCodes_Received = TAXID_DED Then
1780          .cmbTaxCodes_Received_display = Null
1790        Else
1800          .cmbTaxCodes_Received_display = .cmbTaxCodes_Received.Column(1)
1810        End If
1820      End Select

1830      Select Case gblnRevenueExpenseTracking
          Case True
1840        .cmbRevenueCodes_Received.Visible = True
1850        .cmbRevenueCodes_Received_display.Visible = True
1860        .cmbRevenueCodes_Received_lbl2.Visible = False  ' ** Option off msg.
1870      Case False
1880        .cmbRevenueCodes_Received.Visible = False
1890        .cmbRevenueCodes_Received_display.Visible = False
1900        .cmbRevenueCodes_Received_lbl_box.Visible = False
1910        .cmbRevenueCodes_Received_lbl2.Visible = True  ' ** Option off msg.
1920      End Select

1930      Select Case gblnIncomeTaxCoding
          Case True
1940        .cmbTaxCodes_Received.Visible = True
1950        .cmbTaxCodes_Received_display.Visible = True
1960        .cmbTaxCodes_Received_lbl2.Visible = False  ' ** Option off msg.
1970      Case False
1980        .cmbTaxCodes_Received.Visible = False
1990        .cmbTaxCodes_Received_display.Visible = False
2000        .cmbTaxCodes_Received_lbl_box.Visible = False  ' ** Option off msg.
2010        .cmbTaxCodes_Received_lbl2.Visible = True
2020      End Select

2030      If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
2040        .cmbRevenueCodes_Received_box_lbl.ForeColor = WIN_CLR_DISF
2050        .cmbRevenueCodes_Received_box_lbl_dim_hi.Visible = True
2060      End If

2070      curr_id_AfterUpdate  ' ** Procedure: Below.
2080      IncludeCurrency False  ' ** Procedure: Below
2090      DoEvents

2100      DoCmd.Hourglass False
2110      .frmMap_Rec_Detail_Sub.Form.TimerInterval = 500&

2120    End With

EXITP:
2130    Exit Sub

ERRH:
2140    DoCmd.Hourglass False
2150    Select Case ERR.Number
        Case Else
2160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2170    End Select
2180    Resume EXITP

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

2630      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
2640        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
2650      End If

2660      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
2670        lngGTR_Stat = lngGTR_Stat + 1&
2680        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
2690        lngTmp01 = lngTmp01 + 1&
2700        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
2710        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
2720        If lngTmp01 > 1& Then
2730          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
2740        End If
2750        If lngTmp01 = 24& Then
2760          lngTmp01 = 0&
2770          lngTmp02 = lngTmp02 + 1&
2780          If lngTmp02 < 8& Then
2790            .TimerInterval = 50&
2800          Else
2810            blnGTR_Emblem = False
2820          End If
2830        Else
2840          .TimerInterval = 50&
2850        End If
2860      End If  ' ** blnGTR_Emblem.
2870      If gblnGoToReport = True Then
2880        DoCmd.Hourglass True  ' ** Make sure it's still running.
2890        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or _
                (lngGTR_Stat > 96& And lngGTR_Stat < 120&) Or (lngGTR_Stat > 120&)) Then
              ' ** Just let the emblem go.
2900        Else
2910          If blnGTR_Emblem = False Then
2920            DoCmd.Hourglass True  ' ** Make sure it's still running.
2930            blnGTR_Emblem = True
2940            If IsNull(garr_varGoToReport(GTR_CTL4)) = False Then
2950              lngTmp01 = 0&: lngTmp02 = 0&
2960              lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
2970              .GoToReport_arw_map_okr_img.Left = (.GoToReport_arw_map_okr_img.Left - lngTmp03)
2980              .cmdOK.Left = (.cmdOK.Left - lngTmp03)
2990              .cmdReinvest.Left = (.cmdReinvest.Left - lngTmp03)
3000              .cmdCancel.Left = (.cmdCancel.Left - lngTmp03)
3010            End If
3020          End If
3030          DoEvents
3040          Select Case blnGoingToReport
              Case True
3050            Select Case IsNull(garr_varGoToReport(GTR_CTL4))
                Case True
3060              DoCmd.Hourglass False
3070              gblnGoToReport = False
3080              blnGoingToReport = False
3090              .GoToReport_arw_pr_img.Visible = False
3100              .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
3110              GTREmblem_Off  ' ** Procedure: Below.
3120              .cmdCancel.SetFocus
3130              .TimerInterval = 0&
3140            Case False
3150              .TimerInterval = 50&
3160              cmdReinvest_Click  ' ** Procedure: Below.
3170            End Select
3180          Case False
3190            DoCmd.Hourglass True  ' ** Make sure it's still running.
3200            DoEvents
3210            blnGoingToReport = True
3220            Select Case IsNull(garr_varGoToReport(GTR_CTL4))
                Case True
3230              .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3240              .GoToReport_arw_pr_img.Visible = True
3250              If .cmdPrintReport.Enabled = True Then
3260                .cmdPrintReport.SetFocus
3270              End If
3280              DoEvents
                  ' ** Make sure this remains off!
3290              GTREmblem_Off  ' ** Procedure: Below.
3300              Beep
3310              DoCmd.Hourglass False
3320              DoEvents
3330              .TimerInterval = (GTR_WAIT * 2&)
3340            Case False
3350              lngBtn_Top = (.cmdReinvest.Top + (lngTpp))
3360              lngBtn_Left = ((.cmdReinvest.Left - .GoToReport_arw_map_okr_img.Width) + (2& * lngTpp))
3370              .GoToReport_arw_map_okr_img.Top = lngBtn_Top
3380              .GoToReport_arw_map_okr_img.Left = lngBtn_Left
3390              .GoToReport_arw_map_okr_img.Visible = True
3400              .cmdReinvest.SetFocus
3410              DoEvents
3420              .TimerInterval = 50&
3430            End Select
3440          End Select
3450        End If
3460      End If
3470      If blnRecalcTots = True Then
3480        blnRecalcTots = False
3490        blnRetVal = RecalcTots  ' ** Function: Below.
3500        If blnRetVal = False Then
3510          blnRecalcTots = True
3520          .TimerInterval = 250&
3530        End If
3540      End If
3550    End With

EXITP:
3560    Exit Sub

ERRH:
3570    DoCmd.Hourglass False
3580    Select Case ERR.Number
        Case Else
3590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3600    End Select
3610    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3710    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Assure all reports are closed.
3720    Do While Reports.Count > 0
3730      DoCmd.Close acReport, Reports(0).Name
3740    Loop

3750    gstrReportCallingForm = vbNullString

EXITP:
3760    Exit Sub

ERRH:
3770    DoCmd.Hourglass False
3780    Select Case ERR.Number
        Case Else
3790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3800    End Select
3810    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3910    With Me
3920      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3930        Select Case blnPrintReport_Focus
            Case True
3940          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3950          .cmdPrintReport_raised_img.Visible = False
3960        Case False
3970          .cmdPrintReport_raised_img.Visible = True
3980          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3990        End Select
4000        .cmdPrintReport_raised_focus_img.Visible = False
4010        .cmdPrintReport_raised_focus_dots_img.Visible = False
4020        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4030        .cmdPrintReport_raised_img_dis.Visible = False
4040      End If
4050    End With

EXITP:
4060    Exit Sub

ERRH:
4070    Select Case ERR.Number
        Case Else
4080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4090    End Select
4100    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4220    Exit Sub

ERRH:
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Private Sub cmdOK_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRevCode As Long, lngRevType As Long, lngTaxType As Long
        Dim strRevDesc As String, strTaxDesc As String
        Dim varComments As Variant

4310    DoCmd.Hourglass True
4320    DoEvents

4330    With Me
4340      lngRevCode = .cmbRevenueCodes_Received
4350      strRevDesc = .cmbRevenueCodes_Received.Column(1)
4360      lngRevType = .cmbRevenueCodes_Received.Column(2)
4370      lngTaxcode = .cmbTaxCodes_Received
4380      strTaxDesc = .cmbTaxCodes_Received.Column(1)
4390      lngTaxType = .cmbTaxCodes_Received.Column(2)
4400      varComments = .description
4410    End With

4420    Set dbs = CurrentDb
4430    With dbs

4440      Select Case strCallingForm
          Case "frmJournal"
            ' ** Append tblJournal_Map_Staging2 to Journal, by specified [desc], [jsub], [revcode],
            ' ** [taxcode], [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id
4450        Set qdf = .QueryDefs("qryMap_Rec_03a")
4460        With qdf.Parameters
4470          ![astno] = lngAssetNo
4480          ![actno] = strAccountNo_OU
4490        End With
4500      Case "frmJournal_Columns"
            ' ** Append tblJournal_Map_Staging2 to tblJournal_Column, by specified [desc],
            ' ** [jsub], [revcod], [taxcod], [revdesc], [taxdesc], [revtyp], [taxtyp],
            ' ** [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id.
4510        Set qdf = .QueryDefs("qryJournal_Columns_35_06a")
4520        With qdf.Parameters
4530          ![revdesc] = strRevDesc
4540          ![taxdesc] = strTaxDesc
4550          ![revtyp] = lngRevType
4560          ![taxtyp] = lngTaxType
4570        End With
4580      End Select
4590      With qdf.Parameters
4600        ![Desc] = varComments
4610        ![jsub] = "LTCG"  ' ** journalSubtype (VGC: 02/15/09).
4620        ![revcod] = lngRevCode
4630        ![taxcod] = lngTaxcode
4640        ![reinvstd] = False
4650      End With
4660      qdf.Execute
4670      Set qdf = Nothing
4680      DoEvents

4690      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
4700        Set qdf = .QueryDefs("qryMap_01")
4710      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
4720        Set qdf = .QueryDefs("qryJournal_Columns_35_04")
4730      End Select
4740      qdf.Execute
4750      Set qdf = Nothing
4760      DoEvents

4770      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty tblJournal_Map_Staging2.
4780        Set qdf = .QueryDefs("qryMap_Rec_02_16")
4790      Case "frmJournal_Columns"
            ' ** Empty tblJournal_Map_Staging2.
4800        Set qdf = .QueryDefs("qryJournal_Columns_35_05_16")
4810      End Select
4820      qdf.Execute
4830      Set qdf = Nothing

4840      .Close
4850    End With
4860    Set dbs = Nothing
4870    DoEvents

4880    Select Case strCallingForm
        Case "frmJournal"
4890      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4900        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
4910      Else
4920        Forms(strCallingForm).frmJournal_Sub5_Misc.Form.Requery
4930        DoCmd.SelectObject acForm, strCallingForm, False
4940      End If
4950      DoEvents
4960      Forms(strCallingForm).frmJournal_Sub5_Misc.Form.Requery
4970      Forms(strCallingForm).frmJournal_Sub5_Misc.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub5_Misc.
4980    Case "frmJournal_Columns"
4990      gblnSetFocus = True
5000      gstrReturningForm = "frmMap_Return"
5010      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5020        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
5030      Else
5040        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
5050        Forms(strCallingForm).frmJournal_Columns_Sub.Form.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
5060        DoCmd.SelectObject acForm, strCallingForm, False
5070        Forms(strCallingForm).TimerInterval = 250&
5080      End If
5090    End Select

5100    DoCmd.Close acForm, THIS_NAME

EXITP:
5110    DoCmd.Hourglass False
5120    Set qdf = Nothing
5130    Set dbs = Nothing
5140    Exit Sub

ERRH:
5150    DoCmd.Hourglass False
5160    Select Case ERR.Number
        Case Else
5170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5180    End Select
5190    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
5290          .cmdReinvest.SetFocus
5300        End With
5310      End Select
5320    End If

        ' ** Shift keys.
5330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5340      Select Case intRetVal
          Case vbKeyTab
5350        With Me
5360          intRetVal = 0
5370          .cmdPrintReport.SetFocus
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

Private Sub cmdReinvest_Click()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReinvest_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRevCode As Long, lngRevType As Long, lngTaxType As Long
        Dim strRevDesc As String, strTaxDesc As String
        Dim varComments As Variant
        Dim strDocName As String

5510    DoCmd.Hourglass True
5520    DoEvents

5530    With Me
5540      lngRevCode = .cmbRevenueCodes_Received
5550      strRevDesc = .cmbRevenueCodes_Received.Column(1)
5560      lngRevType = .cmbRevenueCodes_Received.Column(2)
5570      lngTaxcode = .cmbTaxCodes_Received
5580      strTaxDesc = .cmbTaxCodes_Received.Column(1)
5590      lngTaxType = .cmbTaxCodes_Received.Column(2)
5600      varComments = .description
5610    End With

5620    Set dbs = CurrentDb
5630    With dbs

5640      Select Case strCallingForm
          Case "frmJournal"
            ' ** Append tblJournal_Map_Staging2 to Journal, by specified [desc], [jsub],
            ' ** [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id
5650        Set qdf = .QueryDefs("qryMap_Rec_03a")
5660        With qdf.Parameters
5670          ![astno] = lngAssetNo
5680          ![actno] = strAccountNo_OU
5690        End With
5700      Case "frmJournal_Columns"
            ' ** Append tblJournal_Map_Staging2 to tblJournal_Column, by specified
            ' ** [desc], [jsub], [revcod], [taxcod], [revdesc], [taxdesc], [revtyp], [taxtyp],
            ' ** GlobalVarGet('gblnGoToReport').  #curr_id
5710        Set qdf = .QueryDefs("qryJournal_Columns_35_06a")
5720        With qdf.Parameters
5730          ![revdesc] = strRevDesc
5740          ![taxdesc] = strTaxDesc
5750          ![revtyp] = lngRevType
5760          ![taxtyp] = lngTaxType
5770        End With
5780      End Select
5790      With qdf.Parameters
5800        ![Desc] = varComments
5810        ![jsub] = "LTCG"  ' ** journalSubtype (VGC: 02/15/09).
5820        ![revcod] = lngRevCode
5830        ![taxcod] = lngTaxcode
5840        ![reinvstd] = True  ' ** It will only update the appropriate side.
5850      End With
5860      qdf.Execute
5870      Set qdf = Nothing
5880      DoEvents

          ' ** Empty tblJournal_Map_Staging3.
5890      Set qdf = .QueryDefs("qryMap_Reinvest_05_07")
5900      qdf.Execute
5910      Set qdf = Nothing
5920      DoEvents

          ' ** Append Journal Map to tblJournal_Map_Staging3.
5930      Set qdf = .QueryDefs("qryMap_Reinvest_05_04")
5940      qdf.Execute
5950      Set qdf = Nothing
5960      DoEvents

5970      Select Case strCallingForm
          Case "frmJournal"
            ' ** Update qryMap_Reinvest_05_05 (tblJournal_Map_Staging3,
            ' ** linked to Journal, with Journal_ID_new).
5980        Set qdf = .QueryDefs("qryMap_Reinvest_05_06")
5990      Case "frmJournal_Columns"
            ' ** Update qryJournal_Columns_30_14 (tblJournal_Map_Staging3,
            ' ** linked to tblJournal_Column, with JrnlCol_ID_new).
6000        Set qdf = .QueryDefs("qryJournal_Columns_30_15")
6010      End Select
6020      qdf.Execute
6030      Set qdf = Nothing
6040      DoEvents

6050      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
6060        Set qdf = .QueryDefs("qryMap_01")
6070      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
6080        Set qdf = .QueryDefs("qryJournal_Columns_35_04")
6090      End Select
6100      qdf.Execute
6110      Set qdf = Nothing
6120      DoEvents

6130      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty tblJournal_Map_Staging2.
6140        Set qdf = .QueryDefs("qryMap_Rec_02_16")
6150      Case "frmJournal_Columns"
            ' ** Empty tblJournal_Map_Staging2.
6160        Set qdf = .QueryDefs("qryJournal_Columns_35_05_16")
6170      End Select
6180      qdf.Execute
6190      Set qdf = Nothing

6200      .Close
6210    End With
6220    Set dbs = Nothing
6230    DoEvents

        'USES tmpAccount FOR SORTING, BUT WHERE DOES THAT GET POPULATED?
        'Empty tmpAccount.  frmMap_Rec
        'Append Account to tmpAccount.
        'Empty tmpAccount  UNLOAD!

        'Empty tblJournal_Map_Staging3.
        'Append Journal Map to tblJournal_Map_Staging3.
        'Empty Journal Map.
        'Empty tblJournal_Map_Staging2.
        'Append ActiveAssets to Journal Map
        'Append Journal Map to tblJournal_Map_Staging.
        'Cancel
        'Empty Journal Map.
        'Empty tblJournal_Map_Staging.
        'Empty tblJournal_Map_Staging3.

6240    If strCallingForm = "frmJournal_Columns" Then
6250      Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
6260    End If

6270    strDocName = "frmMap_Reinvest_Rec_Price"  ' ** NOTE: This will change gstrFormQuerySpec!
6280    DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & CStr(lngAssetNo) & "~" & CStr(dblPerShare) & "~" & _
          Format(datTransDate, "mm/dd/yyyy") & "~" & Format(datAssetDate, "mm/dd/yyyy hh:nn:ss") & "~" & CStr(lngCurrID)
6290    If gblnGoToReport = True Then
6300      Forms(strDocName).TimerInterval = 100&
6310    End If

6320    DoCmd.Close acForm, THIS_NAME

EXITP:
6330    DoCmd.Hourglass False
6340    Set qdf = Nothing
6350    Set dbs = Nothing
6360    Exit Sub

ERRH:
6370    DoCmd.Hourglass False
6380    Select Case ERR.Number
        Case Else
6390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6400    End Select
6410    Resume EXITP

End Sub

Private Sub cmdReinvest_KeyDown(KeyCode As Integer, Shift As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReinvest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6520    intShiftDown = (Shift And acShiftMask) > 0
6530    intAltDown = (Shift And acAltMask) > 0
6540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6560      Select Case intRetVal
          Case vbKeyTab
6570        With Me
6580          intRetVal = 0
6590          .cmdCancel.SetFocus
6600        End With
6610      End Select
6620    End If

        ' ** Shift keys.
6630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6640      Select Case intRetVal
          Case vbKeyTab
6650        With Me
6660          intRetVal = 0
6670          .cmdOK.SetFocus
6680        End With
6690      End Select
6700    End If

EXITP:
6710    KeyCode = intRetVal
6720    Exit Sub

ERRH:
6730    intRetVal = 0
6740    Select Case ERR.Number
        Case Else
6750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6760    End Select
6770    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intTmp01 As Integer
        Dim intX As Integer

6810    DoCmd.Hourglass True
6820    DoEvents

6830    intTmp01 = 0
6840    Do While Forms.Count > 1
6850      intTmp01 = intTmp01 + 1
6860      For intX = (Forms.Count - 1) To 0 Step -1
6870        Select Case Forms(intX).Name
            Case THIS_NAME
              ' ** Skip.
6880        Case "frmJournal", "frmJournal_Columns"
              ' ** Skip.
6890        Case Else
6900          DoCmd.Close acForm, Forms(intX).Name
6910          Exit For  ' ** Do them one-at-a-time.
6920        End Select
6930      Next
6940      If intTmp01 >= 10 Then Exit Do  ' ** Don't get caught in a loop.
6950    Loop

6960    Set dbs = CurrentDb
6970    With dbs
6980      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
6990        Set qdf = .QueryDefs("qryMap_01")
7000      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
7010        Set qdf = .QueryDefs("qryJournal_Columns_35_04")
7020      End Select
7030      qdf.Execute
7040      Set qdf = Nothing
7050      DoEvents
          ' ** Empty tmpAccount.
7060      Set qdf = dbs.QueryDefs("qryMap_Rec_08_01")
7070      qdf.Execute
7080      Set qdf = Nothing
7090      DoEvents
7100      .Close
7110    End With
7120    Set dbs = Nothing
7130    DoEvents

7140    gblnSetFocus = True
7150    Select Case strCallingForm
        Case "frmJournal"
7160      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
7170        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
7180      Else
7190        gstrReturningForm = THIS_NAME
7200        Forms(strCallingForm).frmJournal_Sub5_Misc.Form.Requery
7210        DoCmd.SelectObject acForm, strCallingForm, False
7220        Forms(strCallingForm).TimerInterval = 250&
7230      End If
7240    Case "frmJournal_Columns"
7250      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
7260        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
7270      Else
7280        gstrReturningForm = THIS_NAME
7290        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
7300        DoCmd.SelectObject acForm, strCallingForm, False
7310        Forms(strCallingForm).TimerInterval = 250&
7320      End If
7330    End Select

7340    DoCmd.Close acForm, THIS_NAME

EXITP:
7350    Set qdf = Nothing
7360    Set dbs = Nothing
7370    Exit Sub

ERRH:
7380    DoCmd.Hourglass False
7390    Select Case ERR.Number
        Case Else
7400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7410    End Select
7420    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7520    intShiftDown = (Shift And acShiftMask) > 0
7530    intAltDown = (Shift And acAltMask) > 0
7540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7560      Select Case intRetVal
          Case vbKeyTab
7570        With Me
7580          intRetVal = 0
7590          lngRecsCur = .frmMap_Rec_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Rec_Detail_Sub.
7600          If lngRecsCur > 0 Then
7610            .frmMap_Rec_Detail_Sub.SetFocus
7620            .frmMap_Rec_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMap_Rec_Detail_Sub.
7630            .frmMap_Rec_Detail_Sub.Form.FocusHolder.SetFocus
7640          Else
7650            .description.SetFocus
7660          End If
7670        End With
7680      End Select
7690    End If

        ' ** Shift keys.
7700    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7710      Select Case intRetVal
          Case vbKeyTab
7720        With Me
7730          intRetVal = 0
7740          .cmdReinvest.SetFocus
7750        End With
7760      End Select
7770    End If

EXITP:
7780    KeyCode = intRetVal
7790    Exit Sub

ERRH:
7800    intRetVal = 0
7810    Select Case ERR.Number
        Case Else
7820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7830    End Select
7840    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

7910    With Me
7920      blnPrintReport_Focus = True
7930      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
7940      .cmdPrintReport_raised_img.Visible = False
7950      .cmdPrintReport_raised_focus_img.Visible = False
7960      .cmdPrintReport_raised_focus_dots_img.Visible = False
7970      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7980      .cmdPrintReport_raised_img_dis.Visible = False
7990    End With

EXITP:
8000    Exit Sub

ERRH:
8010    Select Case ERR.Number
        Case Else
8020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8030    End Select
8040    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

8110    With Me
8120      blnPrintReport_MouseDown = True
8130      .cmdPrintReport_sunken_focus_dots_img.Visible = True
8140      .cmdPrintReport_raised_img.Visible = False
8150      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8160      .cmdPrintReport_raised_focus_img.Visible = False
8170      .cmdPrintReport_raised_focus_dots_img.Visible = False
8180      .cmdPrintReport_raised_img_dis.Visible = False
8190    End With

EXITP:
8200    Exit Sub

ERRH:
8210    Select Case ERR.Number
        Case Else
8220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8230    End Select
8240    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

        ' ** qryMapReport_03.
        'strSQL = "SELECT DISTINCTROW account.accountno, account.shortname, Sum([journal map].shareface) AS total_shareface, " & _
        '  "Sum([journal map].pcash) AS total_pcash, " & CoInfo & ", CStr([masterasset].[Description]) & " & _
        '  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy')) AS totdesc " & _
        '  "FROM ([journal map] INNER JOIN account ON [journal map].accountno = account.accountno) " & _
        '  "INNER JOIN masterasset ON [journal map].assetno = masterasset.assetno " & _
        '  "GROUP BY account.accountno, account.shortname, CStr([masterasset].[Description]) & " & _
        '  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & IIf([masterasset].[due] Is Not Null,'  Due ' & " & _
        '  "Format([masterasset].[due],'mm/dd/yyyy'));"

8310    With Me
8320      strDocName = "rptMap_Received"
8330      gstrReportCallingForm = THIS_NAME
8340      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
8350        .Modal = False
8360        DoCmd.OpenReport strDocName, acViewPreview, , , , acViewPreview
8370        DoCmd.Maximize
8380        DoCmd.RunCommand acCmdFitToWindow
8390        .Visible = False
8400      Else
            '##GTR_Ref: rptMap_Received
8410        DoCmd.OpenReport strDocName, acViewNormal, , , , acViewNormal
8420      End If
8430    End With

EXITP:
8440    Exit Sub

ERRH:
8450    Me.Visible = True
8460    DoCmd.Restore
8470    Select Case ERR.Number
        Case Else
8480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8490    End Select
8500    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8620    intShiftDown = (Shift And acShiftMask) > 0
8630    intAltDown = (Shift And acAltMask) > 0
8640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8660      Select Case intRetVal
          Case vbKeyTab
8670        With Me
8680          intRetVal = 0
8690          .cmdOK.SetFocus
8700        End With
8710      End Select
8720    End If

        ' ** Shift keys.
8730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8740      Select Case intRetVal
          Case vbKeyTab
8750        With Me
8760          intRetVal = 0
8770          If .cmbTaxCodes_Received.Visible = True And .cmbTaxCodes_Received.Enabled = True Then
8780            .cmbTaxCodes_Received.SetFocus
8790          ElseIf .cmbRevenueCodes_Received.Visible = True And .cmbRevenueCodes_Received.Enabled = True Then
8800            .cmbRevenueCodes_Received.SetFocus
8810          Else
8820            .description.SetFocus
8830          End If
8840        End With
8850      End Select
8860    End If

EXITP:
8870    KeyCode = intRetVal
8880    Exit Sub

ERRH:
8890    intRetVal = 0
8900    Select Case ERR.Number
        Case Else
8910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8920    End Select
8930    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

9010    With Me
9020      If blnPrintReport_MouseDown = False Then
9030        Select Case blnPrintReport_Focus
            Case True
9040          .cmdPrintReport_raised_focus_dots_img.Visible = True
9050          .cmdPrintReport_raised_focus_img.Visible = False
9060        Case False
9070          .cmdPrintReport_raised_focus_img.Visible = True
9080          .cmdPrintReport_raised_focus_dots_img.Visible = False
9090        End Select
9100        .cmdPrintReport_raised_img.Visible = False
9110        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9120        .cmdPrintReport_sunken_focus_dots_img.Visible = False
9130        .cmdPrintReport_raised_img_dis.Visible = False
9140      End If
9150    End With

EXITP:
9160    Exit Sub

ERRH:
9170    Select Case ERR.Number
        Case Else
9180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9190    End Select
9200    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

9310    With Me
9320      .cmdPrintReport_raised_focus_dots_img.Visible = True
9330      .cmdPrintReport_raised_img.Visible = False
9340      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9350      .cmdPrintReport_raised_focus_img.Visible = False
9360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
9370      .cmdPrintReport_raised_img_dis.Visible = False
9380      blnPrintReport_MouseDown = False
9390    End With

EXITP:
9400    Exit Sub

ERRH:
9410    Select Case ERR.Number
        Case Else
9420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9430    End Select
9440    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

9510    With Me
9520      .cmdPrintReport_raised_img.Visible = True
9530      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9540      .cmdPrintReport_raised_focus_img.Visible = False
9550      .cmdPrintReport_raised_focus_dots_img.Visible = False
9560      .cmdPrintReport_sunken_focus_dots_img.Visible = False
9570      .cmdPrintReport_raised_img_dis.Visible = False
9580      blnPrintReport_Focus = False
9590    End With

EXITP:
9600    Exit Sub

ERRH:
9610    Select Case ERR.Number
        Case Else
9620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9630    End Select
9640    Resume EXITP

End Sub

Private Sub total_shareface_sum_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_sum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9720    intShiftDown = (Shift And acShiftMask) > 0
9730    intAltDown = (Shift And acAltMask) > 0
9740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9770        With Me
9780          intRetVal = 0
9790          .description.SetFocus
9800        End With
9810      End Select
9820    End If

        ' ** Shift keys.
9830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9850        With Me
9860          intRetVal = 0
9870          lngRecsCur = .frmMap_Rec_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Rec_Detail_Sub.
9880          If lngRecsCur > 0& Then
9890            .frmMap_Rec_Detail_Sub.SetFocus
9900            .frmMap_Rec_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Rec_Detail_Sub.
9910            .frmMap_Rec_Detail_Sub.Form.FocusHolder.SetFocus
9920          Else
9930            .cmdCancel.SetFocus
9940          End If
9950        End With
9960      End Select
9970    End If

EXITP:
9980    KeyCode = intRetVal
9990    Exit Sub

ERRH:
10000   intRetVal = 0
10010   Select Case ERR.Number
        Case Else
10020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10030   End Select
10040   Resume EXITP

End Sub

Private Sub total_cash_sum_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "total_cash_sum_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
10170       With Me
10180         intRetVal = 0
10190         .description.SetFocus
10200       End With
10210     End Select
10220   End If

        ' ** Shift keys.
10230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10250       With Me
10260         intRetVal = 0
10270         lngRecsCur = .frmMap_Rec_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Rec_Detail_Sub.
10280         If lngRecsCur > 0& Then
10290           .frmMap_Rec_Detail_Sub.SetFocus
10300           .frmMap_Rec_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Rec_Detail_Sub.
10310           .frmMap_Rec_Detail_Sub.Form.FocusHolder.SetFocus
10320         Else
10330           .cmdCancel.SetFocus
10340         End If
10350       End With
10360     End Select
10370   End If

EXITP:
10380   KeyCode = intRetVal
10390   Exit Sub

ERRH:
10400   intRetVal = 0
10410   Select Case ERR.Number
        Case Else
10420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10430   End Select
10440   Resume EXITP

End Sub

Private Sub Overshort_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "Overshort_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10520   intShiftDown = (Shift And acShiftMask) > 0
10530   intAltDown = (Shift And acAltMask) > 0
10540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10570       With Me
10580         intRetVal = 0
10590         .description.SetFocus
10600       End With
10610     End Select
10620   End If

        ' ** Shift keys.
10630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10650       With Me
10660         intRetVal = 0
10670         lngRecsCur = .frmMap_Rec_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Rec_Detail_Sub.
10680         If lngRecsCur > 0& Then
10690           .frmMap_Rec_Detail_Sub.SetFocus
10700           .frmMap_Rec_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Rec_Detail_Sub.
10710           .frmMap_Rec_Detail_Sub.Form.FocusHolder.SetFocus
10720         Else
10730           .cmdCancel.SetFocus
10740         End If
10750       End With
10760     End Select
10770   End If

EXITP:
10780   KeyCode = intRetVal
10790   Exit Sub

ERRH:
10800   intRetVal = 0
10810   Select Case ERR.Number
        Case Else
10820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10830   End Select
10840   Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

10910   With Me
10920     Select Case .curr_date.Visible
          Case True
10930       .curr_date.Visible = False
10940     Case False
10950       .curr_date = .curr_id.Column(CBX_C_DATE)
10960       .curr_date.Visible = True
10970     End Select
10980   End With

EXITP:
10990   Exit Sub

ERRH:
11000   Select Case ERR.Number
        Case Else
11010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
11020   End Select
11030   Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String
        Dim lngX As Long

11110   With Me

11120     If IsNull(.curr_id) = False Then

11130       lngCurrID = .curr_id.Column(CBX_C_CURRID)

11140       blnFound = False
11150       If lngCurrID <> 150& Then
11160         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
11170           NoChar_Load  ' ** Procedure: Below.
11180         End If
11190         For lngX = 0& To (lngNoChars - 1&)
11200           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
11210             blnFound = True
11220           End If
11230         Next
11240       End If

11250       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
11260         .total_cash_sum.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
11270         .total_cash_sum.FontSize = 10
11280         .total_cash_sum.FontBold = True
11290         .overshort.FontName = "Courier New"
11300         .overshort.FontSize = 10
11310         .overshort.FontBold = True
11320         .pershare.FontName = "Courier New"
11330         .pershare.FontSize = 10
11340         .pershare.FontBold = True
11350       Case False
11360         If .total_cash_sum.FontName <> "Arial" Then
11370           .total_cash_sum.FontName = "Arial"
11380           .total_cash_sum.FontSize = 10
11390           .total_cash_sum.FontBold = False
11400           .overshort.FontName = "Arial"
11410           .overshort.FontSize = 10
11420           .overshort.FontBold = False
11430           .pershare.FontName = "Arial"
11440           .pershare.FontSize = 10
11450           .pershare.FontBold = False
11460         End If
11470       End Select

11480       If lngCurrID = 150& Then  ' ** USD.
11490         strFrmt1 = "Currency"
11500         intDec1 = 2
11510         strFrmt2 = "Currency"
11520         intDec2 = 5
11530         lngBkClr = CLR_DISABLED_BG
11540       Else
11550         strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
11560         If .curr_id.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
11570           intDec1 = 2
11580         Else
11590           intDec1 = .curr_id.Column(CBX_C_DEC)  ' ** Decimal places.
11600         End If
11610         intDec2 = 5
11620         intLen = Len(strTmp01)
11630         Select Case intLen
              Case 1
                'strTmp01 = strTmp01
11640         Case 2
11650           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
11660         Case 3
11670           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
11680         Case 4
11690           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
11700         Case 5
11710           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                  "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
11720         End Select
11730         Select Case intDec1
              Case 0
11740           strFrmt1 = strTmp01 & "#,##0;(" & strTmp01 & "#,##0);" & strTmp01 & "0;" & strTmp01 & "0"
11750           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11760         Case 1  ' ** None have this currently.
11770           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
11780           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11790         Case 2
11800           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
11810           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11820         Case 3
11830           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
11840           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11850         End Select
11860         lngBkClr = CLR_LTGRN
11870       End If
11880       .total_cash_sum.Format = strFrmt1
11890       .total_cash_sum.DecimalPlaces = intDec1
11900       .total_cash_sum.BackColor = lngBkClr
11910       .overshort.Format = strFrmt1
11920       .overshort.DecimalPlaces = intDec1
11930       .overshort.BackColor = lngBkClr
11940       .pershare.Format = strFrmt2
11950       .pershare.DecimalPlaces = intDec2
11960       .pershare.BackColor = lngBkClr

11970       If lngCurrID = 150& Then
11980         .total_cash_sum_usd.Visible = False
11990         .Overshort_usd.Visible = False
12000         .pershare_usd.Visible = False
12010       Else
12020         .total_cash_sum_usd = 0
12030         .total_cash_sum_usd.Visible = True
12040         .Overshort_usd = 0
12050         .Overshort_usd.Visible = True
12060         .pershare_usd = 0
12070         .pershare_usd.Visible = True
12080       End If

12090     End If

12100   End With

EXITP:
12110   Exit Sub

ERRH:
12120   Select Case ERR.Number
        Case Else
12130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12140   End Select
12150   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

12210   With Me
12220     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
12230     .curr_id.SetFocus
12240     Response = acDataErrContinue
12250   End With

EXITP:
12260   Exit Sub

ERRH:
12270   Select Case ERR.Number
        Case Else
12280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12290   End Select
12300   Resume EXITP

End Sub

Private Sub description_GotFocus()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

12410   With Me
12420     .description.SelLength = 0
12430     .description.SelStart = 999
12440   End With

EXITP:
12450   Exit Sub

ERRH:
12460   Select Case ERR.Number
        Case Else
12470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12480   End Select
12490   Resume EXITP

End Sub

Private Sub description_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12520   intShiftDown = (Shift And acShiftMask) > 0
12530   intAltDown = (Shift And acAltMask) > 0
12540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12570       With Me
12580         intRetVal = 0
12590         If .cmbRevenueCodes_Received.Visible = True And .cmbRevenueCodes_Received.Enabled = True Then
12600           .cmbRevenueCodes_Received.SetFocus
12610         ElseIf .cmbTaxCodes_Received.Visible = True And .cmbTaxCodes_Received.Enabled = True Then
12620           .cmbTaxCodes_Received.SetFocus
12630         Else
12640           .cmdReinvest.SetFocus   ' ** Yes, go to here.
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
12730         lngRecsCur = .frmMap_Rec_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Rec_Detail_Sub.
12740         If lngRecsCur > 0& Then
12750           .frmMap_Rec_Detail_Sub.SetFocus
12760           .frmMap_Rec_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Rec_Detail_Sub.
12770           .frmMap_Rec_Detail_Sub.Form.FocusHolder.SetFocus
12780         Else
12790           .cmdCancel.SetFocus
12800         End If
12810       End With
12820     End Select
12830   End If

EXITP:
12840   KeyCode = intRetVal
12850   Exit Sub

ERRH:
12860   intRetVal = 0
12870   Select Case ERR.Number
        Case Else
12880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12890   End Select
12900   Resume EXITP

End Sub

Private Sub description_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "description_MouseMove"

13010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
13020   Exit Sub

ERRH:
13030   Select Case ERR.Number
        Case Else
13040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13050   End Select
13060   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_display_GotFocus()

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_display_GotFocus"

13110   With Me
13120     .cmbRevenueCodes_Received.SetFocus
13130   End With

EXITP:
13140   Exit Sub

ERRH:
13150   Select Case ERR.Number
        Case Else
13160     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13170   End Select
13180   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_display_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13220   intShiftDown = (Shift And acShiftMask) > 0
13230   intAltDown = (Shift And acAltMask) > 0
13240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13270       With Me
13280         intRetVal = 0
13290         .cmbRevenueCodes_Received.SetFocus
13300       End With
13310     End Select
13320   End If

        ' ** Shift keys.
13330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13350       With Me
13360         intRetVal = 0
13370         .description.SetFocus
13380       End With
13390     End Select
13400   End If

EXITP:
13410   KeyCode = intRetVal
13420   Exit Sub

ERRH:
13430   intRetVal = 0
13440   Select Case ERR.Number
        Case Else
13450     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13460   End Select
13470   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_Enter()

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_Enter"

13510   With Me
13520     Select Case IsNull(.cmbRevenueCodes_Received)
          Case True
13530       ExpandCombo .cmbRevenueCodes_Received  ' ** Module Procedure: modUtilities.
13540     Case False
13550       If .cmbRevenueCodes_Received = REVID_INC Or .cmbRevenueCodes_Received = REVID_EXP Then
13560         ExpandCombo .cmbRevenueCodes_Received  ' ** Module Procedure: modUtilities.
13570       End If
13580     End Select
13590   End With

EXITP:
13600   Exit Sub

ERRH:
13610   Select Case ERR.Number
        Case Else
13620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13630   End Select
13640   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_AfterUpdate()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_AfterUpdate"

13710   With Me
13720     Select Case IsNull(.cmbRevenueCodes_Received)
          Case True
13730       .cmbRevenueCodes_Received = REVID_INC
13740     Case False
            ' ** Nothing at the moment.
13750     End Select
13760     If .cmbRevenueCodes_Received = REVID_INC Or .cmbRevenueCodes_Received = REVID_EXP Then
13770       .cmbRevenueCodes_Received_display = Null
13780     Else
13790       .cmbRevenueCodes_Received_display = .cmbRevenueCodes_Received.Column(1)
13800     End If
13810   End With

EXITP:
13820   Exit Sub

ERRH:
13830   Select Case ERR.Number
        Case Else
13840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13850   End Select
13860   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Received_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Received_KeyDown"

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
13990         If .cmbTaxCodes_Received.Visible = True And .cmbTaxCodes_Received.Enabled = True Then
14000           .cmbTaxCodes_Received.SetFocus
14010         Else
14020           .cmdReinvest.SetFocus
14030         End If
14040       End With
14050     End Select
14060   End If

        ' ** Shift keys.
14070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14090       With Me
14100         intRetVal = 0
14110         .description.SetFocus
14120       End With
14130     End Select
14140   End If

EXITP:
14150   KeyCode = intRetVal
14160   Exit Sub

ERRH:
14170   intRetVal = 0
14180   Select Case ERR.Number
        Case Else
14190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14200   End Select
14210   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Received_display_GotFocus()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_display_GotFocus"

14310   With Me
14320     .cmbTaxCodes_Received.SetFocus
14330   End With

EXITP:
14340   Exit Sub

ERRH:
14350   Select Case ERR.Number
        Case Else
14360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14370   End Select
14380   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Received_display_KeyDown(KeyCode As Integer, Shift As Integer)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14420   intShiftDown = (Shift And acShiftMask) > 0
14430   intAltDown = (Shift And acAltMask) > 0
14440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14470       With Me
14480         intRetVal = 0
14490         .cmbTaxCodes_Received.SetFocus
14500       End With
14510     End Select
14520   End If

        ' ** Shift keys.
14530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14550       With Me
14560         intRetVal = 0
14570         If .cmbRevenueCodes_Received.Visible = True And .cmbRevenueCodes_Received.Enabled = True Then
14580           .cmbRevenueCodes_Received.SetFocus
14590         Else
14600           .description.SetFocus
14610         End If
14620       End With
14630     End Select
14640   End If

EXITP:
14650   KeyCode = intRetVal
14660   Exit Sub

ERRH:
14670   intRetVal = 0
14680   Select Case ERR.Number
        Case Else
14690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14700   End Select
14710   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Received_Enter()

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_Enter"

14810   With Me
14820     Select Case IsNull(.cmbTaxCodes_Received)
          Case True
14830       ExpandCombo .cmbTaxCodes_Received  ' ** Module Procedure: modUtilities.
14840     Case False
14850       If .cmbTaxCodes_Received = TAXID_INC Or .cmbTaxCodes_Received = TAXID_DED Then
14860         ExpandCombo .cmbTaxCodes_Received  ' ** Module Procedure: modUtilities.
14870       End If
14880     End Select
14890   End With

EXITP:
14900   Exit Sub

ERRH:
14910   Select Case ERR.Number
        Case Else
14920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14930   End Select
14940   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Received_AfterUpdate()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_AfterUpdate"

15010   With Me
15020     Select Case IsNull(.cmbTaxCodes_Received)
          Case True
15030       .cmbTaxCodes_Received = TAXID_INC
15040     Case False
            ' ** Nothing at the moment.
15050     End Select
15060     If .cmbTaxCodes_Received = TAXID_INC Or .cmbTaxCodes_Received = TAXID_DED Then
15070       .cmbTaxCodes_Received_display = Null
15080     Else
15090       .cmbTaxCodes_Received_display = .cmbTaxCodes_Received.Column(1)
15100     End If
15110   End With

EXITP:
15120   Exit Sub

ERRH:
15130   Select Case ERR.Number
        Case Else
15140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15150   End Select
15160   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Received_KeyDown(KeyCode As Integer, Shift As Integer)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Received_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15220   intShiftDown = (Shift And acShiftMask) > 0
15230   intAltDown = (Shift And acAltMask) > 0
15240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15270       With Me
15280         intRetVal = 0
15290         .cmdReinvest.SetFocus
15300       End With
15310     End Select
15320   End If

        ' ** Shift keys.
15330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15350       With Me
15360         intRetVal = 0
15370         If .cmbRevenueCodes_Received.Visible = True And .cmbRevenueCodes_Received.Enabled = True Then
15380           .cmbRevenueCodes_Received.SetFocus
15390         Else
15400           .description.SetFocus
15410         End If
15420       End With
15430     End Select
15440   End If

EXITP:
15450   KeyCode = intRetVal
15460   Exit Sub

ERRH:
15470   intRetVal = 0
15480   Select Case ERR.Number
        Case Else
15490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15500   End Select
15510   Resume EXITP

End Sub

Public Function RecalcTots()
' ** I give up trying to get it from the subform!

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim dblTmp01 As Double, dblTmp02 As Double
        Dim lngX As Long
        Dim blnRetVal As Boolean

15610   With Me
15620     blnRetVal = False
15630     Set rst = .frmMap_Rec_Detail_Sub.Form.RecordsetClone
15640     With rst
15650       If .BOF = True And .EOF = True Then
              ' ** Ho-Hum...
15660       Else
15670         .MoveLast
15680         lngRecs = .RecordCount
15690         .MoveFirst
15700         dblTmp01 = 0#: dblTmp02 = 0#
15710         For lngX = 1& To lngRecs
15720           dblTmp01 = dblTmp01 + ![total_shareface]
15730           dblTmp02 = dblTmp02 + ![total_pcash]
15740           If lngX < lngRecs Then .MoveNext
15750         Next
15760       End If
15770       .Close
15780     End With
15790     If dblTmp01 <> 0# Then
15800       blnRetVal = True
15810     End If
15820     .total_shareface_sum = dblTmp01
15830     .total_cash_sum = dblTmp02
15840   End With

EXITP:
15850   Set rst = Nothing
15860   RecalcTots = blnRetVal
15870   Exit Function

ERRH:
15880   blnRetVal = False
15890   Select Case ERR.Number
        Case 3167  ' ** Record is deleted.
15900     blnRetVal = True   ' ** Other wise it'll get stuck in a loop.
15910   Case Else
15920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15930   End Select
15940   Resume EXITP

End Function

Public Sub IncludeCurrency(blnCalc As Boolean)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim dblRate2 As Double
        Dim blnFound As Boolean
        Dim dblTmp01 As Double

16010   With Me
16020     Select Case blnCalc
          Case True
16030       If blnCurrID = True Then
16040         blnFound = False
16050         Do While blnFound = False
16060           If IsNull(.total_shareface_sum) = False Then
16070             If .total_shareface_sum <> 0 Then
16080               blnFound = True
16090               dblRate2 = .curr_id.Column(CBX_C_RATE2)
16100               dblTmp01 = (.total_cash_sum * dblRate2)
16110               dblTmp01 = Round(dblTmp01, 2)
16120               .total_cash_sum_usd = dblTmp01
16130               dblTmp01 = (.overshort * dblRate2)
16140               dblTmp01 = Round(dblTmp01, 2)
16150               .Overshort_usd = dblTmp01
16160               dblTmp01 = (.pershare * dblRate2)
16170               dblTmp01 = Round(dblTmp01, 5)
16180               .pershare_usd = dblTmp01
16190               If lngCurrID <> 150& Then
16200                 .total_cash_sum_usd.Visible = True
16210                 .Overshort_usd.Visible = True
16220                 .pershare_usd.Visible = True
16230               End If
16240             End If
16250           End If
16260         Loop
16270       End If
16280     Case False
16290       .curr_id.Enabled = False
16300       .curr_id.Locked = True
16310       .curr_id.ForeColor = CLR_DISABLED_FG
16320       .curr_id.BackColor = CLR_VLTGRN
16330       If blnCurrID = True And lngCurrID <> 150& Then
16340         .curr_id.Visible = True
16350         .curr_id_cmd.Visible = True
16360         .curr_id_cmd.Enabled = True
16370       Else
16380         .curr_id.Visible = False
16390         .curr_id_cmd.Visible = False
16400         .curr_id_cmd.Enabled = False
16410         .cmbRevenueCodes_Received_vline05.Visible = False
16420         .cmbRevenueCodes_Received_vline06.Visible = False
16430         .cmbRevenueCodes_Received_box.Width = (.cmbRevenueCodes_Received_vline06.Left - .cmbRevenueCodes_Received_box.Left)
16440         .overshort.Top = .Overshort_alt_line.Top
16450         .Overshort_lbl.Top = (.Overshort_alt_line.Top + lngTpp)
16460         .pershare.Top = .pershare_alt_line.Top
16470         .pershare_lbl.Top = (.pershare_alt_line.Top + lngTpp)
16480       End If
16490       .total_cash_sum_usd.Visible = False
16500       .Overshort_usd.Visible = False
16510       .pershare_usd.Visible = False
16520     End Select
16530   End With

EXITP:
16540   Exit Sub

ERRH:
16550   DoCmd.Hourglass False
16560   Select Case ERR.Number
        Case Else
16570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16580   End Select
16590   Resume EXITP

End Sub

Private Sub IncludeCurrency_Sub()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long

16610   With Me
16620     If blnCurrID = True And lngCurrID <> 150& Then

16630       With .frmMap_Rec_Detail_Sub.Form
16640         lngTmp01 = .Width - .total_pcash_alt_line.Left
16650       End With

16660       lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
16670       lngMonitorNum = 1&: lngTmp02 = 0&
16680       EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
16690       If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

16700       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
16710       DoCmd.MoveSize (lngFrm_Left - (lngTmp01 / 2)), lngTmp02, (lngFrm_Width + lngTmp01), lngFrm_Height  'lngFrm_Top
16720       If lngMonitorNum > 1& Then
16730         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
16740       End If

16750       .Width = (.Width + lngTmp01)
16760       .frmMap_Rec_Detail_Sub.Width = (.frmMap_Rec_Detail_Sub.Width + lngTmp01)
16770       .frmMap_Rec_Detail_Sub_box.Width = (.frmMap_Rec_Detail_Sub_box.Width + lngTmp01)

16780       .Header_vline01.Left = .Width
16790       .Header_vline02.Left = .Width
16800       .Footer_vline01.Left = .Width
16810       .Footer_vline02.Left = .Width
16820       .Header_hline01.Width = .Width
16830       .Header_hline02.Width = .Width
16840       .Footer_hline01.Width = .Width
16850       .Footer_hline02.Width = .Width

16860       .total_cash_sum_usd.Left = ((.total_cash_sum.Left + .total_cash_sum.Width) + (4& * lngTpp))
16870       .total_cash_sum_usd.Top = .total_cash_sum.Top
16880       .Overshort_usd.Left = ((.overshort.Left + .overshort.Width) + (4& * lngTpp))
16890       .overshort.Top = .Overshort_alt_line.Top
16900       .Overshort_lbl.Top = (.Overshort_alt_line.Top + lngTpp)
16910       .Overshort_usd.Top = .Overshort_alt_line.Top
16920       .pershare_usd.Left = ((.total_cash_sum_usd.Left + .total_cash_sum_usd.Width) - .pershare_usd.Width)
16930       .pershare.Top = .pershare_alt_line.Top
16940       .pershare_lbl.Top = (.pershare_alt_line.Top + lngTpp)
16950       .pershare_usd.Top = (.pershare_alt_line.Top + lngTpp)

16960       .total_cash_sum_usd.ForeColor = CLR_DISABLED_FG
16970       .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
16980       .Overshort_usd.ForeColor = CLR_DISABLED_FG
16990       .Overshort_usd.BackColor = CLR_DISABLED_BG
17000       .pershare_usd.ForeColor = CLR_DISABLED_FG
17010       .pershare_usd.BackColor = CLR_DISABLED_BG

17020       .maptype.Left = (.maptype.Left + lngTmp01)
17030       .cmdCancel.Left = (.cmdCancel.Left + lngTmp01)
17040       .cmdReinvest.Left = (.cmdReinvest.Left + lngTmp01)
17050       .cmdOK.Left = (.cmdOK.Left + lngTmp01)

17060       lngTmp01 = (.GoToReport_Emblem_01_img.Left + lngTmp01)
17070       For lngX = 1& To 24&
17080         .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Left = lngTmp01
17090       Next

17100     End If
17110   End With

EXITP:
17120   Exit Sub

ERRH:
17130   DoCmd.Hourglass False
17140   Select Case ERR.Number
        Case Else
17150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17160   End Select
17170   Resume EXITP

End Sub

Private Sub NoChar_Load()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

17210   Set dbs = CurrentDb
17220   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
17230     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
17240     Set rst = qdf.OpenRecordset
17250     With rst
17260       .MoveLast
17270       lngNoChars = .RecordCount
17280       .MoveFirst
17290       arr_varNoChar = .GetRows(lngNoChars)
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
17300       .Close
17310     End With
17320     .Close
17330   End With

EXITP:
17340   Set rst = Nothing
17350   Set qdf = Nothing
17360   Set dbs = Nothing
17370   Exit Sub

ERRH:
17380   DoCmd.Hourglass False
17390   Select Case ERR.Number
        Case Else
17400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17410   End Select
17420   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

17510   With Me
17520     blnGTR_Emblem = False
17530     For lngX = 1& To 24&
17540       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
17550     Next
17560   End With

EXITP:
17570   Exit Sub

ERRH:
17580   Select Case ERR.Number
        Case Else
17590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17600   End Select
17610   Resume EXITP

End Sub
