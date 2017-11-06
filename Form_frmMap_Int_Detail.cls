VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Int_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Int_Detail"

'VGC 08/30/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   IncludeCurrency_Sub()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Interest}
' **   OK:               O {cmdOK}
' **   Print Report:     P {cmdPrintReport}
' **   Reinvest:         R {cmdReinvest}
' **   Tax Codes:        T {cmbTaxCodes_Interest}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots}

' ** MapType = "Interest"

'rptMap_Interest
'qryMapReport_02

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

                ' ** strCallingForm & "~" & dblOverUnder & "~" & dblPerShare & "~" & intTaxCode & "~" &
                ' ** lngAssetNo & "~" & dblShareface & "~" & dblAmount & "~" & lngCurrID
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
360             dblAmount = Val(Left(strTmp01, (intPos01 - 1)))
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
620                 Set qdf = .QueryDefs("qryMap_Int_02_04")
630               Case "frmJournal_Columns"
                    ' ** Journal Map, grouped and summed, by accountno.
640                 Set qdf = .QueryDefs("qryJournal_Columns_31_05_04")
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
790             Set rst = Nothing
800             DoEvents

810             Select Case strCallingForm
                Case "frmJournal"
820               If .frmMap_Int_Detail_Sub.Form.RecordSource <> "qryMap_Int_02_04" Then .frmMap_Int_Detail_Sub.Form.RecordSource = "qryMap_Int_02_04"
830             Case "frmJournal_Columns"
840               If .frmMap_Int_Detail_Sub.Form.RecordSource <> "qryJournal_Columns_31_05_04" Then .frmMap_Int_Detail_Sub.Form.RecordSource = "qryJournal_Columns_31_05_04"
850             End Select

860             .maptype = "Interest"
870             .txtNumberRecords = "Number of Records = " & CStr(lngRecs)
880             .overshort = dblOverUnder
890             .pershare = dblPerShare    ' ** Rate %, 3 dec.

900             If .cmdReinvest.Enabled = True Then
910               .cmdReinvest.SetFocus
920             ElseIf .cmdOK.Enabled = True Then
930               .cmdOK.SetFocus
940             Else
950               .description.SetFocus
960             End If

970             blnRecalcTots = True
980             .TimerInterval = 500&

990             DoEvents
1000            DoCmd.SelectObject acForm, THIS_NAME, False
1010            DoEvents

1020            blnIsOpen = False

1030          Else
1040            Cancel = -1
1050          End If
1060        Else
1070          Cancel = -1
1080        End If
1090      Else
1100        Cancel = -1
1110      End If
1120    End With

1130    If Cancel = -1 Then
1140      Beep
1150      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1160    End If

EXITP:
1170    Set rst = Nothing
1180    Set qdf = Nothing
1190    Set dbs = Nothing
1200    Exit Sub

ERRH:
1210    DoCmd.Hourglass False
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

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

1500      .transdate = gdatStartDate
1510      .cmbTaxCodes_Interest = lngTaxcode

1520      Select Case IsNull(.cmbRevenueCodes_Interest)
          Case True
1530        .cmbRevenueCodes_Interest = REVID_INC
1540        .cmbRevenueCodes_Interest_display = Null
1550      Case False
1560        If .cmbRevenueCodes_Interest = REVID_INC Or .cmbRevenueCodes_Interest = REVID_EXP Then
1570          .cmbRevenueCodes_Interest_display = Null
1580        Else
1590          .cmbRevenueCodes_Interest_display = .cmbRevenueCodes_Interest.Column(1)
1600        End If
1610      End Select

1620      Select Case IsNull(.cmbTaxCodes_Interest)
          Case True
1630        .cmbTaxCodes_Interest = TAXID_INC
1640        .cmbTaxCodes_Interest_display = Null
1650      Case False
1660        If .cmbTaxCodes_Interest = TAXID_INC Or .cmbTaxCodes_Interest = TAXID_DED Then
1670          .cmbTaxCodes_Interest_display = Null
1680        Else
1690          .cmbTaxCodes_Interest_display = .cmbTaxCodes_Interest.Column(1)
1700        End If
1710      End Select

1720      Select Case gblnRevenueExpenseTracking
          Case True
1730        .cmbRevenueCodes_Interest.Visible = True
1740        .cmbRevenueCodes_Interest_display.Visible = True
1750        .cmbRevenueCodes_Interest_lbl2.Visible = False  ' ** Option off msg.
1760      Case False
1770        .cmbRevenueCodes_Interest.Visible = False
1780        .cmbRevenueCodes_Interest_display.Visible = False
1790        .cmbRevenueCodes_Interest_lbl_box.Visible = False
1800        .cmbRevenueCodes_Interest_lbl2.Visible = True  ' ** Option off msg.
1810      End Select

1820      Select Case gblnIncomeTaxCoding
          Case True
1830        .cmbTaxCodes_Interest.Visible = True
1840        .cmbTaxCodes_Interest_display.Visible = True
1850        .cmbTaxCodes_Interest_lbl2.Visible = False  ' ** Option off msg.
1860      Case False
1870        .cmbTaxCodes_Interest.Visible = False
1880        .cmbTaxCodes_Interest_display.Visible = False
1890        .cmbTaxCodes_Interest_lbl_box.Visible = False  ' ** Option off msg.
1900        .cmbTaxCodes_Interest_lbl2.Visible = True
1910      End Select

1920      If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
1930        .cmbRevenueCodes_Interest_box_lbl.ForeColor = WIN_CLR_DISF
1940        .cmbRevenueCodes_Interest_box_lbl_dim_hi.Visible = True
1950      End If

1960      curr_id_AfterUpdate  ' ** Procedure: Below.
1970      IncludeCurrency False  ' ** Procedure: Below
1980      DoEvents

1990      DoCmd.Hourglass False
2000      .frmMap_Int_Detail_Sub.Form.TimerInterval = 500&

2010    End With

EXITP:
2020    Exit Sub

ERRH:
2030    DoCmd.Hourglass False
2040    Select Case ERR.Number
        Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Comments:         M {description}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc Totals:    F5 {RecalcTots}

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyEscape
2170        intRetVal = 0
2180        cmdCancel_Click  ' ** Procedure: Below
2190      Case vbKeyF5
2200        intRetVal = 0
2210        RecalcTots  ' ** Function: Below.
2220      End Select
2230    End If

        ' ** Ctrl keys.
2240    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2250      Select Case intRetVal
          Case vbKeyM
2260        With Me
2270          intRetVal = 0
2280          .description.SetFocus
2290        End With
2300      End Select
2310    End If

        ' ** Ctrl-Shift keys.
2320    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2330      Select Case intRetVal
          Case vbKeyF
2340        With Me
2350          intRetVal = 0
2360          .FocusHolder.SetFocus
2370        End With
2380      End Select
2390    End If

EXITP:
2400    KeyCode = intRetVal
2410    Exit Sub

ERRH:
2420    intRetVal = 0
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub Form_Timer()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim lngBtn_Top As Long, lngBtn_Left As Long
        Dim lngTmp03 As Long
        Dim blnRetVal As Boolean

        Static lngTmp01 As Long, lngTmp02 As Long

2510    With Me
2520      .TimerInterval = 0&
2530      If blnGTR_Emblem = True Then
            ' ** GTR_WAIT = 1500&
2540        lngGTR_Stat = lngGTR_Stat + 1&
2550        If lngTmp01 = 0& And lngTmp02 = 0& Then .GoToReport_Emblem_01_img.Visible = True
2560        lngTmp01 = lngTmp01 + 1&
2570        .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01), 2) & "_img").Visible = True
2580        If lngTmp01 = 1 Then .GoToReport_Emblem_24_img.Visible = False
2590        If lngTmp01 > 1& Then
2600          .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_img").Visible = False
2610        End If
2620        If lngTmp01 = 24& Then
2630          lngTmp01 = 0&
2640          lngTmp02 = lngTmp02 + 1&
2650          If lngTmp02 < 8& Then
2660            .TimerInterval = 50&
2670          Else
2680            blnGTR_Emblem = False
2690          End If
2700        Else
2710          .TimerInterval = 50&
2720        End If
2730      End If  ' ** blnGTR_Emblem.
2740      If gblnGoToReport = True Then
2750        DoCmd.Hourglass True  ' ** Make sure it's still running.
2760        If ((lngGTR_Stat > 0& And lngGTR_Stat < 48&) Or (lngGTR_Stat > 48& And lngGTR_Stat < 96&) Or _
                (lngGTR_Stat > 96& And lngGTR_Stat < 120&) Or (lngGTR_Stat > 120&)) Then
              ' ** Just let the emblem go.
2770        Else
2780          If blnGTR_Emblem = False Then
2790            DoCmd.Hourglass True  ' ** Make sure it's still running.
2800            blnGTR_Emblem = True
2810            If IsNull(garr_varGoToReport(GTR_CTL4)) = False Then
2820              lngTmp01 = 0&: lngTmp02 = 0&
2830              lngTmp03 = (.GoToReport_Emblem_01_img.Width + (8& * lngTpp))
2840              .GoToReport_arw_map_okr_img.Left = (.GoToReport_arw_map_okr_img.Left - lngTmp03)
2850              .cmdOK.Left = (.cmdOK.Left - lngTmp03)
2860              .cmdReinvest.Left = (.cmdReinvest.Left - lngTmp03)
2870              .cmdCancel.Left = (.cmdCancel.Left - lngTmp03)
2880            End If
2890          End If
2900          DoEvents
2910          Select Case blnGoingToReport
              Case True
2920            Select Case IsNull(garr_varGoToReport(GTR_CTL4))
                Case True
2930              DoCmd.Hourglass False
2940              gblnGoToReport = False
2950              blnGoingToReport = False
2960              .GoToReport_arw_pr_img.Visible = False
2970              .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
2980              GTREmblem_Off  ' ** Procedure: Below.
2990              .cmdCancel.SetFocus
3000              .TimerInterval = 0&
3010            Case False
3020              .TimerInterval = 50&
3030              cmdReinvest_Click  ' ** Procedure: Below.
3040            End Select
3050          Case False
3060            DoCmd.Hourglass True  ' ** Make sure it's still running.
3070            DoEvents
3080            blnGoingToReport = True
3090            Select Case IsNull(garr_varGoToReport(GTR_CTL4))
                Case True
3100              .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3110              .GoToReport_arw_pr_img.Visible = True
3120              If .cmdPrintReport.Enabled = True Then
3130                .cmdPrintReport.SetFocus
3140              End If
3150              DoEvents
                  ' ** Make sure this remains off!
3160              GTREmblem_Off  ' ** Procedure: Below.
3170              Beep
3180              DoCmd.Hourglass False
3190              DoEvents
3200              .TimerInterval = (GTR_WAIT * 2&)
3210            Case False
3220              lngBtn_Top = (.cmdReinvest.Top + (lngTpp))
3230              lngBtn_Left = ((.cmdReinvest.Left - .GoToReport_arw_map_okr_img.Width) + (2& * lngTpp))
3240              .GoToReport_arw_map_okr_img.Top = lngBtn_Top
3250              .GoToReport_arw_map_okr_img.Left = lngBtn_Left
3260              .GoToReport_arw_map_okr_img.Visible = True
3270              .cmdReinvest.SetFocus
3280              DoEvents
3290              .TimerInterval = 50&
3300            End Select
3310          End Select
3320        End If
3330      End If
3340      If blnRecalcTots = True Then
3350        blnRecalcTots = False
3360        blnRetVal = RecalcTots  ' ** Function: Below.
3370        If blnRetVal = False Then
3380          blnRecalcTots = True
3390          .TimerInterval = 500&
3400        End If
3410      End If
3420    End With

EXITP:
3430    Exit Sub

ERRH:
3440    DoCmd.Hourglass False
3450    Select Case ERR.Number
        Case Else
3460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3470    End Select
3480    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3510    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Assure all reports are closed.
3520    Do While Reports.Count > 0
3530      DoCmd.Close acReport, Reports(0).Name
3540    Loop

3550    gstrReportCallingForm = vbNullString

EXITP:
3560    Exit Sub

ERRH:
3570    Select Case ERR.Number
        Case Else
3580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3590    End Select
3600    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3710    With Me
3720      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3730        Select Case blnPrintReport_Focus
            Case True
3740          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3750          .cmdPrintReport_raised_img.Visible = False
3760        Case False
3770          .cmdPrintReport_raised_img.Visible = True
3780          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3790        End Select
3800        .cmdPrintReport_raised_focus_img.Visible = False
3810        .cmdPrintReport_raised_focus_dots_img.Visible = False
3820        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3830        .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub cmdOK_Click()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRevCode As Long, lngRevType As Long, lngTaxType As Long
        Dim strRevDesc As String, strTaxDesc As String
        Dim varComments As Variant

4110    DoCmd.Hourglass True
4120    DoEvents

4130    With Me
4140      lngRevCode = .cmbRevenueCodes_Interest
4150      strRevDesc = .cmbRevenueCodes_Interest.Column(1)
4160      lngRevType = .cmbRevenueCodes_Interest.Column(2)
4170      lngTaxcode = .cmbTaxCodes_Interest
4180      strTaxDesc = .cmbTaxCodes_Interest.Column(1)
4190      lngTaxType = .cmbTaxCodes_Interest.Column(2)
4200      varComments = .description
4210    End With

4220    Set dbs = CurrentDb
4230    With dbs

4240      Select Case strCallingForm
          Case "frmJournal"
            ' ** Append Journal Map to Journal, by specified [desc], [revcod], [taxcod],
            ' ** [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id
4250        Set qdf = .QueryDefs("qryMap_Int_03")
4260      Case "frmJournal_Columns"
            ' ** Append Journal Map to tblJournal_Column, by specified
            ' ** [desc], [taxcod], [revcod], [revdesc], [taxdesc], [revtyp], [taxtyp],
            ' ** [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id.
4270        Set qdf = .QueryDefs("qryJournal_Columns_31_06")
4280        With qdf.Parameters
4290          ![revdesc] = strRevDesc
4300          ![taxdesc] = strTaxDesc
4310          ![revtyp] = lngRevType
4320          ![taxtyp] = lngTaxType
4330        End With
4340      End Select
4350      With qdf.Parameters
4360        ![Desc] = Me.description
4370        ![revcod] = lngRevCode
4380        ![taxcod] = lngTaxcode
4390        ![reinvstd] = False
4400      End With
4410      qdf.Execute
4420      Set qdf = Nothing
4430      DoEvents

4440      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
4450        Set qdf = .QueryDefs("qryMap_01")
4460      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
4470        Set qdf = .QueryDefs("qryJournal_Columns_31_04")
4480      End Select
4490      qdf.Execute
4500      Set qdf = Nothing
4510      DoEvents
          ' ** Empty tmpAccount.
4520      Set qdf = .QueryDefs("qryMap_Int_08_01")
4530      qdf.Execute
4540      Set qdf = Nothing
4550      DoEvents

4560      .Close
4570    End With
4580    DoEvents

4590    Select Case strCallingForm
        Case "frmJournal"
4600      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4610        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
4620      Else
4630        Forms(strCallingForm).frmJournal_Sub2_Interest.Form.Requery
4640        DoCmd.SelectObject acForm, strCallingForm, False
4650      End If
4660      DoEvents
4670      Forms(strCallingForm).frmJournal_Sub2_Interest.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub2_Interest.
4680    Case "frmJournal_Columns"
4690      gblnSetFocus = True
4700      gstrReturningForm = "frmMap_Return"
4710      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4720        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
4730      Else
4740        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
4750        Forms(strCallingForm).frmJournal_Columns_Sub.Form.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
4760        DoCmd.SelectObject acForm, strCallingForm, False
4770        Forms(strCallingForm).TimerInterval = 250&
4780      End If
4790    End Select

4800    DoCmd.Close acForm, THIS_NAME

EXITP:
4810    DoCmd.Hourglass False
4820    Set qdf = Nothing
4830    Set dbs = Nothing
4840    Exit Sub

ERRH:
4850    DoCmd.Hourglass False
4860    Select Case ERR.Number
        Case Else
4870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4880    End Select
4890    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
4990          .cmdReinvest.SetFocus
5000        End With
5010      End Select
5020    End If

        ' ** Shift keys.
5030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5040      Select Case intRetVal
          Case vbKeyTab
5050        With Me
5060          intRetVal = 0
5070          .cmdPrintReport.SetFocus
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

Private Sub cmdReinvest_Click()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReinvest_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRevCode As Long, lngRevType As Long, lngTaxType As Long
        Dim strRevDesc As String, strTaxDesc As String
        Dim varComments As Variant
        Dim strDocName As String

5210    DoCmd.Hourglass True
5220    DoEvents

5230    With Me
5240      lngRevCode = .cmbRevenueCodes_Interest
5250      strRevDesc = .cmbRevenueCodes_Interest.Column(1)
5260      lngRevType = .cmbRevenueCodes_Interest.Column(2)
5270      lngTaxcode = .cmbTaxCodes_Interest
5280      strTaxDesc = .cmbTaxCodes_Interest.Column(1)
5290      lngTaxType = .cmbTaxCodes_Interest.Column(2)
5300      varComments = .description
5310    End With

5320    Set dbs = CurrentDb
5330    With dbs

5340      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty tblMark.
5350        Set qdf = .QueryDefs("qryMap_Int_03m")
5360        qdf.Execute
5370        Set qdf = Nothing
5380        DoEvents
            ' ** Append qryMap_Int_03n (Journal, grouped, with Max(ID)) to tblMark.
5390        Set qdf = .QueryDefs("qryMap_Int_03o")
5400        qdf.Execute
5410        Set qdf = Nothing
5420        DoEvents
            ' ** Append Journal Map to Journal, by specified [desc], [revcod], [taxcod],
            ' ** [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id
5430        Set qdf = .QueryDefs("qryMap_Int_03")
5440      Case "frmJournal_Columns"
            ' ** Empty tblMark.
5450        Set qdf = .QueryDefs("qryJournal_Columns_31_06m")
5460        qdf.Execute
5470        Set qdf = Nothing
5480        DoEvents
            ' ** Append qryJournal_Columns_31_06b (tblJournal_Column, grouped, with Max(JrnlCol_ID))to tblMark.
5490        Set qdf = .QueryDefs("qryJournal_Columns_31_06o")
5500        qdf.Execute
5510        Set qdf = Nothing
5520        DoEvents
            ' ** Append Journal Map to tblJournal_Column, by specified
            ' ** [desc], [taxcod], [revcod], [revdesc], [taxdesc], [revtyp], [taxtyp],
            ' ** [reinvstd], GlobalVarGet('gblnGoToReport').  #curr_id
5530        Set qdf = .QueryDefs("qryJournal_Columns_31_06")
5540        With qdf.Parameters
5550          ![revdesc] = strRevDesc
5560          ![taxdesc] = strTaxDesc
5570          ![revtyp] = lngRevType
5580          ![taxtyp] = lngTaxType
5590        End With
5600      End Select
5610      With qdf.Parameters
5620        ![Desc] = varComments
5630        ![revcod] = lngRevCode
5640        ![taxcod] = lngTaxcode
5650        ![reinvstd] = True  ' ** It will only update the appropriate side.
5660      End With
5670      qdf.Execute
5680      Set qdf = Nothing
5690      DoEvents

          ' ** Empty tblJournal_Map_Staging3.
5700      Set qdf = .QueryDefs("qryMap_Reinvest_05_07")
5710      qdf.Execute
5720      Set qdf = Nothing
5730      DoEvents

          ' ** Append Journal Map to tblJournal_Map_Staging3.
5740      Set qdf = .QueryDefs("qryMap_Reinvest_05_04")
5750      qdf.Execute
5760      Set qdf = Nothing
5770      DoEvents

5780      Select Case strCallingForm
          Case "frmJournal"
            ' ** Update qryMap_Reinvest_05_05 (tblJournal_Map_Staging3,
            ' ** linked to Journal, with Journal_ID_new).
5790        Set qdf = .QueryDefs("qryMap_Reinvest_05_06")
5800      Case "frmJournal_Columns"
            ' ** Update qryJournal_Columns_30_14 (tblJournal_Map_Staging3,
            ' ** linked to tblJournal_Column, with JrnlCol_ID_new).
5810        Set qdf = .QueryDefs("qryJournal_Columns_30_15")
5820      End Select
5830      qdf.Execute
5840      Set qdf = Nothing
5850      DoEvents

5860      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
5870        Set qdf = .QueryDefs("qryMap_01")
5880      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
5890        Set qdf = .QueryDefs("qryJournal_Columns_31_04")
5900      End Select
5910      qdf.Execute
5920      Set qdf = Nothing

5930      .Close
5940    End With
5950    Set dbs = Nothing
5960    DoEvents

5970    If strCallingForm = "frmJournal_Columns" Then
5980      Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
5990    End If

6000    strDocName = "frmMap_Reinvest_DivInt_Price"
6010    DoCmd.OpenForm strDocName, , , , , , strCallingForm & "~" & "Interest" & "~" & CStr(lngAssetNo) & "~" & _
          CStr(dblAmount) & "~" & CStr(dblShareface) & "~" & CStr(lngCurrID)
6020    If gblnGoToReport = True Then
6030      Forms(strDocName).TimerInterval = 100&
6040    End If

6050    DoCmd.Close acForm, THIS_NAME

EXITP:
6060    DoCmd.Hourglass False
6070    Set qdf = Nothing
6080    Set dbs = Nothing
6090    Exit Sub

ERRH:
6100    DoCmd.Hourglass False
6110    Select Case ERR.Number
        Case Else
6120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6130    End Select
6140    Resume EXITP

End Sub

Private Sub cmdReinvest_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReinvest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6220    intShiftDown = (Shift And acShiftMask) > 0
6230    intAltDown = (Shift And acAltMask) > 0
6240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6260      Select Case intRetVal
          Case vbKeyTab
6270        With Me
6280          intRetVal = 0
6290          .cmdCancel.SetFocus
6300        End With
6310      End Select
6320    End If

        ' ** Shift keys.
6330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6340      Select Case intRetVal
          Case vbKeyTab
6350        With Me
6360          intRetVal = 0
6370          .cmdOK.SetFocus
6380        End With
6390      End Select
6400    End If

EXITP:
6410    KeyCode = intRetVal
6420    Exit Sub

ERRH:
6430    intRetVal = 0
6440    Select Case ERR.Number
        Case Else
6450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6460    End Select
6470    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intTmp01 As Integer
        Dim intX As Integer

6510    DoCmd.Hourglass True
6520    DoEvents

6530    intTmp01 = 0
6540    Do While Forms.Count > 1
6550      intTmp01 = intTmp01 + 1
6560      For intX = (Forms.Count - 1) To 0 Step -1
6570        Select Case Forms(intX).Name
            Case THIS_NAME
              ' ** Skip.
6580        Case "frmJournal", "frmJournal_Columns"
              ' ** Skip.
6590        Case Else
6600          DoCmd.Close acForm, Forms(intX).Name
6610          Exit For  ' ** Do them one-at-a-time.
6620        End Select
6630      Next
6640      If intTmp01 >= 10 Then Exit Do  ' ** Don't get caught in a loop.
6650    Loop

6660    Set dbs = CurrentDb
6670    With dbs
6680      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
6690        Set qdf = .QueryDefs("qryMap_01")
6700      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
6710        Set qdf = .QueryDefs("qryJournal_Columns_31_04")
6720      End Select
6730      qdf.Execute
6740      Set qdf = Nothing
6750      DoEvents
          ' ** Empty tmpAccount.
6760      Set qdf = .QueryDefs("qryMap_Int_08_01")
6770      qdf.Execute
6780      Set qdf = Nothing
6790      DoEvents
6800      .Close
6810    End With
6820    Set dbs = Nothing
6830    DoEvents

6840    gblnSetFocus = True
6850    Select Case strCallingForm
        Case "frmJournal"
6860      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
6870        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
6880      Else
6890        gstrReturningForm = THIS_NAME
6900        Forms(strCallingForm).frmJournal_Sub2_Interest.Form.Requery
6910        DoCmd.SelectObject acForm, strCallingForm, False
6920        Forms(strCallingForm).TimerInterval = 250&
6930      End If
6940    Case "frmJournal_Columns"
6950      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
6960        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
6970      Else
6980        gstrReturningForm = THIS_NAME
6990        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
7000        DoCmd.SelectObject acForm, strCallingForm, False
7010        Forms(strCallingForm).TimerInterval = 250&
7020      End If
7030    End Select

7040    DoCmd.Close acForm, THIS_NAME

EXITP:
7050    Set qdf = Nothing
7060    Set dbs = Nothing
7070    Exit Sub

ERRH:
7080    DoCmd.Hourglass False
7090    Select Case ERR.Number
        Case Else
7100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7110    End Select
7120    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7220    intShiftDown = (Shift And acShiftMask) > 0
7230    intAltDown = (Shift And acAltMask) > 0
7240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7260      Select Case intRetVal
          Case vbKeyTab
7270        With Me
7280          intRetVal = 0
7290          lngRecsCur = .frmMap_Int_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Int_Detail_Sub.
7300          If lngRecsCur > 0 Then
7310            .frmMap_Int_Detail_Sub.SetFocus
7320            .frmMap_Int_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMap_Int_Detail_Sub.
7330            .frmMap_Int_Detail_Sub.Form.FocusHolder.SetFocus
7340          Else
7350            .description.SetFocus
7360          End If
7370        End With
7380      End Select
7390    End If

        ' ** Shift keys.
7400    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7410      Select Case intRetVal
          Case vbKeyTab
7420        With Me
7430          intRetVal = 0
7440          .cmdReinvest.SetFocus
7450        End With
7460      End Select
7470    End If

EXITP:
7480    KeyCode = intRetVal
7490    Exit Sub

ERRH:
7500    intRetVal = 0
7510    Select Case ERR.Number
        Case Else
7520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7530    End Select
7540    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

7610    With Me
7620      blnPrintReport_Focus = True
7630      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
7640      .cmdPrintReport_raised_img.Visible = False
7650      .cmdPrintReport_raised_focus_img.Visible = False
7660      .cmdPrintReport_raised_focus_dots_img.Visible = False
7670      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7680      .cmdPrintReport_raised_img_dis.Visible = False
7690    End With

EXITP:
7700    Exit Sub

ERRH:
7710    Select Case ERR.Number
        Case Else
7720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7730    End Select
7740    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

7810    With Me
7820      blnPrintReport_MouseDown = True
7830      .cmdPrintReport_sunken_focus_dots_img.Visible = True
7840      .cmdPrintReport_raised_img.Visible = False
7850      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7860      .cmdPrintReport_raised_focus_img.Visible = False
7870      .cmdPrintReport_raised_focus_dots_img.Visible = False
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

Private Sub cmdPrintReport_Click()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

        ' ** qryMapReport_02.
        'strSQL = "SELECT DISTINCTROW account.accountno, account.shortname, Sum([journal map].shareface) AS total_shareface, " & _
        '  "Sum([journal map].icash) AS total_icash, " & CoInfo & ", " & _
        '  "CStr([masterasset].[Description]) & IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy')) AS totdesc " & _
        '  "FROM ([journal map] INNER JOIN account ON [journal map].accountno = account.accountno) " & _
        '  "INNER JOIN masterasset ON [journal map].assetno = masterasset.assetno " & _
        '  "GROUP BY account.accountno, account.shortname, " & _
        '  "CStr([masterasset].[Description]) & IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'));"

8010    With Me
8020      strDocName = "rptMap_Interest"
8030      gstrReportCallingForm = THIS_NAME
8040      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
8050        .Modal = False
8060        DoCmd.OpenReport strDocName, acViewPreview, , , , acViewPreview
8070        DoCmd.Maximize
8080        DoCmd.RunCommand acCmdFitToWindow
8090        .Visible = False
8100      Else
            '##GTR_Ref: rptMap_Interest
8110        DoCmd.OpenReport strDocName, acViewNormal, , , , acViewNormal
8120      End If
8130    End With

EXITP:
8140    Exit Sub

ERRH:
8150    Me.Visible = True
8160    DoCmd.Restore
8170    Select Case ERR.Number
        Case Else
8180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8190    End Select
8200    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8320    intShiftDown = (Shift And acShiftMask) > 0
8330    intAltDown = (Shift And acAltMask) > 0
8340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8360      Select Case intRetVal
          Case vbKeyTab
8370        With Me
8380          intRetVal = 0
8390          .cmdOK.SetFocus
8400        End With
8410      End Select
8420    End If

        ' ** Shift keys.
8430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8440      Select Case intRetVal
          Case vbKeyTab
8450        With Me
8460          intRetVal = 0
8470          If .cmbTaxCodes_Interest.Visible = True And .cmbTaxCodes_Interest.Enabled = True Then
8480            .cmbTaxCodes_Interest.SetFocus
8490          ElseIf .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
8500            .cmbRevenueCodes_Interest.SetFocus
8510          Else
8520            .description.SetFocus
8530          End If
8540        End With
8550      End Select
8560    End If

EXITP:
8570    KeyCode = intRetVal
8580    Exit Sub

ERRH:
8590    intRetVal = 0
8600    Select Case ERR.Number
        Case Else
8610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8620    End Select
8630    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

8710    With Me
8720      If blnPrintReport_MouseDown = False Then
8730        Select Case blnPrintReport_Focus
            Case True
8740          .cmdPrintReport_raised_focus_dots_img.Visible = True
8750          .cmdPrintReport_raised_focus_img.Visible = False
8760        Case False
8770          .cmdPrintReport_raised_focus_img.Visible = True
8780          .cmdPrintReport_raised_focus_dots_img.Visible = False
8790        End Select
8800        .cmdPrintReport_raised_img.Visible = False
8810        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8820        .cmdPrintReport_sunken_focus_dots_img.Visible = False
8830        .cmdPrintReport_raised_img_dis.Visible = False
8840      End If
8850    End With

EXITP:
8860    Exit Sub

ERRH:
8870    Select Case ERR.Number
        Case Else
8880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8890    End Select
8900    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

9010    With Me
9020      .cmdPrintReport_raised_focus_dots_img.Visible = True
9030      .cmdPrintReport_raised_img.Visible = False
9040      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9050      .cmdPrintReport_raised_focus_img.Visible = False
9060      .cmdPrintReport_sunken_focus_dots_img.Visible = False
9070      .cmdPrintReport_raised_img_dis.Visible = False
9080      blnPrintReport_MouseDown = False
9090    End With

EXITP:
9100    Exit Sub

ERRH:
9110    Select Case ERR.Number
        Case Else
9120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9130    End Select
9140    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

9210    With Me
9220      .cmdPrintReport_raised_img.Visible = True
9230      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9240      .cmdPrintReport_raised_focus_img.Visible = False
9250      .cmdPrintReport_raised_focus_dots_img.Visible = False
9260      .cmdPrintReport_sunken_focus_dots_img.Visible = False
9270      .cmdPrintReport_raised_img_dis.Visible = False
9280      blnPrintReport_Focus = False
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

Private Sub total_shareface_sum_KeyDown(KeyCode As Integer, Shift As Integer)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_sum_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
9470        With Me
9480          intRetVal = 0
9490          .description.SetFocus
9500        End With
9510      End Select
9520    End If

        ' ** Shift keys.
9530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9550        With Me
9560          intRetVal = 0
9570          lngRecsCur = .frmMap_Int_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Int_Detail_Sub.
9580          If lngRecsCur > 0& Then
9590            .frmMap_Int_Detail_Sub.SetFocus
9600            .frmMap_Int_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Int_Detail_Sub.
9610            .frmMap_Int_Detail_Sub.Form.FocusHolder.SetFocus
9620          Else
9630            .cmdCancel.SetFocus
9640          End If
9650        End With
9660      End Select
9670    End If

EXITP:
9680    KeyCode = intRetVal
9690    Exit Sub

ERRH:
9700    intRetVal = 0
9710    Select Case ERR.Number
        Case Else
9720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9730    End Select
9740    Resume EXITP

End Sub

Private Sub total_cash_sum_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "total_cash_sum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9820    intShiftDown = (Shift And acShiftMask) > 0
9830    intAltDown = (Shift And acAltMask) > 0
9840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9870        With Me
9880          intRetVal = 0
9890          .description.SetFocus
9900        End With
9910      End Select
9920    End If

        ' ** Shift keys.
9930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9950        With Me
9960          intRetVal = 0
9970          lngRecsCur = .frmMap_Int_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Int_Detail_Sub.
9980          If lngRecsCur > 0& Then
9990            .frmMap_Int_Detail_Sub.SetFocus
10000           .frmMap_Int_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Int_Detail_Sub.
10010           .frmMap_Int_Detail_Sub.Form.FocusHolder.SetFocus
10020         Else
10030           .cmdCancel.SetFocus
10040         End If
10050       End With
10060     End Select
10070   End If

EXITP:
10080   KeyCode = intRetVal
10090   Exit Sub

ERRH:
10100   intRetVal = 0
10110   Select Case ERR.Number
        Case Else
10120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10130   End Select
10140   Resume EXITP

End Sub

Private Sub Overshort_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "Overshort_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10220   intShiftDown = (Shift And acShiftMask) > 0
10230   intAltDown = (Shift And acAltMask) > 0
10240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10270       With Me
10280         intRetVal = 0
10290         .description.SetFocus
10300       End With
10310     End Select
10320   End If

        ' ** Shift keys.
10330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10350       With Me
10360         intRetVal = 0
10370         lngRecsCur = .frmMap_Int_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Int_Detail_Sub.
10380         If lngRecsCur > 0& Then
10390           .frmMap_Int_Detail_Sub.SetFocus
10400           .frmMap_Int_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Int_Detail_Sub.
10410           .frmMap_Int_Detail_Sub.Form.FocusHolder.SetFocus
10420         Else
10430           .cmdCancel.SetFocus
10440         End If
10450       End With
10460     End Select
10470   End If

EXITP:
10480   KeyCode = intRetVal
10490   Exit Sub

ERRH:
10500   intRetVal = 0
10510   Select Case ERR.Number
        Case Else
10520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10530   End Select
10540   Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

10610   With Me
10620     Select Case .curr_date.Visible
          Case True
10630       .curr_date.Visible = False
10640     Case False
10650       .curr_date = .curr_id.Column(CBX_C_DATE)
10660       .curr_date.Visible = True
10670     End Select
10680   End With

EXITP:
10690   Exit Sub

ERRH:
10700   Select Case ERR.Number
        Case Else
10710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
10720   End Select
10730   Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String
        Dim lngX As Long

10810   With Me

10820     If IsNull(.curr_id) = False Then

10830       lngCurrID = .curr_id.Column(CBX_C_CURRID)

10840       blnFound = False
10850       If lngCurrID <> 150& Then
10860         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
10870           NoChar_Load  ' ** Procedure: Below.
10880         End If
10890         For lngX = 0& To (lngNoChars - 1&)
10900           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
10910             blnFound = True
10920           End If
10930         Next
10940       End If

10950       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
10960         .total_cash_sum.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
10970         .total_cash_sum.FontSize = 10
10980         .total_cash_sum.FontBold = True
10990         .overshort.FontName = "Courier New"
11000         .overshort.FontSize = 10
11010         .overshort.FontBold = True
11020         .pershare.FontName = "Courier New"
11030         .pershare.FontSize = 10
11040         .pershare.FontBold = True
11050       Case False
11060         If .total_cash_sum.FontName <> "Arial" Then
11070           .total_cash_sum.FontName = "Arial"
11080           .total_cash_sum.FontSize = 10
11090           .total_cash_sum.FontBold = False
11100           .overshort.FontName = "Arial"
11110           .overshort.FontSize = 10
11120           .overshort.FontBold = False
11130           .pershare.FontName = "Arial"
11140           .pershare.FontSize = 10
11150           .pershare.FontBold = False
11160         End If
11170       End Select

11180       If lngCurrID = 150& Then  ' ** USD.
11190         strFrmt1 = "Currency"
11200         intDec1 = 2
11210         strFrmt2 = "Standard"
11220         intDec2 = 5
11230         lngBkClr = CLR_DISABLED_BG
11240       Else
11250         strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
11260         If .curr_id.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
11270           intDec1 = 2
11280         Else
11290           intDec1 = .curr_id.Column(CBX_C_DEC)  ' ** Decimal places.
11300         End If
11310         intDec2 = 5
11320         intLen = Len(strTmp01)
11330         Select Case intLen
              Case 1
                'strTmp01 = strTmp01
11340         Case 2
11350           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
11360         Case 3
11370           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
11380         Case 4
11390           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
11400         Case 5
11410           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                  "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
11420         End Select
11430         Select Case intDec1
              Case 0
11440           strFrmt1 = strTmp01 & "#,##0;(" & strTmp01 & "#,##0);" & strTmp01 & "0;" & strTmp01 & "0"
11450           strFrmt2 = "#,##0.00000;(" & "#,##0.00000);" & "0.00000;" & "0.00000"
11460         Case 1  ' ** None have this currently.
11470           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
11480           strFrmt2 = "#,##0.00000;(" & "#,##0.00000);" & "0.00000;" & "0.00000"
11490         Case 2
11500           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
11510           strFrmt2 = "#,##0.00000;(" & "#,##0.00000);" & "0.00000;" & "0.00000"
11520         Case 3
11530           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
11540           strFrmt2 = "#,##0.00000;(" & "#,##0.00000);" & "0.00000;" & "0.00000"
11550         End Select
11560         lngBkClr = CLR_LTGRN
11570       End If
11580       .total_cash_sum.Format = strFrmt1
11590       .total_cash_sum.DecimalPlaces = intDec1
11600       .total_cash_sum.BackColor = lngBkClr
11610       .overshort.Format = strFrmt1
11620       .overshort.DecimalPlaces = intDec1
11630       .overshort.BackColor = lngBkClr
11640       .pershare.Format = strFrmt2
11650       .pershare.DecimalPlaces = intDec2
11660       .pershare.BackColor = lngBkClr

11670       If lngCurrID = 150& Then
11680         .total_cash_sum_usd.Visible = False
11690         .Overshort_usd.Visible = False
11700         .pershare_usd.Visible = False
11710       Else
11720         .total_cash_sum_usd = 0
11730         .total_cash_sum_usd.Visible = True
11740         .Overshort_usd = 0
11750         .Overshort_usd.Visible = True
11760         .pershare_usd = 0
11770         .pershare_usd.Visible = False
11780       End If

11790     End If

11800   End With

EXITP:
11810   Exit Sub

ERRH:
11820   Select Case ERR.Number
        Case Else
11830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11840   End Select
11850   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

11910   With Me
11920     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
11930     .curr_id.SetFocus
11940     Response = acDataErrContinue
11950   End With

EXITP:
11960   Exit Sub

ERRH:
11970   Select Case ERR.Number
        Case Else
11980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11990   End Select
12000   Resume EXITP

End Sub

Private Sub description_GotFocus()

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

12110   With Me
12120     .description.SelLength = 0
12130     .description.SelStart = 999
12140   End With

EXITP:
12150   Exit Sub

ERRH:
12160   Select Case ERR.Number
        Case Else
12170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12180   End Select
12190   Resume EXITP

End Sub

Private Sub description_KeyDown(KeyCode As Integer, Shift As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12220   intShiftDown = (Shift And acShiftMask) > 0
12230   intAltDown = (Shift And acAltMask) > 0
12240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12270       With Me
12280         intRetVal = 0
12290         If .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
12300           .cmbRevenueCodes_Interest.SetFocus
12310         ElseIf .cmbTaxCodes_Interest.Visible = True And .cmbTaxCodes_Interest.Enabled = True Then
12320           .cmbTaxCodes_Interest.SetFocus
12330         Else
12340           .cmdReinvest.SetFocus   ' ** Yes, go to here.
12350         End If
12360       End With
12370     End Select
12380   End If

        ' ** Shift keys.
12390   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12400     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12410       With Me
12420         intRetVal = 0
12430         lngRecsCur = .frmMap_Int_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Int_Detail_Sub.
12440         If lngRecsCur > 0& Then
12450           .frmMap_Int_Detail_Sub.SetFocus
12460           .frmMap_Int_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Int_Detail_Sub.
12470           .frmMap_Int_Detail_Sub.Form.FocusHolder.SetFocus
12480         Else
12490           .cmdCancel.SetFocus
12500         End If
12510       End With
12520     End Select
12530   End If

EXITP:
12540   KeyCode = intRetVal
12550   Exit Sub

ERRH:
12560   intRetVal = 0
12570   Select Case ERR.Number
        Case Else
12580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12590   End Select
12600   Resume EXITP

End Sub

Private Sub description_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "description_MouseMove"

12710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12720   Exit Sub

ERRH:
12730   Select Case ERR.Number
        Case Else
12740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12750   End Select
12760   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_display_GotFocus()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_display_GotFocus"

12810   With Me
12820     .cmbRevenueCodes_Interest.SetFocus
12830   End With

EXITP:
12840   Exit Sub

ERRH:
12850   Select Case ERR.Number
        Case Else
12860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12870   End Select
12880   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_display_KeyDown(KeyCode As Integer, Shift As Integer)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_display_KeyDown"

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
12990         .cmbRevenueCodes_Interest.SetFocus
13000       End With
13010     End Select
13020   End If

        ' ** Shift keys.
13030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13050       With Me
13060         intRetVal = 0
13070         .description.SetFocus
13080       End With
13090     End Select
13100   End If

EXITP:
13110   KeyCode = intRetVal
13120   Exit Sub

ERRH:
13130   intRetVal = 0
13140   Select Case ERR.Number
        Case Else
13150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13160   End Select
13170   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_Enter()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_Enter"

13210   With Me
13220     Select Case IsNull(.cmbRevenueCodes_Interest)
          Case True
13230       ExpandCombo .cmbRevenueCodes_Interest  ' ** Module Procedure: modUtilities.
13240     Case False
13250       If .cmbRevenueCodes_Interest = REVID_INC Or .cmbRevenueCodes_Interest = REVID_EXP Then
13260         ExpandCombo .cmbRevenueCodes_Interest  ' ** Module Procedure: modUtilities.
13270       End If
13280     End Select
13290   End With

EXITP:
13300   Exit Sub

ERRH:
13310   Select Case ERR.Number
        Case Else
13320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13330   End Select
13340   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_AfterUpdate()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_AfterUpdate"

13410   With Me
13420     Select Case IsNull(.cmbRevenueCodes_Interest)
          Case True
13430       .cmbRevenueCodes_Interest = REVID_INC
13440     Case False
            ' ** Nothing at the moment.
13450     End Select
13460     If .cmbRevenueCodes_Interest = REVID_INC Or .cmbRevenueCodes_Interest = REVID_EXP Then
13470       .cmbRevenueCodes_Interest_display = Null
13480     Else
13490       .cmbRevenueCodes_Interest_display = .cmbRevenueCodes_Interest.Column(1)
13500     End If
13510   End With

EXITP:
13520   Exit Sub

ERRH:
13530   Select Case ERR.Number
        Case Else
13540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13550   End Select
13560   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Interest_KeyDown(KeyCode As Integer, Shift As Integer)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Interest_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13620   intShiftDown = (Shift And acShiftMask) > 0
13630   intAltDown = (Shift And acAltMask) > 0
13640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13670       With Me
13680         intRetVal = 0
13690         If .cmbTaxCodes_Interest.Visible = True And .cmbTaxCodes_Interest.Enabled = True Then
13700           .cmbTaxCodes_Interest.SetFocus
13710         Else
13720           .cmdReinvest.SetFocus
13730         End If
13740       End With
13750     End Select
13760   End If

        ' ** Shift keys.
13770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13790       With Me
13800         intRetVal = 0
13810         .description.SetFocus
13820       End With
13830     End Select
13840   End If

EXITP:
13850   KeyCode = intRetVal
13860   Exit Sub

ERRH:
13870   intRetVal = 0
13880   Select Case ERR.Number
        Case Else
13890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13900   End Select
13910   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_display_GotFocus()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_display_GotFocus"

14010   With Me
14020     .cmbTaxCodes_Interest.SetFocus
14030   End With

EXITP:
14040   Exit Sub

ERRH:
14050   Select Case ERR.Number
        Case Else
14060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14070   End Select
14080   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_display_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_display_KeyDown"

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
14190         .cmbTaxCodes_Interest.SetFocus
14200       End With
14210     End Select
14220   End If

        ' ** Shift keys.
14230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14250       With Me
14260         intRetVal = 0
14270         If .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
14280           .cmbRevenueCodes_Interest.SetFocus
14290         Else
14300           .description.SetFocus
14310         End If
14320       End With
14330     End Select
14340   End If

EXITP:
14350   KeyCode = intRetVal
14360   Exit Sub

ERRH:
14370   intRetVal = 0
14380   Select Case ERR.Number
        Case Else
14390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14400   End Select
14410   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_Enter()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_Enter"

14510   With Me
14520     Select Case IsNull(.cmbTaxCodes_Interest)
          Case True
14530       ExpandCombo .cmbTaxCodes_Interest  ' ** Module Procedure: modUtilities.
14540     Case False
14550       If .cmbTaxCodes_Interest = TAXID_INC Or .cmbTaxCodes_Interest = TAXID_DED Then
14560         ExpandCombo .cmbTaxCodes_Interest  ' ** Module Procedure: modUtilities.
14570       End If
14580     End Select
14590   End With

EXITP:
14600   Exit Sub

ERRH:
14610   Select Case ERR.Number
        Case Else
14620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14630   End Select
14640   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_AfterUpdate()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_AfterUpdate"

14710   With Me
14720     Select Case IsNull(.cmbTaxCodes_Interest)
          Case True
14730       .cmbTaxCodes_Interest = TAXID_INC
14740     Case False
            ' ** Nothing at the moment.
14750     End Select
14760     If .cmbTaxCodes_Interest = TAXID_INC Or .cmbTaxCodes_Interest = TAXID_DED Then
14770       .cmbTaxCodes_Interest_display = Null
14780     Else
14790       .cmbTaxCodes_Interest_display = .cmbTaxCodes_Interest.Column(1)
14800     End If
14810   End With

EXITP:
14820   Exit Sub

ERRH:
14830   Select Case ERR.Number
        Case Else
14840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14850   End Select
14860   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Interest_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Interest_KeyDown"

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
14990         .cmdReinvest.SetFocus
15000       End With
15010     End Select
15020   End If

        ' ** Shift keys.
15030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15050       With Me
15060         intRetVal = 0
15070         If .cmbRevenueCodes_Interest.Visible = True And .cmbRevenueCodes_Interest.Enabled = True Then
15080           .cmbRevenueCodes_Interest.SetFocus
15090         Else
15100           .description.SetFocus
15110         End If
15120       End With
15130     End Select
15140   End If

EXITP:
15150   KeyCode = intRetVal
15160   Exit Sub

ERRH:
15170   intRetVal = 0
15180   Select Case ERR.Number
        Case Else
15190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15200   End Select
15210   Resume EXITP

End Sub

Public Function RecalcTots()
' ** I give up trying to get it from the subform!

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim dblTmp01 As Double, dblTmp02 As Double
        Dim lngX As Long
        Dim blnRetVal As Boolean

15310   With Me
15320     blnRetVal = False
15330     Set rst = .frmMap_Int_Detail_Sub.Form.RecordsetClone
15340     With rst
15350       If .BOF = True And .EOF = True Then
              ' ** Ho-Hum...
15360       Else
15370         .MoveLast
15380         lngRecs = .RecordCount
15390         .MoveFirst
15400         dblTmp01 = 0#: dblTmp02 = 0#
15410         For lngX = 1& To lngRecs
15420           dblTmp01 = dblTmp01 + ![total_shareface]
15430           dblTmp02 = dblTmp02 + ![total_icash]
15440           If lngX < lngRecs Then .MoveNext
15450         Next
15460       End If
15470       .Close
15480     End With
15490     If dblTmp01 <> 0# Then
15500       blnRetVal = True
15510     End If
15520     .total_shareface_sum = dblTmp01
15530     .total_cash_sum = dblTmp02
15540   End With

EXITP:
15550   Set rst = Nothing
15560   RecalcTots = blnRetVal
15570   Exit Function

ERRH:
15580   blnRetVal = False
15590   Select Case ERR.Number
        Case Else
15600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15610   End Select
15620   Resume EXITP

End Function

Public Sub IncludeCurrency(blnCalc As Boolean)

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim dblRate2 As Double
        Dim blnFound As Boolean
        Dim dblTmp01 As Double

15710   With Me
15720     Select Case blnCalc
          Case True
15730       If blnCurrID = True Then
15740         blnFound = False
15750         Do While blnFound = False
15760           If IsNull(.total_cash_sum) = False Then
15770             If .total_cash_sum <> 0 Then
15780               blnFound = True
15790               dblRate2 = .curr_id.Column(CBX_C_RATE2)
15800               dblTmp01 = (.total_cash_sum * dblRate2)
15810               dblTmp01 = Round(dblTmp01, 2)
15820               .total_cash_sum_usd = dblTmp01
15830               dblTmp01 = (.overshort * dblRate2)
15840               dblTmp01 = Round(dblTmp01, 2)
15850               .Overshort_usd = dblTmp01
15860               If lngCurrID <> 150& Then
15870                 .total_cash_sum_usd.Visible = True
15880                 .Overshort_usd.Visible = True
15890               End If
15900             End If
15910           End If
15920         Loop
15930       End If
15940     Case False
15950       .curr_id.Enabled = False
15960       .curr_id.Locked = True
15970       .curr_id.ForeColor = CLR_DISABLED_FG
15980       .curr_id.BackColor = CLR_VLTGRN
15990       If blnCurrID = True And lngCurrID <> 150& Then
16000         .curr_id.Visible = True
16010         .curr_id_cmd.Visible = True
16020         .curr_id_cmd.Enabled = True
16030       Else
16040         .curr_id.Visible = False
16050         .curr_id_cmd.Visible = False
16060         .curr_id_cmd.Enabled = False
16070         .cmbRevenueCodes_Interest_vline05.Visible = False
16080         .cmbRevenueCodes_Interest_vline06.Visible = False
16090         .cmbRevenueCodes_Interest_box.Width = (.cmbRevenueCodes_Interest_vline06.Left - .cmbRevenueCodes_Interest_box.Left)
16100         .overshort.Top = .Overshort_alt_line.Top
16110         .Overshort_lbl.Top = (.Overshort_alt_line.Top + lngTpp)
16120         .pershare.Top = .pershare_alt_line.Top
16130         .pershare_lbl.Top = (.pershare_alt_line.Top + lngTpp)
16140         .pershare_lbl2.Top = (.pershare_alt_line.Top + lngTpp)
16150       End If
16160       .total_cash_sum_usd.Visible = False
16170       .Overshort_usd.Visible = False
16180       .pershare_usd.Visible = False
16190       .pershare_usd_lbl2.Visible = False
16200     End Select
16210   End With

EXITP:
16220   Exit Sub

ERRH:
16230   DoCmd.Hourglass False
16240   Select Case ERR.Number
        Case Else
16250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16260   End Select
16270   Resume EXITP

End Sub

Private Sub IncludeCurrency_Sub()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim lngX As Long

16310   With Me
16320     If blnCurrID = True And lngCurrID <> 150& Then

16330       With .frmMap_Int_Detail_Sub.Form
16340         lngTmp01 = .Width - .total_icash_alt_line.Left
16350       End With

16360       lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
16370       lngMonitorNum = 1&: lngTmp02 = 0&
16380       EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
16390       If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

16400       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
16410       DoCmd.MoveSize (lngFrm_Left - (lngTmp01 / 2)), lngTmp02, (lngFrm_Width + lngTmp01), lngFrm_Height  'lngFrm_Top
16420       If lngMonitorNum > 1& Then
16430         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
16440       End If

16450       .Width = (.Width + lngTmp01)
16460       .frmMap_Int_Detail_Sub.Width = (.frmMap_Int_Detail_Sub.Width + lngTmp01)
16470       .frmMap_Int_Detail_Sub_box.Width = (.frmMap_Int_Detail_Sub_box.Width + lngTmp01)

16480       .Header_vline01.Left = .Width
16490       .Header_vline02.Left = .Width
16500       .Footer_vline01.Left = .Width
16510       .Footer_vline02.Left = .Width
16520       .Header_hline01.Width = .Width
16530       .Header_hline02.Width = .Width
16540       .Footer_hline01.Width = .Width
16550       .Footer_hline02.Width = .Width

16560       .total_cash_sum_usd.Left = ((.total_cash_sum.Left + .total_cash_sum.Width) + (4& * lngTpp))
16570       .total_cash_sum_usd.Top = .total_cash_sum.Top
16580       .Overshort_usd.Left = ((.overshort.Left + .overshort.Width) + (4& * lngTpp))
16590       .overshort.Top = .Overshort_alt_line.Top
16600       .Overshort_lbl.Top = (.Overshort_alt_line.Top + lngTpp)
16610       .Overshort_usd.Top = .Overshort_alt_line.Top
16620       .pershare_usd.Left = ((.total_cash_sum_usd.Left + .total_cash_sum_usd.Width) - .pershare_usd.Width)
16630       .pershare.Top = .pershare_alt_line.Top
16640       .pershare.RightMargin = 0&
16650       .pershare_lbl.Top = (.pershare_alt_line.Top + lngTpp)
16660       .pershare_lbl2.Top = (.pershare_alt_line.Top + lngTpp)
16670       .pershare_usd.Top = .pershare_alt_line.Top

16680       .total_cash_sum_usd.ForeColor = CLR_DISABLED_FG
16690       .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
16700       .Overshort_usd.ForeColor = CLR_DISABLED_FG
16710       .Overshort_usd.BackColor = CLR_DISABLED_BG
16720       .pershare_usd.ForeColor = CLR_DISABLED_FG
16730       .pershare_usd.BackColor = CLR_DISABLED_BG

16740       .maptype.Left = (.maptype.Left + lngTmp01)
16750       .cmdCancel.Left = (.cmdCancel.Left + lngTmp01)
16760       .cmdReinvest.Left = (.cmdReinvest.Left + lngTmp01)
16770       .cmdOK.Left = (.cmdOK.Left + lngTmp01)

16780       lngTmp01 = (.GoToReport_Emblem_01_img.Left + lngTmp01)
16790       For lngX = 1& To 24&
16800         .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Left = lngTmp01
16810       Next

16820     End If
16830   End With

EXITP:
16840   Exit Sub

ERRH:
16850   DoCmd.Hourglass False
16860   Select Case ERR.Number
        Case Else
16870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16880   End Select
16890   Resume EXITP

End Sub

Private Sub NoChar_Load()

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

16910   Set dbs = CurrentDb
16920   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
16930     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
16940     Set rst = qdf.OpenRecordset
16950     With rst
16960       .MoveLast
16970       lngNoChars = .RecordCount
16980       .MoveFirst
16990       arr_varNoChar = .GetRows(lngNoChars)
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
17000       .Close
17010     End With
17020     .Close
17030   End With

EXITP:
17040   Set rst = Nothing
17050   Set qdf = Nothing
17060   Set dbs = Nothing
17070   Exit Sub

ERRH:
17080   DoCmd.Hourglass False
17090   Select Case ERR.Number
        Case Else
17100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17110   End Select
17120   Resume EXITP

End Sub

Public Sub GTREmblem_Off()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "GTREmblem_Off"

        Dim lngX As Long

17210   With Me
17220     blnGTR_Emblem = False
17230     For lngX = 1& To 24&
17240       .Controls("GoToReport_Emblem_" & Right("00" & CStr(lngX), 2) & "_img").Visible = False
17250     Next
17260   End With

EXITP:
17270   Exit Sub

ERRH:
17280   Select Case ERR.Number
        Case Else
17290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17300   End Select
17310   Resume EXITP

End Sub
