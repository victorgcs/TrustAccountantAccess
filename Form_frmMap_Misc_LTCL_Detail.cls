VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Misc_LTCL_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Misc_LTCL_Detail"

'VGC 08/30/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   IncludeCurrency_Sub()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Misc}
' **   OK:               O {cmdOK}
' **   Print Report:     P {cmdPrintReport}
' **   Short Name:       S {shortname on frmMap_Misc_LTCL_Detail_Sub}
' **   Tax Codes:        T {cmbTaxCodes_Misc}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots}

' ** MapType = "Misc."

' ** Report:
' **   rptMap_Misc_LTCL
' **     qryMapReport_07

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long
Private CLR_DISABLED_BG_R As Long
Private CLR_DISABLED_BG_VLR As Long

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
Private blnGoingToReport As Boolean
Private blnCurrID As Boolean, lngCurrID As Long
Private blnRecalcTots As Boolean
Private strAccountNo_OU As String, dblAmount As Double, dblPerShare As Double
Private dblOverUnder As Double, lngTaxcode As Long, lngAssetNo As Long, datTransDate As Date
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String, strCallersCallingForm As String
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

170             strCallingForm = "frmMap_Misc_LTCL"  ' ** The only form that does.

                ' ** strCallersCallingForm & "~" & cstr(lngAssetNo) & "~" & CStr(dblOverUnder) & "~" & CStr(dblPerShare) & "~" & _
                ' **   Format(datTransDate, "mm/dd/yyyy")
180             intPos01 = InStr(strTmp01, "~")
190             strCallersCallingForm = Left(strTmp01, (intPos01 - 1))
200             strTmp01 = Mid(strTmp01, (intPos01 + 1))
210             intPos01 = InStr(strTmp01, "~")
220             lngAssetNo = Val(Left(strTmp01, (intPos01 - 1)))
230             strTmp01 = Mid(strTmp01, (intPos01 + 1))
240             intPos01 = InStr(strTmp01, "~")
250             dblOverUnder = Val(Left(strTmp01, (intPos01 - 1)))
260             strTmp01 = Mid(strTmp01, (intPos01 + 1))
270             intPos01 = InStr(strTmp01, "~")
280             dblPerShare = Val(Left(strTmp01, (intPos01 - 1)))
290             strTmp01 = Mid(strTmp01, (intPos01 + 1))
300             intPos01 = InStr(strTmp01, "~")
310             datTransDate = CDate(Left(strTmp01, (intPos01 - 1)))
320             strTmp01 = Mid(strTmp01, (intPos01 + 1))
330             intPos01 = InStr(strTmp01, "~")
340             dblAmount = Val(Mid(strTmp01, (intPos01 - 1)))
350             lngCurrID = Val(Mid(strTmp01, (intPos01 + 1)))
360             lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

370             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
380             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

                ' ** Variables are fed empty, then populated ByRef.
390             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

400             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

410             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
420             If lngMonitorCnt > 1& Then
430               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
440               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
450               If lngMonitorNum > 1& Then
460                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
470               End If
480             End If

490             CLR_DISABLED_FG = CLR_DKGRY
500             CLR_DISABLED_BG = CLR_LTTEAL
510             CLR_DISABLED_BG_R = CLR_LTRED
520             CLR_DISABLED_BG_VLR = CLR_VLTRED

530             blnCurrID = Pref_CurrID  ' ** Module Function: modPreferenceFuncs.
540             .curr_id = lngCurrID
550             IncludeCurrency_Sub  ' ** Procedure: Below.

560             Select Case gblnAccountNoWithType
                Case True
570               strAccountNo_OU = "99-INCOME O/U"
580             Case False
590               strAccountNo_OU = "INCOME O/U"
600             End Select

610             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

620             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

630             Set dbs = CurrentDb
640             With dbs
650               Select Case strCallersCallingForm
                  Case "frmJournal"
                    ' ** tblJournal_Map_Staging2, linked to Account.
660                 Set qdf = .QueryDefs("qryMap_Misc_LTCL_Detail_Sub_01_02")
670               Case "frmJournal_Columns"
                    ' ** tblJournal_Map_Staging2, linked to Account.
680                 Set qdf = .QueryDefs("qryMap_Misc_LTCL_Detail_Sub_51_02")
690               End Select
700               Set rst = qdf.OpenRecordset
710               With rst
720                 If .BOF = True And .EOF = True Then
730                   lngRecs = 0&
740                 Else
750                   .MoveLast
760                   lngRecs = .RecordCount
770                 End If
780                 .Close
790               End With
800               .Close
810             End With

820             Select Case strCallersCallingForm
                Case "frmJournal"
830               If .frmMap_Misc_LTCL_Detail_Sub.Form.RecordSource <> "qryMap_Misc_LTCL_Detail_Sub_01_02" Then
840                 .frmMap_Misc_LTCL_Detail_Sub.Form.RecordSource = "qryMap_Misc_LTCL_Detail_Sub_01_02"
850               End If
860             Case "frmJournal_Columns"
870               If .frmMap_Misc_LTCL_Detail_Sub.Form.RecordSource <> "qryMap_Misc_LTCL_Detail_Sub_51_02" Then
880                 .frmMap_Misc_LTCL_Detail_Sub.Form.RecordSource = "qryMap_Misc_LTCL_Detail_Sub_51_02"
890               End If
900             End Select

910             .maptype = "Misc."
920             .txtNumberRecords = "Number Records = " & CStr(lngRecs)
930             .overshort = dblOverUnder

940             .cmbRecurringItems = "Long Term Capital Loss"
950             .description = "LTCL"
960             .pershare = dblPerShare  '5 dec.
970             .amount = dblAmount

980             If .cmdOK.Enabled = True Then
990               .cmdOK.SetFocus
1000            Else
1010              .description.SetFocus
1020            End If

1030            blnRecalcTots = True
1040            .TimerInterval = 500&

1050            DoEvents
1060            DoCmd.SelectObject acForm, THIS_NAME, False
1070            DoEvents

1080            blnIsOpen = False

1090          Else
1100            Cancel = -1
1110          End If
1120        Else
1130          Cancel = -1
1140        End If
1150      Else
1160        Cancel = -1
1170      End If
1180    End With

1190    If Cancel = -1 Then
1200      Beep
1210      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1220    End If

EXITP:
1230    Set rst = Nothing
1240    Set qdf = Nothing
1250    Set dbs = Nothing
1260    Exit Sub

ERRH:
1270    DoCmd.Hourglass False
1280    Select Case ERR.Number
        Case Else
1290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1300    End Select
1310    Resume EXITP

End Sub

Private Sub Form_Load()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1410    With Me

1420      .total_shareface_sum.ForeColor = CLR_DISABLED_FG
1430      .total_shareface_sum.BackColor = CLR_DISABLED_BG
1440      .total_cash_sum.ForeColor = CLR_DISABLED_FG
1450      .total_cash_sum.BackColor = CLR_DISABLED_BG_VLR
1460      .overshort.ForeColor = CLR_DISABLED_FG
1470      .overshort.BackColor = CLR_DISABLED_BG
1480      .pershare.ForeColor = CLR_DISABLED_FG
1490      .pershare.BackColor = CLR_DISABLED_BG_VLR

          ' ** Borrowing these variables from the Court Reports.
1500      If gstrCrtRpt_Version <> vbNullString Then
1510        .cusip = gstrCrtRpt_Ordinal
1520        .cusip.ForeColor = CLR_DISABLED_FG
1530        .cusip.BackColor = CLR_DISABLED_BG
1540        .cusip.Visible = True
1550        .totdesc = gstrCrtRpt_Version
1560        .totdesc.ForeColor = CLR_DISABLED_FG
1570        .totdesc.BackColor = CLR_DISABLED_BG
1580        .totdesc.Visible = True
1590      End If

1600      .transdate = datTransDate

1610      Select Case IsNull(.cmbRevenueCodes_Misc)
          Case True
1620        .cmbRevenueCodes_Misc = REVID_INC
1630        .cmbRevenueCodes_Misc_display = Null
1640      Case False
1650        If .cmbRevenueCodes_Misc = REVID_INC Or .cmbRevenueCodes_Misc = REVID_EXP Then
1660          .cmbRevenueCodes_Misc_display = Null
1670        Else
1680          .cmbRevenueCodes_Misc_display = .cmbRevenueCodes_Misc.Column(1)
1690        End If
1700      End Select

1710      Select Case IsNull(.cmbTaxCodes_Misc)
          Case True
1720        .cmbTaxCodes_Misc = TAXID_INC
1730        .cmbTaxCodes_Misc_display = Null
1740      Case False
1750        If .cmbTaxCodes_Misc = TAXID_INC Or .cmbTaxCodes_Misc = TAXID_DED Then
1760          .cmbTaxCodes_Misc_display = Null
1770        Else
1780          .cmbTaxCodes_Misc_display = .cmbTaxCodes_Misc.Column(1)
1790        End If
1800      End Select

1810      Select Case gblnRevenueExpenseTracking
          Case True
1820        .cmbRevenueCodes_Misc.Visible = True
1830        .cmbRevenueCodes_Misc_display.Visible = True
1840        .cmbRevenueCodes_Misc_lbl2.Visible = False  ' ** Option off msg.
1850      Case False
1860        .cmbRevenueCodes_Misc.Visible = False
1870        .cmbRevenueCodes_Misc_display.Visible = False
1880        .cmbRevenueCodes_Misc_lbl_box.Visible = False
1890        .cmbRevenueCodes_Misc_lbl2.Visible = True  ' ** Option off msg.
1900      End Select

1910      Select Case gblnIncomeTaxCoding
          Case True
1920        .cmbTaxCodes_Misc.Visible = True
1930        .cmbTaxCodes_Misc_display.Visible = True
1940        .cmbTaxCodes_Misc_lbl2.Visible = False  ' ** Option off msg.
1950      Case False
1960        .cmbTaxCodes_Misc.Visible = False
1970        .cmbTaxCodes_Misc_display.Visible = False
1980        .cmbTaxCodes_Misc_lbl_box.Visible = False  ' ** Option off msg.
1990        .cmbTaxCodes_Misc_lbl2.Visible = True
2000      End Select

2010      If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
2020        .cmbRevenueCodes_Misc_box_lbl.ForeColor = WIN_CLR_DISF
2030        .cmbRevenueCodes_Misc_box_lbl_dim_hi.Visible = True
2040      End If

2050      curr_id_AfterUpdate  ' ** Procedure: Below.
2060      IncludeCurrency False  ' ** Procedure: Below
2070      DoEvents

2080      DoCmd.Hourglass False
2090      .frmMap_Misc_LTCL_Detail_Sub.Form.TimerInterval = 500&

2100    End With

EXITP:
2110    Exit Sub

ERRH:
2120    DoCmd.Hourglass False
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Short Name:       S {shortname on frmMap_Misc_LTCL_Detail_Sub}

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

        ' ** Alt keys.
2340    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2350      Select Case intRetVal
          Case vbKeyS
2360        With Me
2370          intRetVal = 0
2380          .frmMap_Misc_LTCL_Detail_Sub.SetFocus
2390          .frmMap_Misc_LTCL_Detail_Sub.Form.shortname.SetFocus
2400        End With
2410      End Select
2420    End If

        ' ** Ctrl keys.
2430    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2440      Select Case intRetVal
          Case vbKeyM
2450        With Me
2460          intRetVal = 0
2470          .description.SetFocus
2480        End With
2490      End Select
2500    End If

        ' ** Ctrl-Shift keys.
2510    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2520      Select Case intRetVal
          Case vbKeyF
2530        With Me
2540          intRetVal = 0
2550          .FocusHolder.SetFocus
2560        End With
2570      End Select
2580    End If

EXITP:
2590    KeyCode = intRetVal
2600    Exit Sub

ERRH:
2610    intRetVal = 0
2620    Select Case ERR.Number
        Case Else
2630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2640    End Select
2650    Resume EXITP

End Sub

Private Sub Form_Timer()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim blnRetVal As Boolean

2710    With Me
2720      .TimerInterval = 0&
2730      If gblnGoToReport = True Then
2740        Select Case blnGoingToReport
            Case True
2750          DoCmd.Hourglass False
2760          gblnGoToReport = False
2770          blnGoingToReport = False
2780          .GoToReport_arw_pr_img.Visible = False
2790          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
2800          .cmdCancel.SetFocus
2810          DoEvents
2820        Case False
2830          DoCmd.Hourglass True  ' ** Make sure it's still running.
2840          DoEvents
2850          blnGoingToReport = True
2860          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
2870          .GoToReport_arw_pr_img.Visible = True
2880          If .cmdPrintReport.Enabled = True Then
2890            .cmdPrintReport.SetFocus
2900          End If
2910          DoEvents
2920          .cmdOK.Enabled = False
2930          Beep
2940          DoCmd.Hourglass False
2950          DoEvents
2960          .TimerInterval = (GTR_WAIT * 2&)
2970        End Select
2980      End If
2990      If blnRecalcTots = True Then
3000        blnRecalcTots = False
3010        blnRetVal = RecalcTots  ' ** Function: Below.
3020        If blnRetVal = False Then
3030          blnRecalcTots = True
3040          .TimerInterval = 500&
3050        End If
3060      End If
3070    End With

EXITP:
3080    Exit Sub

ERRH:
3090    DoCmd.Hourglass False
3100    Select Case ERR.Number
        Case Else
3110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3120    End Select
3130    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3210    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3220    Set dbs = CurrentDb
        ' ** Empty tmpAccount.
3230    Set qdf = dbs.QueryDefs("qryMap_Misc_LTCL_08_01")
3240    qdf.Execute
3250    Set qdf = Nothing
3260    DoEvents
3270    dbs.Close
3280    Set dbs = Nothing
3290    DoEvents

        ' ** Assure all reports are closed.
3300    If Reports.Count > 0 Then
3310      Do While Reports.Count > 0
3320        DoCmd.Close acReport, Reports(0).Name
3330      Loop
3340    End If

3350    gstrReportCallingForm = vbNullString

EXITP:
3360    Set qdf = Nothing
3370    Set dbs = Nothing
3380    Exit Sub

ERRH:
3390    DoCmd.Hourglass False
3400    Select Case ERR.Number
        Case Else
3410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3420    End Select
3430    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3510    With Me
3520      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3530        Select Case blnPrintReport_Focus
            Case True
3540          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3550          .cmdPrintReport_raised_img.Visible = False
3560        Case False
3570          .cmdPrintReport_raised_img.Visible = True
3580          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3590        End Select
3600        .cmdPrintReport_raised_focus_img.Visible = False
3610        .cmdPrintReport_raised_focus_dots_img.Visible = False
3620        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3630        .cmdPrintReport_raised_img_dis.Visible = False
3640      End If
3650    End With

EXITP:
3660    Exit Sub

ERRH:
3670    Select Case ERR.Number
        Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

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

Private Sub cmdOK_Click()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRevCode As Long, lngRevType As Long, lngTaxType As Long
        Dim strRevDesc As String, strTaxDesc As String
        Dim varComments As Variant, varRecurItem As Variant

3910    DoCmd.Hourglass True
3920    DoEvents

3930    With Me
3940      lngRevCode = .cmbRevenueCodes_Misc
3950      strRevDesc = .cmbRevenueCodes_Misc.Column(1)
3960      lngRevType = .cmbRevenueCodes_Misc.Column(2)
3970      lngTaxcode = .cmbTaxCodes_Misc
3980      strTaxDesc = .cmbTaxCodes_Misc.Column(1)
3990      lngTaxType = .cmbTaxCodes_Misc.Column(2)
4000      varComments = .description
4010      varRecurItem = .cmbRecurringItems
4020      If IsNull(varRecurItem) = True Then
4030        varRecurItem = "Long Term Capital Loss"
4040      End If
4050    End With

4060    Set dbs = CurrentDb
4070    With dbs

4080      Select Case strCallersCallingForm
          Case "frmJournal"
            ' ** Append tblJournal _Map_Staging2, linked to qryMap_Misc_LTCL_Detail_12b (qryMap_Misc_LTCL_Detail_12a
            ' ** (Journal Map, grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with cnt,
            ' ** Location_ID_min, Location_ID_max), qryMap_Misc_LTCL_Detail_12e (Ledger, just 'INCOME O/U', 'Received',
            ' ** grouped by accountno, with Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno]),
            ' ** to Journal, by specified [desc], [jsub], [revcode], [taxcode], [recitm], GlobalVarGet('gblnGoToReport').
4090        Set qdf = .QueryDefs("qryMap_Misc_LTCL_Detail_02")
4100        With qdf.Parameters
4110          ![astno] = lngAssetNo
4120          ![actno] = strAccountNo_OU
4130        End With
4140      Case "frmJournal_Columns"
            ' ** Append tblJournal_Map_Staging2, linked to qryMap_Misc_LTCL_Detail_52d (qryMap_Misc_LTCL_Detail_52c
            ' ** (qryMap_Misc_LTCL_62b (qryMap_Misc_LTCL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
            ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max), with Location_ID), linked to
            ' ** Location, with Loc_Name, Loc_Name_display), to tblJournal_Column, by specified [desc], [jsub], [revcod],
            ' ** [taxcod], [revdesc], [taxdesc], [revtyp], [taxtyp], [recitm], GlobalVarGet('gblnGoToReport').
4150        Set qdf = .QueryDefs("qryMap_Misc_LTCL_Detail_52")
4160        With qdf.Parameters
4170          ![revdesc] = strRevDesc
4180          ![taxdesc] = strTaxDesc
4190          ![revtyp] = lngRevType
4200          ![taxtyp] = lngTaxType
4210        End With
4220      End Select
4230      With qdf.Parameters
4240        ![Desc] = varComments
4250        ![jsub] = "LTCL"
4260        ![revcod] = lngRevCode
4270        ![taxcod] = lngTaxcode
4280        ![recitm] = varRecurItem
4290      End With
4300      qdf.Execute
4310      Set qdf = Nothing

4320      Select Case strCallersCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
4330        Set qdf = .QueryDefs("qryMap_01")
4340      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
4350        Set qdf = .QueryDefs("qryJournal_Columns_35_04")
4360      End Select
4370      qdf.Execute
4380      Set qdf = Nothing

4390      Select Case strCallersCallingForm
          Case "frmJournal"
            ' ** Empty tblJournal_Map_Staging2.
4400        Set qdf = .QueryDefs("qryMap_Misc_LTCL_18")
4410      Case "frmJournal_Columns"
            ' ** Empty tblJournal_Map_Staging2.
4420        Set qdf = .QueryDefs("qryMap_Misc_LTCL_68")
4430      End Select
4440      qdf.Execute
4450      Set qdf = Nothing

4460      .Close
4470    End With
4480    Set dbs = Nothing

4490    Select Case strCallersCallingForm
        Case "frmJournal"
4500      If IsLoaded(strCallersCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4510        DoCmd.OpenForm strCallersCallingForm, , , , , , THIS_NAME
4520      Else
4530        Forms(strCallersCallingForm).frmJournal_Sub5_Misc.Form.Requery
4540        DoCmd.SelectObject acForm, strCallersCallingForm, False
4550      End If
4560      DoEvents
4570  On Error Resume Next
4580      Forms(strCallersCallingForm).frmJournal_Sub5_Misc.Form.Requery
4590      Forms(strCallersCallingForm).frmJournal_Sub5_Misc.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub5_Misc.
4600  On Error GoTo ERRH
4610    Case "frmJournal_Columns"
4620      gblnSetFocus = True
4630      gstrReturningForm = "frmMap_Return"
4640      If IsLoaded(strCallersCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4650        DoCmd.OpenForm strCallersCallingForm, , , , , , "frmMenu_Post"
4660      Else
4670        Forms(strCallersCallingForm).frmJournal_Columns_Sub.Form.Requery
4680        Forms(strCallersCallingForm).frmJournal_Columns_Sub.Form.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
4690        DoCmd.SelectObject acForm, strCallersCallingForm, False
4700        Forms(strCallersCallingForm).TimerInterval = 250&
4710      End If
4720    End Select

4730    DoCmd.Close acForm, THIS_NAME

EXITP:
4740    DoCmd.Hourglass False
4750    Set qdf = Nothing
4760    Set dbs = Nothing
4770    Exit Sub

ERRH:
4780    DoCmd.Hourglass False
4790    Select Case ERR.Number
        Case Else
4800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4810    End Select
4820    Resume EXITP

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
4990          .cmdCancel.SetFocus
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

Public Sub cmdCancel_Click()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intTmp01 As Integer
        Dim intX As Integer

5210    DoCmd.Hourglass True
5220    DoEvents

5230    intTmp01 = 0
5240    Do While Forms.Count > 1
5250      intTmp01 = intTmp01 + 1
5260      For intX = (Forms.Count - 1) To 0 Step -1
5270        Select Case Forms(intX).Name
            Case THIS_NAME
              ' ** Skip.
5280        Case "frmJournal", "frmJournal_Columns", "frmSiteMap", "frmSiteMap_Journal"
              ' ** Skip.
5290        Case Else
5300          DoCmd.Close acForm, Forms(intX).Name
5310          Exit For  ' ** Do them one-at-a-time.
5320        End Select
5330      Next
5340      If intTmp01 >= 10 Then Exit Do  ' ** Don't get caught in a loop.
5350    Loop

5360    Set dbs = CurrentDb
5370    With dbs
5380      Select Case strCallersCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
5390        Set qdf = .QueryDefs("qryMap_01")
5400      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
5410        Set qdf = .QueryDefs("qryJournal_Columns_35_04")
5420      End Select
5430      qdf.Execute
5440      .Close
5450    End With

5460    gblnSetFocus = True
5470    Select Case strCallersCallingForm
        Case "frmJournal"
5480      If IsLoaded(strCallersCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5490        DoCmd.OpenForm strCallersCallingForm, , , , , , THIS_NAME
5500      Else
5510        gstrReturningForm = THIS_NAME
5520        Forms(strCallersCallingForm).frmJournal_Sub5_Misc.Form.Requery
5530        DoCmd.SelectObject acForm, strCallersCallingForm, False
5540        Forms(strCallersCallingForm).TimerInterval = 250&
5550      End If
5560    Case "frmJournal_Columns"
5570      If IsLoaded(strCallersCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5580        DoCmd.OpenForm strCallersCallingForm, , , , , , "frmMenu_Post"
5590      Else
5600        gstrReturningForm = THIS_NAME
5610        Forms(strCallersCallingForm).frmJournal_Columns_Sub.Form.Requery
5620        DoCmd.SelectObject acForm, strCallersCallingForm, False
5630        Forms(strCallersCallingForm).TimerInterval = 250&
5640      End If
5650    End Select

5660    DoCmd.Close acForm, THIS_NAME

EXITP:
5670    Set qdf = Nothing
5680    Set dbs = Nothing
5690    Exit Sub

ERRH:
5700    DoCmd.Hourglass False
5710    Select Case ERR.Number
        Case Else
5720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5730    End Select
5740    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5820    intShiftDown = (Shift And acShiftMask) > 0
5830    intAltDown = (Shift And acAltMask) > 0
5840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5860      Select Case intRetVal
          Case vbKeyTab
5870        With Me
5880          intRetVal = 0
5890          lngRecsCur = .frmMap_Misc_LTCL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_LTCL_Detail_Sub.
5900          If lngRecsCur > 0 Then
5910            .frmMap_Misc_LTCL_Detail_Sub.SetFocus
5920            .frmMap_Misc_LTCL_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMap_Misc_LTCL_Detail_Sub.
5930            .frmMap_Misc_LTCL_Detail_Sub.Form.FocusHolder.SetFocus
5940          Else
5950            .description.SetFocus
5960          End If
5970        End With
5980      End Select
5990    End If

        ' ** Shift keys.
6000    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6010      Select Case intRetVal
          Case vbKeyTab
6020        With Me
6030          intRetVal = 0
6040          .cmdOK.SetFocus
6050        End With
6060      End Select
6070    End If

EXITP:
6080    KeyCode = intRetVal
6090    Exit Sub

ERRH:
6100    intRetVal = 0
6110    Select Case ERR.Number
        Case Else
6120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6130    End Select
6140    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

6210    With Me
6220      blnPrintReport_Focus = True
6230      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
6240      .cmdPrintReport_raised_img.Visible = False
6250      .cmdPrintReport_raised_focus_img.Visible = False
6260      .cmdPrintReport_raised_focus_dots_img.Visible = False
6270      .cmdPrintReport_sunken_focus_dots_img.Visible = False
6280      .cmdPrintReport_raised_img_dis.Visible = False
6290    End With

EXITP:
6300    Exit Sub

ERRH:
6310    Select Case ERR.Number
        Case Else
6320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6330    End Select
6340    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

6410    With Me
6420      blnPrintReport_MouseDown = True
6430      .cmdPrintReport_sunken_focus_dots_img.Visible = True
6440      .cmdPrintReport_raised_img.Visible = False
6450      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6460      .cmdPrintReport_raised_focus_img.Visible = False
6470      .cmdPrintReport_raised_focus_dots_img.Visible = False
6480      .cmdPrintReport_raised_img_dis.Visible = False
6490    End With

EXITP:
6500    Exit Sub

ERRH:
6510    Select Case ERR.Number
        Case Else
6520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6530    End Select
6540    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

6610    With Me
6620      strDocName = "rptMap_Misc_LTCL"
6630      gstrReportCallingForm = THIS_NAME
6640      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
6650        .Modal = False
6660        DoCmd.OpenReport strDocName, acViewPreview, , , , strCallersCallingForm & "~" & CStr(acViewPreview)
6670        DoCmd.Maximize
6680        DoCmd.RunCommand acCmdFitToWindow
6690        Select Case strCallersCallingForm
            Case "frmJournal"
6700          Forms(strCallersCallingForm).Visible = False
6710          .Visible = False
6720        Case "frmJournal_Columns"
6730          .Visible = False
6740        End Select
6750      Else
            '##GTR_Ref: rptMap_Misc_LTCL
6760        DoCmd.OpenReport strDocName, acViewNormal, , , , strCallersCallingForm & "~" & CStr(acViewNormal)
6770      End If
6780    End With

EXITP:
6790    Exit Sub

ERRH:
6800    Forms(strCallersCallingForm).Visible = True
6810    Me.Visible = True
6820    DoCmd.Restore
6830    Select Case ERR.Number
        Case Else
6840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6850    End Select
6860    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6920    intShiftDown = (Shift And acShiftMask) > 0
6930    intAltDown = (Shift And acAltMask) > 0
6940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6960      Select Case intRetVal
          Case vbKeyTab
6970        With Me
6980          intRetVal = 0
6990          .cmdOK.SetFocus
7000        End With
7010      End Select
7020    End If

        ' ** Shift keys.
7030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7040      Select Case intRetVal
          Case vbKeyTab
7050        With Me
7060          intRetVal = 0
7070          If .cmbTaxCodes_Misc.Visible = True And .cmbTaxCodes_Misc.Enabled = True Then
7080            .cmbTaxCodes_Misc.SetFocus
7090          ElseIf .cmbRevenueCodes_Misc.Visible = True And .cmbRevenueCodes_Misc.Enabled = True Then
7100            .cmbRevenueCodes_Misc.SetFocus
7110          Else
7120            .description.SetFocus
7130          End If
7140        End With
7150      End Select
7160    End If

EXITP:
7170    KeyCode = intRetVal
7180    Exit Sub

ERRH:
7190    intRetVal = 0
7200    Select Case ERR.Number
        Case Else
7210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7220    End Select
7230    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

7310    With Me
7320      If blnPrintReport_MouseDown = False Then
7330        Select Case blnPrintReport_Focus
            Case True
7340          .cmdPrintReport_raised_focus_dots_img.Visible = True
7350          .cmdPrintReport_raised_focus_img.Visible = False
7360        Case False
7370          .cmdPrintReport_raised_focus_img.Visible = True
7380          .cmdPrintReport_raised_focus_dots_img.Visible = False
7390        End Select
7400        .cmdPrintReport_raised_img.Visible = False
7410        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7420        .cmdPrintReport_sunken_focus_dots_img.Visible = False
7430        .cmdPrintReport_raised_img_dis.Visible = False
7440      End If
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

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

7610    With Me
7620      .cmdPrintReport_raised_focus_dots_img.Visible = True
7630      .cmdPrintReport_raised_img.Visible = False
7640      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7650      .cmdPrintReport_raised_focus_img.Visible = False
7660      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7670      .cmdPrintReport_raised_img_dis.Visible = False
7680      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

7810    With Me
7820      .cmdPrintReport_raised_img.Visible = True
7830      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7840      .cmdPrintReport_raised_focus_img.Visible = False
7850      .cmdPrintReport_raised_focus_dots_img.Visible = False
7860      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7870      .cmdPrintReport_raised_img_dis.Visible = False
7880      blnPrintReport_Focus = False
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

Private Sub total_shareface_sum_KeyDown(KeyCode As Integer, Shift As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_sum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8020    intShiftDown = (Shift And acShiftMask) > 0
8030    intAltDown = (Shift And acAltMask) > 0
8040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8070        With Me
8080          intRetVal = 0
8090          .description.SetFocus
8100        End With
8110      End Select
8120    End If

        ' ** Shift keys.
8130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8150        With Me
8160          intRetVal = 0
8170          lngRecsCur = .frmMap_Misc_LTCL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_LTCL_Detail_Sub.
8180          If lngRecsCur > 0& Then
8190            .frmMap_Misc_LTCL_Detail_Sub.SetFocus
8200            .frmMap_Misc_LTCL_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Misc_LTCL_Detail_Sub.
8210            .frmMap_Misc_LTCL_Detail_Sub.Form.FocusHolder.SetFocus
8220          Else
8230            .cmdCancel.SetFocus
8240          End If
8250        End With
8260      End Select
8270    End If

EXITP:
8280    KeyCode = intRetVal
8290    Exit Sub

ERRH:
8300    intRetVal = 0
8310    Select Case ERR.Number
        Case Else
8320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8330    End Select
8340    Resume EXITP

End Sub

Private Sub total_cash_sum_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "total_cash_sum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8420    intShiftDown = (Shift And acShiftMask) > 0
8430    intAltDown = (Shift And acAltMask) > 0
8440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8470        With Me
8480          intRetVal = 0
8490          .description.SetFocus
8500        End With
8510      End Select
8520    End If

        ' ** Shift keys.
8530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8550        With Me
8560          intRetVal = 0
8570          lngRecsCur = .frmMap_Misc_LTCL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_LTCL_Detail_Sub.
8580          If lngRecsCur > 0& Then
8590            .frmMap_Misc_LTCL_Detail_Sub.SetFocus
8600            .frmMap_Misc_LTCL_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Misc_LTCL_Detail_Sub.
8610            .frmMap_Misc_LTCL_Detail_Sub.Form.FocusHolder.SetFocus
8620          Else
8630            .cmdCancel.SetFocus
8640          End If
8650        End With
8660      End Select
8670    End If

EXITP:
8680    KeyCode = intRetVal
8690    Exit Sub

ERRH:
8700    intRetVal = 0
8710    Select Case ERR.Number
        Case Else
8720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8730    End Select
8740    Resume EXITP

End Sub

Private Sub Overshort_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "Overshort_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8820    intShiftDown = (Shift And acShiftMask) > 0
8830    intAltDown = (Shift And acAltMask) > 0
8840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8870        With Me
8880          intRetVal = 0
8890          .description.SetFocus
8900        End With
8910      End Select
8920    End If

        ' ** Shift keys.
8930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8950        With Me
8960          intRetVal = 0
8970          lngRecsCur = .frmMap_Misc_LTCL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_LTCL_Detail_Sub.
8980          If lngRecsCur > 0& Then
8990            .frmMap_Misc_LTCL_Detail_Sub.SetFocus
9000            .frmMap_Misc_LTCL_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Misc_LTCL_Detail_Sub.
9010            .frmMap_Misc_LTCL_Detail_Sub.Form.FocusHolder.SetFocus
9020          Else
9030            .cmdCancel.SetFocus
9040          End If
9050        End With
9060      End Select
9070    End If

EXITP:
9080    KeyCode = intRetVal
9090    Exit Sub

ERRH:
9100    intRetVal = 0
9110    Select Case ERR.Number
        Case Else
9120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9130    End Select
9140    Resume EXITP

End Sub

Private Sub cmbRecurringItems_Enter()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_Enter"

9210    With Me
9220      If IsNull(.cmbRecurringItems) = True Or .cmbRecurringItems.text = vbNullString Then
9230        ExpandCombo .cmbRecurringItems  ' ** Module Procedure: modUtilities.
9240      End If
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

Private Sub cmbRecurringItems_AfterUpdate()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_AfterUpdate"

9410    With Me
9420      If IsNull(.cmbRecurringItems) = True Then
9430        MsgBox "A Description/Recurring Item is required for a Misc. Journal Type.", vbInformation + vbOKOnly, "Entry Required"
9440        .cmbRecurringItems.SetFocus
9450      Else
9460        If InStr(.cmbRecurringItems.text, Chr(34)) > 0 Then
9470          MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
9480          .cmbRecurringItems.SetFocus
9490        End If
9500      End If
9510    End With

EXITP:
9520    Exit Sub

ERRH:
9530    Select Case ERR.Number
        Case Else
9540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9550    End Select
9560    Resume EXITP

End Sub

Private Sub cmbRecurringItems_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_KeyDown"

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
9770          lngRecsCur = .frmMap_Misc_LTCL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_LTCL_Detail_Sub.
9780          If lngRecsCur > 0& Then
9790            .frmMap_Misc_LTCL_Detail_Sub.SetFocus
9800            .frmMap_Misc_LTCL_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Misc_LTCL_Detail_Sub.
9810            .frmMap_Misc_LTCL_Detail_Sub.Form.FocusHolder.SetFocus
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

Private Sub cmbRecurringItems_NotInList(NewData As String, Response As Integer)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_NotInList"

10010   With Me
10020     .cmbRecurringItems.SelStart = 0
10030     .cmbRecurringItems.SelLength = Len(.cmbRecurringItems.text)
10040   End With

EXITP:
10050   Exit Sub

ERRH:
10060   Select Case ERR.Number
        Case Else
10070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10080   End Select
10090   Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

10110   With Me
10120     Select Case .curr_date.Visible
          Case True
10130       .curr_date.Visible = False
10140     Case False
10150       .curr_date = .curr_id.Column(CBX_C_DATE)
10160       .curr_date.Visible = True
10170     End Select
10180   End With

EXITP:
10190   Exit Sub

ERRH:
10200   Select Case ERR.Number
        Case Else
10210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
10220   End Select
10230   Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String
        Dim lngX As Long

10310   With Me

10320     If IsNull(.curr_id) = False Then

10330       lngCurrID = .curr_id.Column(CBX_C_CURRID)

10340       blnFound = False
10350       If lngCurrID <> 150& Then
10360         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
10370           NoChar_Load  ' ** Procedure: Below.
10380         End If
10390         For lngX = 0& To (lngNoChars - 1&)
10400           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
10410             blnFound = True
10420           End If
10430         Next
10440       End If

10450       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
10460         .total_cash_sum.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
10470         .total_cash_sum.FontSize = 10
10480         .total_cash_sum.FontBold = True
10490         .overshort.FontName = "Courier New"
10500         .overshort.FontSize = 10
10510         .overshort.FontBold = True
10520         .pershare.FontName = "Courier New"
10530         .pershare.FontSize = 10
10540         .pershare.FontBold = True
10550       Case False
10560         If .total_cash_sum.FontName <> "Arial" Then
10570           .total_cash_sum.FontName = "Arial"
10580           .total_cash_sum.FontSize = 10
10590           .total_cash_sum.FontBold = False
10600           .overshort.FontName = "Arial"
10610           .overshort.FontSize = 10
10620           .overshort.FontBold = False
10630           .pershare.FontName = "Arial"
10640           .pershare.FontSize = 10
10650           .pershare.FontBold = False
10660         End If
10670       End Select

10680       If lngCurrID = 150& Then  ' ** USD.
10690         strFrmt1 = "Currency"
10700         intDec1 = 2
10710         strFrmt2 = "Currency"
10720         intDec2 = 5
10730         lngBkClr = CLR_DISABLED_BG_VLR
10740       Else
10750         strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
10760         If .curr_id.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
10770           intDec1 = 2
10780         Else
10790           intDec1 = .curr_id.Column(CBX_C_DEC)  ' ** Decimal places.
10800         End If
10810         intDec2 = 5
10820         intLen = Len(strTmp01)
10830         Select Case intLen
              Case 1
                'strTmp01 = strTmp01
10840         Case 2
10850           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
10860         Case 3
10870           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
10880         Case 4
10890           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
10900         Case 5
10910           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                  "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
10920         End Select
10930         Select Case intDec1
              Case 0
10940           strFrmt1 = strTmp01 & "#,##0;(" & strTmp01 & "#,##0);" & strTmp01 & "0;" & strTmp01 & "0"
10950           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
10960         Case 1  ' ** None have this currently.
10970           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
10980           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
10990         Case 2
11000           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
11010           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11020         Case 3
11030           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
11040           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11050         End Select
11060         lngBkClr = CLR_VLTPRP
11070       End If
11080       .total_cash_sum.Format = strFrmt1
11090       .total_cash_sum.DecimalPlaces = intDec1
11100       .total_cash_sum.BackColor = lngBkClr
11110       .overshort.Format = strFrmt1
11120       .overshort.DecimalPlaces = intDec1
11130       .overshort.BackColor = CLR_DISABLED_BG
11140       .pershare.Format = strFrmt2
11150       .pershare.DecimalPlaces = intDec2
11160       .pershare.BackColor = lngBkClr

11170       If lngCurrID = 150& Then
11180         .total_cash_sum_usd.Visible = False
11190         .Overshort_usd.Visible = False
11200         .pershare_usd.Visible = False
11210       Else
11220         .total_cash_sum_usd = 0
11230         .total_cash_sum_usd.BackColor = CLR_VLTRED
11240         .total_cash_sum_usd.Visible = True
11250         .Overshort_usd = 0
11260         .Overshort_usd.Visible = True
11270         .pershare_usd = 0
11280         .pershare_usd.BackColor = CLR_VLTRED
11290         .pershare_usd.Visible = True
11300       End If

11310     End If

11320   End With

EXITP:
11330   Exit Sub

ERRH:
11340   Select Case ERR.Number
        Case Else
11350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11360   End Select
11370   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

11410   With Me
11420     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
11430     .curr_id.SetFocus
11440     Response = acDataErrContinue
11450   End With

EXITP:
11460   Exit Sub

ERRH:
11470   Select Case ERR.Number
        Case Else
11480     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11490   End Select
11500   Resume EXITP

End Sub

Private Sub description_GotFocus()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

11610   With Me
11620     .description.SelLength = 0
11630     .description.SelStart = 999
11640   End With

EXITP:
11650   Exit Sub

ERRH:
11660   Select Case ERR.Number
        Case Else
11670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11680   End Select
11690   Resume EXITP

End Sub

Private Sub description_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11720   intShiftDown = (Shift And acShiftMask) > 0
11730   intAltDown = (Shift And acAltMask) > 0
11740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11770       With Me
11780         intRetVal = 0
11790         If .cmbRevenueCodes_Misc.Visible = True And .cmbRevenueCodes_Misc.Enabled = True Then
11800           .cmbRevenueCodes_Misc.SetFocus
11810         ElseIf .cmbTaxCodes_Misc.Visible = True And .cmbTaxCodes_Misc.Enabled = True Then
11820           .cmbTaxCodes_Misc.SetFocus
11830         Else
11840           .cmdOK.SetFocus
11850         End If
11860       End With
11870     End Select
11880   End If

        ' ** Shift keys.
11890   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11900     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11910       With Me
11920         intRetVal = 0
11930         .cmbRecurringItems.SetFocus
11940       End With
11950     End Select
11960   End If

EXITP:
11970   KeyCode = intRetVal
11980   Exit Sub

ERRH:
11990   intRetVal = 0
12000   Select Case ERR.Number
        Case Else
12010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12020   End Select
12030   Resume EXITP

End Sub

Private Sub description_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "description_MouseMove"

12110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12120   Exit Sub

ERRH:
12130   Select Case ERR.Number
        Case Else
12140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12150   End Select
12160   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_display_GotFocus()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_display_GotFocus"

12210   With Me
12220     .cmbRevenueCodes_Misc.SetFocus
12230   End With

EXITP:
12240   Exit Sub

ERRH:
12250   Select Case ERR.Number
        Case Else
12260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12270   End Select
12280   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_display_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12320   intShiftDown = (Shift And acShiftMask) > 0
12330   intAltDown = (Shift And acAltMask) > 0
12340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12370       With Me
12380         intRetVal = 0
12390         .cmbRevenueCodes_Misc.SetFocus
12400       End With
12410     End Select
12420   End If

        ' ** Shift keys.
12430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12450       With Me
12460         intRetVal = 0
12470         .description.SetFocus
12480       End With
12490     End Select
12500   End If

EXITP:
12510   KeyCode = intRetVal
12520   Exit Sub

ERRH:
12530   intRetVal = 0
12540   Select Case ERR.Number
        Case Else
12550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12560   End Select
12570   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_Enter()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_Enter"

12610   With Me
12620     Select Case IsNull(.cmbRevenueCodes_Misc)
          Case True
12630       ExpandCombo .cmbRevenueCodes_Misc  ' ** Module Procedure: modUtilities.
12640     Case False
12650       If .cmbRevenueCodes_Misc = REVID_INC Or .cmbRevenueCodes_Misc = REVID_EXP Then
12660         ExpandCombo .cmbRevenueCodes_Misc  ' ** Module Procedure: modUtilities.
12670       End If
12680     End Select
12690   End With

EXITP:
12700   Exit Sub

ERRH:
12710   Select Case ERR.Number
        Case Else
12720     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12730   End Select
12740   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_AfterUpdate()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_AfterUpdate"

12810   With Me
12820     Select Case IsNull(.cmbRevenueCodes_Misc)
          Case True
12830       .cmbRevenueCodes_Misc = REVID_INC
12840     Case False
            ' ** Nothing at the moment.
12850     End Select
12860     If .cmbRevenueCodes_Misc = REVID_INC Or .cmbRevenueCodes_Misc = REVID_EXP Then
12870       .cmbRevenueCodes_Misc_display = Null
12880     Else
12890       .cmbRevenueCodes_Misc_display = .cmbRevenueCodes_Misc.Column(1)
12900     End If
12910   End With

EXITP:
12920   Exit Sub

ERRH:
12930   Select Case ERR.Number
        Case Else
12940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12950   End Select
12960   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_KeyDown(KeyCode As Integer, Shift As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13020   intShiftDown = (Shift And acShiftMask) > 0
13030   intAltDown = (Shift And acAltMask) > 0
13040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13070       With Me
13080         intRetVal = 0
13090         If .cmbTaxCodes_Misc.Visible = True And .cmbTaxCodes_Misc.Enabled = True Then
13100           .cmbTaxCodes_Misc.SetFocus
13110         ElseIf .cmdPrintReport.Enabled = True Then
13120           .cmdPrintReport.SetFocus
13130         Else
13140           .cmdOK.SetFocus
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
13230         .description.SetFocus
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

Private Sub cmbTaxCodes_Misc_display_GotFocus()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_display_GotFocus"

13410   With Me
13420     .cmbTaxCodes_Misc.SetFocus
13430   End With

EXITP:
13440   Exit Sub

ERRH:
13450   Select Case ERR.Number
        Case Else
13460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13470   End Select
13480   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Misc_display_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13520   intShiftDown = (Shift And acShiftMask) > 0
13530   intAltDown = (Shift And acAltMask) > 0
13540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13570       With Me
13580         intRetVal = 0
13590         .cmbTaxCodes_Misc.SetFocus
13600       End With
13610     End Select
13620   End If

        ' ** Shift keys.
13630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13650       With Me
13660         intRetVal = 0
13670         If .cmbRevenueCodes_Misc.Visible = True And .cmbRevenueCodes_Misc.Enabled = True Then
13680           .cmbRevenueCodes_Misc.SetFocus
13690         Else
13700           .description.SetFocus
13710         End If
13720       End With
13730     End Select
13740   End If

EXITP:
13750   KeyCode = intRetVal
13760   Exit Sub

ERRH:
13770   intRetVal = 0
13780   Select Case ERR.Number
        Case Else
13790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13800   End Select
13810   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Misc_Enter()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_Enter"

13910   With Me
13920     Select Case IsNull(.cmbTaxCodes_Misc)
          Case True
13930       ExpandCombo .cmbTaxCodes_Misc  ' ** Module Procedure: modUtilities.
13940     Case False
13950       If .cmbTaxCodes_Misc = TAXID_INC Or .cmbTaxCodes_Misc = TAXID_DED Then
13960         ExpandCombo .cmbTaxCodes_Misc  ' ** Module Procedure: modUtilities.
13970       End If
13980     End Select
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

Private Sub cmbTaxCodes_Misc_AfterUpdate()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_AfterUpdate"

14110   With Me
14120     Select Case IsNull(.cmbTaxCodes_Misc)
          Case True
14130       .cmbTaxCodes_Misc = TAXID_INC
14140     Case False
            ' ** Nothing at the moment.
14150     End Select
14160     If .cmbTaxCodes_Misc = TAXID_INC Or .cmbTaxCodes_Misc = TAXID_DED Then
14170       .cmbTaxCodes_Misc_display = Null
14180     Else
14190       .cmbTaxCodes_Misc_display = .cmbTaxCodes_Misc.Column(1)
14200     End If
14210   End With

EXITP:
14220   Exit Sub

ERRH:
14230   Select Case ERR.Number
        Case Else
14240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14250   End Select
14260   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Misc_KeyDown(KeyCode As Integer, Shift As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_KeyDown"

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
14390         If .cmdPrintReport.Enabled = True Then
14400           .cmdPrintReport.SetFocus
14410         Else
14420           .cmdOK.SetFocus
14430         End If
14440       End With
14450     End Select
14460   End If

        ' ** Shift keys.
14470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14490       With Me
14500         intRetVal = 0
14510         If .cmbRevenueCodes_Misc.Visible = True And .cmbRevenueCodes_Misc.Enabled = True Then
14520           .cmbRevenueCodes_Misc.SetFocus
14530         Else
14540           .description.SetFocus
14550         End If
14560       End With
14570     End Select
14580   End If

EXITP:
14590   KeyCode = intRetVal
14600   Exit Sub

ERRH:
14610   intRetVal = 0
14620   Select Case ERR.Number
        Case Else
14630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14640   End Select
14650   Resume EXITP

End Sub

Public Function RecalcTots() As Boolean
' ** I give up trying to get it from the subform!

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim dblTmp01 As Double, dblTmp02 As Double
        Dim lngX As Long
        Dim blnRetVal As Boolean

14710   With Me
14720     blnRetVal = False
14730     Set rst = .frmMap_Misc_LTCL_Detail_Sub.Form.RecordsetClone
14740     With rst
14750       If .BOF = True And .EOF = True Then
              ' ** Ho-Hum...
14760       Else
14770         .MoveLast
14780         lngRecs = .RecordCount
14790         .MoveFirst
14800         dblTmp01 = 0#: dblTmp02 = 0#
14810         For lngX = 1& To lngRecs
14820           dblTmp01 = dblTmp01 + ![total_shareface]
14830           dblTmp02 = dblTmp02 + ![total_pcash]
14840           If lngX < lngRecs Then .MoveNext
14850         Next
14860       End If
14870       .Close
14880     End With
14890     If dblTmp01 <> 0# Then
14900       blnRetVal = True
14910     End If
14920     .total_shareface_sum = dblTmp01
14930     .total_cash_sum = dblTmp02
14940   End With

EXITP:
14950   Set rst = Nothing
14960   RecalcTots = blnRetVal
14970   Exit Function

ERRH:
14980   blnRetVal = False
14990   Select Case ERR.Number
        Case Else
15000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15010   End Select
15020   Resume EXITP

End Function

Public Sub IncludeCurrency(blnCalc As Boolean)

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim dblRate2 As Double
        Dim blnFound As Boolean
        Dim dblTmp01 As Double

15110   With Me
15120     Select Case blnCalc
          Case True
15130       If blnCurrID = True Then
15140         blnFound = False
15150         Do While blnFound = False
15160           If IsNull(.total_cash_sum) = False Then
15170             If .total_cash_sum <> 0 Then
15180               blnFound = True
15190               dblRate2 = .curr_id.Column(CBX_C_RATE2)
15200               dblTmp01 = (.total_cash_sum * dblRate2)
15210               dblTmp01 = Round(dblTmp01, 2)
15220               .total_cash_sum_usd = dblTmp01
15230               dblTmp01 = (.overshort * dblRate2)
15240               dblTmp01 = Round(dblTmp01, 2)
15250               .Overshort_usd = dblTmp01
15260               dblTmp01 = (.pershare * dblRate2)
15270               dblTmp01 = Round(dblTmp01, 5)
15280               .pershare_usd = dblTmp01
15290               If lngCurrID <> 150& Then
15300                 .total_cash_sum_usd.Visible = True
15310                 .Overshort_usd.Visible = True
15320                 .pershare_usd.Visible = True
15330               End If
15340             End If
15350           End If
15360         Loop
15370       End If
15380     Case False
15390       .curr_id.Enabled = False
15400       .curr_id.Locked = True
15410       .curr_id.ForeColor = CLR_DISABLED_FG
15420       .curr_id.BackColor = CLR_VLTGRN
15430       If blnCurrID = True And lngCurrID <> 150& Then
15440         .curr_id.Visible = True
15450         .curr_id_cmd.Visible = True
15460         .curr_id_cmd.Enabled = True
15470       Else
15480         .curr_id.Visible = False
15490         .curr_id_cmd.Visible = False
15500         .curr_id_cmd.Enabled = False
15510         .cmbRevenueCodes_Misc_vline05.Visible = False
15520         .cmbRevenueCodes_Misc_vline06.Visible = False
15530         .cmbRevenueCodes_Misc_box.Width = (.cmbRevenueCodes_Misc_vline06.Left - .cmbRevenueCodes_Misc_box.Left)
15540         .overshort.Top = .Overshort_alt_line.Top
15550         .Overshort_lbl.Top = (.Overshort_alt_line.Top + lngTpp)
15560         .pershare.Top = .pershare_alt_line.Top
15570         .pershare_lbl.Top = (.pershare_alt_line.Top + lngTpp)
15580       End If
15590       .total_cash_sum_usd.Visible = False
15600       .Overshort_usd.Visible = False
15610       .pershare_usd.Visible = False
15620     End Select
15630   End With

EXITP:
15640   Exit Sub

ERRH:
15650   DoCmd.Hourglass False
15660   Select Case ERR.Number
        Case Else
15670     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15680   End Select
15690   Resume EXITP

End Sub

Private Sub IncludeCurrency_Sub()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngTmp01 As Long, lngTmp02 As Long

15710   With Me
15720     If blnCurrID = True And lngCurrID <> 150& Then

15730       With .frmMap_Misc_LTCL_Detail_Sub.Form
15740         lngTmp01 = .Width - .total_pcash_alt_line.Left
15750       End With

15760       lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
15770       lngMonitorNum = 1&: lngTmp02 = 0&
15780       EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
15790       If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

15800       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
15810       DoCmd.MoveSize (lngFrm_Left - (lngTmp01 / 2)), lngTmp02, (lngFrm_Width + lngTmp01), lngFrm_Height  'lngFrm_Top
15820       If lngMonitorNum > 1& Then
15830         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
15840       End If

15850       .Width = (.Width + lngTmp01)
15860       .frmMap_Misc_LTCL_Detail_Sub.Width = (.frmMap_Misc_LTCL_Detail_Sub.Width + lngTmp01)
15870       .frmMap_Misc_LTCL_Detail_Sub_box.Width = (.frmMap_Misc_LTCL_Detail_Sub_box.Width + lngTmp01)

15880       .Header_vline01.Left = .Width
15890       .Header_vline02.Left = .Width
15900       .Footer_vline01.Left = .Width
15910       .Footer_vline02.Left = .Width
15920       .Header_hline01.Width = .Width
15930       .Header_hline02.Width = .Width
15940       .Footer_hline01.Width = .Width
15950       .Footer_hline02.Width = .Width

15960       .total_cash_sum_usd.Left = ((.total_cash_sum.Left + .total_cash_sum.Width) + (4& * lngTpp))
15970       .total_cash_sum_usd.Top = .total_cash_sum.Top
15980       .Overshort_usd.Left = ((.overshort.Left + .overshort.Width) + (4& * lngTpp))
15990       .overshort.Top = .Overshort_alt_line.Top
16000       .Overshort_lbl.Top = (.Overshort_alt_line.Top + lngTpp)
16010       .Overshort_usd.Top = .Overshort_alt_line.Top
16020       .pershare_usd.Left = ((.total_cash_sum_usd.Left + .total_cash_sum_usd.Width) - .pershare_usd.Width)
16030       .pershare.Top = .pershare_alt_line.Top
16040       .pershare_lbl.Top = (.pershare_alt_line.Top + lngTpp)
16050       .pershare_usd.Top = .pershare_alt_line.Top

16060       .total_cash_sum_usd.ForeColor = CLR_DISABLED_FG
16070       .total_cash_sum_usd.BackColor = CLR_VLTRED
16080       .Overshort_usd.ForeColor = CLR_DISABLED_FG
16090       .Overshort_usd.BackColor = CLR_DISABLED_BG
16100       .pershare_usd.ForeColor = CLR_DISABLED_FG
16110       .pershare_usd.BackColor = CLR_VLTRED

16120       .maptype.Left = (.maptype.Left + lngTmp01)
16130       .cmdCancel.Left = (.cmdCancel.Left + lngTmp01)
16140       .cmdOK.Left = (.cmdOK.Left + lngTmp01)

16150     End If
16160   End With

EXITP:
16170   Exit Sub

ERRH:
16180   DoCmd.Hourglass False
16190   Select Case ERR.Number
        Case Else
16200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16210   End Select
16220   Resume EXITP

End Sub

Private Sub NoChar_Load()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

16310   Set dbs = CurrentDb
16320   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
16330     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
16340     Set rst = qdf.OpenRecordset
16350     With rst
16360       .MoveLast
16370       lngNoChars = .RecordCount
16380       .MoveFirst
16390       arr_varNoChar = .GetRows(lngNoChars)
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
16400       .Close
16410     End With
16420     .Close
16430   End With

EXITP:
16440   Set rst = Nothing
16450   Set qdf = Nothing
16460   Set dbs = Nothing
16470   Exit Sub

ERRH:
16480   DoCmd.Hourglass False
16490   Select Case ERR.Number
        Case Else
16500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16510   End Select
16520   Resume EXITP

End Sub
