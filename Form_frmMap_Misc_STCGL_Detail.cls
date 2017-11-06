VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Misc_STCGL_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Misc_STCGL_Detail"

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
' **   Short Name:       S {shortname on frmMap_Misc_STCGL_Detail_Sub}
' **   Tax Codes:        T {cmbTaxCodes_Misc}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots}

' ** MapType = "Misc."

' ** Report:
' **   rptMap_Misc_STCG
' **     qryMapReport_08
' **   rptMap_Misc_STCL
' **     qryMapReport_09

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

170             strCallingForm = "frmMap_Misc_STCGL"  ' ** The only form that does.

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
660                 Set qdf = .QueryDefs("qryMap_Misc_STCGL_Detail_Sub_01_02")
670               Case "frmJournal_Columns"
                    ' ** Journal Map, grouped and summed, by accountno.
680                 Set qdf = .QueryDefs("qryMap_Misc_STCGL_Detail_Sub_51_02")
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
830               If .frmMap_Misc_STCGL_Detail_Sub.Form.RecordSource <> "qryMap_Misc_STCGL_Detail_Sub_01_02" Then
840                 .frmMap_Misc_STCGL_Detail_Sub.Form.RecordSource = "qryMap_Misc_STCGL_Detail_Sub_01_02"
850               End If
860             Case "frmJournal_Columns"
870               If .frmMap_Misc_STCGL_Detail_Sub.Form.RecordSource <> "qryMap_Misc_STCGL_Detail_Sub_51_02" Then
880                 .frmMap_Misc_STCGL_Detail_Sub.Form.RecordSource = "qryMap_Misc_STCGL_Detail_Sub_51_02"
890               End If
900             End Select

910             .maptype = "Misc."
920             .txtNumberRecords = "Number Records = " & CStr(lngRecs)
930             .overshort = dblOverUnder

940             If dblPerShare > 0 Then
950               .cmbRecurringItems = "Short Term Capital Gain"
960               .description = "STCG"
970               .Header_Addition_lbl.Visible = True
980               .Header_Reduction_lbl.Visible = False
990               .Header_lbl.Visible = False
1000              .cusip.Left = ((.Header_Addition_lbl.Left + .Header_Addition_lbl.Width) + (8& * lngTpp))
1010              .totdesc.Left = ((.Header_Addition_lbl.Left + .Header_Addition_lbl.Width) + (8& * lngTpp))
1020            Else
1030              .cmbRecurringItems = "Short Term Capital Loss"
1040              .description = "STCL"
1050              .Header_Reduction_lbl.Visible = True
1060              .Header_Addition_lbl.Visible = False
1070              .Header_lbl.Visible = False
1080              .cusip.Left = ((.Header_Reduction_lbl.Left + .Header_Reduction_lbl.Width) + (8& * lngTpp))
1090              .totdesc.Left = ((.Header_Reduction_lbl.Left + .Header_Reduction_lbl.Width) + (8& * lngTpp))
1100            End If
1110            .pershare = dblPerShare  '5 dec.
1120            .amount = dblAmount

1130            If .cmdOK.Enabled = True Then
1140              .cmdOK.SetFocus
1150            Else
1160              .description.SetFocus
1170            End If

1180            blnRecalcTots = True
1190            .TimerInterval = 500&

1200            DoEvents
1210            DoCmd.SelectObject acForm, THIS_NAME, False
1220            DoEvents

1230            blnIsOpen = False

1240          Else
1250            Cancel = -1
1260          End If
1270        Else
1280          Cancel = -1
1290        End If
1300      Else
1310        Cancel = -1
1320      End If
1330    End With

1340    If Cancel = -1 Then
1350      Beep
1360      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1370    End If

EXITP:
1380    Set rst = Nothing
1390    Set qdf = Nothing
1400    Set dbs = Nothing
1410    Exit Sub

ERRH:
1420    DoCmd.Hourglass False
1430    Select Case ERR.Number
        Case Else
1440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1450    End Select
1460    Resume EXITP

End Sub

Private Sub Form_Load()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1510    With Me

1520      .total_shareface_sum.ForeColor = CLR_DISABLED_FG
1530      .total_shareface_sum.BackColor = CLR_DISABLED_BG
1540      .total_cash_sum.ForeColor = CLR_DISABLED_FG
1550      Select Case IsNull(.total_cash_sum)
          Case True
1560        .total_cash_sum.BackColor = CLR_DISABLED_BG
1570      Case False
1580        If .total_cash_sum >= 0 Then
1590          .total_cash_sum.BackColor = CLR_DISABLED_BG
1600        Else
1610          .total_cash_sum.BackColor = CLR_DISABLED_BG_VLR
1620        End If
1630      End Select
1640      .overshort.ForeColor = CLR_DISABLED_FG
1650      .overshort.BackColor = CLR_DISABLED_BG
1660      .pershare.ForeColor = CLR_DISABLED_FG
1670      Select Case IsNull(.pershare)
          Case True
1680        .pershare.BackColor = CLR_DISABLED_BG
1690      Case False
1700        If .pershare >= 0 Then
1710          .pershare.BackColor = CLR_DISABLED_BG
1720        Else
1730          .pershare.BackColor = CLR_DISABLED_BG_VLR
1740        End If
1750      End Select

          ' ** Borrowing these variables from the Court Reports.
1760      If gstrCrtRpt_Version <> vbNullString Then
1770        .cusip = gstrCrtRpt_Ordinal
1780        .cusip.ForeColor = CLR_DISABLED_FG
1790        .cusip.BackColor = CLR_DISABLED_BG
1800        .cusip.Visible = True
1810        .totdesc = gstrCrtRpt_Version
1820        .totdesc.ForeColor = CLR_DISABLED_FG
1830        .totdesc.BackColor = CLR_DISABLED_BG
1840        .totdesc.Visible = True
1850      End If

1860      .transdate = datTransDate

1870      Select Case IsNull(.cmbRevenueCodes_Misc)
          Case True
1880        .cmbRevenueCodes_Misc = REVID_INC
1890        .cmbRevenueCodes_Misc_display = Null
1900      Case False
1910        If .cmbRevenueCodes_Misc = REVID_INC Or .cmbRevenueCodes_Misc = REVID_EXP Then
1920          .cmbRevenueCodes_Misc_display = Null
1930        Else
1940          .cmbRevenueCodes_Misc_display = .cmbRevenueCodes_Misc.Column(1)
1950        End If
1960      End Select

1970      Select Case IsNull(.cmbTaxCodes_Misc)
          Case True
1980        .cmbTaxCodes_Misc = TAXID_INC
1990        .cmbTaxCodes_Misc_display = Null
2000      Case False
2010        If .cmbTaxCodes_Misc = TAXID_INC Or .cmbTaxCodes_Misc = TAXID_DED Then
2020          .cmbTaxCodes_Misc_display = Null
2030        Else
2040          .cmbTaxCodes_Misc_display = .cmbTaxCodes_Misc.Column(1)
2050        End If
2060      End Select

2070      Select Case gblnRevenueExpenseTracking
          Case True
2080        .cmbRevenueCodes_Misc.Visible = True
2090        .cmbRevenueCodes_Misc_display.Visible = True
2100        .cmbRevenueCodes_Misc_lbl2.Visible = False  ' ** Option off msg.
2110      Case False
2120        .cmbRevenueCodes_Misc.Visible = False
2130        .cmbRevenueCodes_Misc_display.Visible = False
2140        .cmbRevenueCodes_Misc_lbl_box.Visible = False
2150        .cmbRevenueCodes_Misc_lbl2.Visible = True  ' ** Option off msg.
2160      End Select

2170      Select Case gblnIncomeTaxCoding
          Case True
2180        .cmbTaxCodes_Misc.Visible = True
2190        .cmbTaxCodes_Misc_display.Visible = True
2200        .cmbTaxCodes_Misc_lbl2.Visible = False  ' ** Option off msg.
2210      Case False
2220        .cmbTaxCodes_Misc.Visible = False
2230        .cmbTaxCodes_Misc_display.Visible = False
2240        .cmbTaxCodes_Misc_lbl_box.Visible = False  ' ** Option off msg.
2250        .cmbTaxCodes_Misc_lbl2.Visible = True
2260      End Select

2270      If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
2280        .cmbRevenueCodes_Misc_box_lbl.ForeColor = WIN_CLR_DISF
2290        .cmbRevenueCodes_Misc_box_lbl_dim_hi.Visible = True
2300      End If

2310      curr_id_AfterUpdate  ' ** Procedure: Below.
2320      IncludeCurrency False  ' ** Procedure: Below
2330      DoEvents

2340      DoCmd.Hourglass False
2350      .frmMap_Misc_STCGL_Detail_Sub.Form.TimerInterval = 2000&

2360    End With

EXITP:
2370    Exit Sub

ERRH:
2380    DoCmd.Hourglass False
2390    Select Case ERR.Number
        Case Else
2400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2410    End Select
2420    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2520    intShiftDown = (Shift And acShiftMask) > 0
2530    intAltDown = (Shift And acAltMask) > 0
2540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Short Name:       S {shortname on frmMap_Misc_STCGL_Detail_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Comments:         M {description}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc Totals:    F5 {RecalcTots}

        ' ** Plain keys.
2550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2560      Select Case intRetVal
          Case vbKeyEscape
2570        intRetVal = 0
2580        cmdCancel_Click  ' ** Procedure: Below.
2590      Case vbKeyF5
2600        intRetVal = 0
2610        RecalcTots  ' ** Function: Below.
2620      End Select
2630    End If

        ' ** Alt keys.
2640    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2650      Select Case intRetVal
          Case vbKeyS
2660        With Me
2670          intRetVal = 0
2680          .frmMap_Misc_STCGL_Detail_Sub.SetFocus
2690          .frmMap_Misc_STCGL_Detail_Sub.Form.shortname.SetFocus
2700        End With
2710      End Select
2720    End If

        ' ** Ctrl keys.
2730    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2740      Select Case intRetVal
          Case vbKeyM
2750        With Me
2760          intRetVal = 0
2770          .description.SetFocus
2780        End With
2790      End Select
2800    End If

        ' ** Ctrl-Shift keys.
2810    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2820      Select Case intRetVal
          Case vbKeyF
2830        With Me
2840          intRetVal = 0
2850          .FocusHolder.SetFocus
2860        End With
2870      End Select
2880    End If

EXITP:
2890    KeyCode = intRetVal
2900    Exit Sub

ERRH:
2910    intRetVal = 0
2920    Select Case ERR.Number
        Case Else
2930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2940    End Select
2950    Resume EXITP

End Sub

Private Sub Form_Timer()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim blnRetVal As Boolean

3010    With Me
3020      .TimerInterval = 0&
3030      If gblnGoToReport = True Then
3040        Select Case blnGoingToReport
            Case True
3050          DoCmd.Hourglass False
3060          gblnGoToReport = False
3070          blnGoingToReport = False
3080          .GoToReport_arw_pr_img.Visible = False
3090          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
3100          .cmdCancel.SetFocus
3110        Case False
3120          DoCmd.Hourglass True  ' ** Make sure it's still running.
3130          DoEvents
3140          blnGoingToReport = True
3150          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3160          .GoToReport_arw_pr_img.Visible = True
3170          If .cmdPrintReport.Enabled = True Then
3180            .cmdPrintReport.SetFocus
3190          End If
3200          DoEvents
3210          .cmdOK.Enabled = False
3220          Beep
3230          DoCmd.Hourglass False
3240          DoEvents
3250          .TimerInterval = (GTR_WAIT * 2&)
3260        End Select
3270      End If
3280      If blnRecalcTots = True Then
3290        blnRecalcTots = False
3300        blnRetVal = RecalcTots  ' ** Function: Below.
3310        If blnRetVal = False Then
3320          blnRecalcTots = True
3330          .TimerInterval = 500&
3340        End If
3350      End If
3360    End With

EXITP:
3370    Exit Sub

ERRH:
3380    DoCmd.Hourglass False
3390    Select Case ERR.Number
        Case Else
3400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3410    End Select
3420    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3510    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3520    Set dbs = CurrentDb
        ' ** Empty tmpAccount.
3530    Set qdf = dbs.QueryDefs("qryMap_Misc_STCGL_08_01")
3540    qdf.Execute
3550    Set qdf = Nothing
3560    DoEvents
3570    dbs.Close
3580    Set dbs = Nothing
3590    DoEvents

        ' ** Assure all reports are closed.
3600    Do While Reports.Count > 0
3610      DoCmd.Close acReport, Reports(0).Name
3620    Loop

3630    gstrReportCallingForm = vbNullString

EXITP:
3640    Set qdf = Nothing
3650    Set dbs = Nothing
3660    Exit Sub

ERRH:
3670    DoCmd.Hourglass False
3680    Select Case ERR.Number
        Case Else
3690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3700    End Select
3710    Resume EXITP

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
        Dim varComments As Variant, varRecurItem As Variant

4210    DoCmd.Hourglass True
4220    DoEvents

4230    With Me
4240      lngRevCode = .cmbRevenueCodes_Misc
4250      strRevDesc = .cmbRevenueCodes_Misc.Column(1)
4260      lngRevType = .cmbRevenueCodes_Misc.Column(2)
4270      lngTaxcode = .cmbTaxCodes_Misc
4280      strTaxDesc = .cmbTaxCodes_Misc.Column(1)
4290      lngTaxType = .cmbTaxCodes_Misc.Column(2)
4300      varComments = .description
4310      varRecurItem = .cmbRecurringItems
4320      If IsNull(varRecurItem) = True Then
4330        If dblPerShare > 0# Then
4340          varRecurItem = "Short Term Capital Gain"
4350        Else
4360          varRecurItem = "Short Term Capital Loss"
4370        End If
4380      End If
4390    End With

4400    Set dbs = CurrentDb
4410    With dbs

4420      Select Case strCallersCallingForm
          Case "frmJournal"
4430        If dblPerShare > 0# Then
              ' ** Append tblJournal_Map_Staging2, linked to qryMap_Misc_STCGL_12b (qryMap_Misc_STCGL_12a
              ' ** (Journal Map, grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with cnt,
              ' ** Location_ID_min, Location_ID_max), qryMap_Misc_STCGL_12e (Ledger, just 'INCOME O/U', 'Received',
              ' ** grouped by accountno, with Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno]),
              ' ** to Journal, by specified [desc], [jsub], [revcode], [taxcode], [recitm], GlobalVarGet('gblnGoToReport').
4440          Set qdf = .QueryDefs("qryMap_Misc_STCGL_Detail_02a")
4450        Else
              ' ** Append tblJournal_Map_Staging2, linked to qryMap_Misc_STCGL_12b (qryMap_Misc_STCGL_12a
              ' ** (Journal Map, grouped by accountno, assetno, Location_ID), grouped by accountno, assetno, with cnt,
              ' ** Location_ID_min, Location_ID_max), qryMap_Misc_STCGL_12f (Ledger, just 'INCOME O/U', 'Paid',
              ' ** grouped by accountno, with Location_ID_min, Location_ID_max, cnt, by specified [actno], [astno]),
              ' ** to Journal, by specified [desc], [jsub], [revcode], [taxcode], [recitm], GlobalVarGet('gblnGoToReport').
4460          Set qdf = .QueryDefs("qryMap_Misc_STCGL_Detail_02b")
4470        End If
4480        With qdf.Parameters
4490          ![astno] = lngAssetNo
4500          ![actno] = strAccountNo_OU
4510        End With
4520      Case "frmJournal_Columns"
4530        If dblPerShare > 0# Then
              ' ** Append tblJournal_Map_Staging2, linked to qryMap_Misc_STCGL_Detail_52d (qryMap_Misc_STCGL_Detail_52c
              ' ** (qryMap_Misc_STCGL_62b (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
              ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max), with Location_ID), linked to
              ' ** Location, with Loc_Name, Loc_Name_display), to tblJournal_Column, by specified [desc], [jsub], [revcod],
              ' ** [taxcod], [revdesc], [taxdesc], [revtyp], [taxtyp], [recitm], GlobalVarGet('gblnGoToReport').
4540          Set qdf = .QueryDefs("qryMap_Misc_STCGL_Detail_52a")
4550        Else
              ' ** Append tblJournal_Map_Staging2, linked to qryMap_Misc_STCGL_Detail_52d (qryMap_Misc_STCGL_Detail_52c
              ' ** (qryMap_Misc_STCGL_62b (qryMap_Misc_STCGL_62a (Journal Map, grouped by accountno, assetno, Location_ID),
              ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max), with Location_ID), linked to
              ' ** Location, with Loc_Name, Loc_Name_display), to tblJournal_Column, by specified [desc], [jsub], [revcod],
              ' ** [taxcod], [revdesc], [taxdesc], [revtyp], [taxtyp], [recitm], GlobalVarGet('gblnGoToReport').
4560          Set qdf = .QueryDefs("qryMap_Misc_STCGL_Detail_52b")
4570        End If
4580        With qdf.Parameters
4590          ![revdesc] = strRevDesc
4600          ![taxdesc] = strTaxDesc
4610          ![revtyp] = lngRevType
4620          ![taxtyp] = lngTaxType
4630        End With
4640      End Select
4650      With qdf.Parameters
4660        ![Desc] = varComments
4670        If dblPerShare > 0# Then
4680          ![jsub] = "STCG"
4690        Else
4700          ![jsub] = "STCL"
4710        End If
4720        ![revcod] = lngRevCode
4730        ![taxcod] = lngTaxcode
4740        ![recitm] = varRecurItem
4750      End With
4760      qdf.Execute
4770      Set qdf = Nothing

4780      Select Case strCallersCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
4790        Set qdf = .QueryDefs("qryMap_01")
4800      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
4810        Set qdf = .QueryDefs("qryJournal_Columns_35_04")
4820      End Select
4830      qdf.Execute
4840      Set qdf = Nothing

4850      Select Case strCallersCallingForm
          Case "frmJournal"
            ' ** Empty tblJournal_Map_Staging2.
4860        Set qdf = .QueryDefs("qryMap_Misc_STCGL_18")
4870      Case "frmJournal_Columns"
            ' ** Empty tblJournal_Map_Staging2.
4880        Set qdf = .QueryDefs("qryMap_Misc_STCGL_68")
4890      End Select
4900      qdf.Execute
4910      Set qdf = Nothing

4920      .Close
4930    End With
4940    Set dbs = Nothing

4950    Select Case strCallersCallingForm
        Case "frmJournal"
4960      If IsLoaded(strCallersCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4970        DoCmd.OpenForm strCallersCallingForm, , , , , , THIS_NAME
4980      Else
4990        Forms(strCallersCallingForm).frmJournal_Sub5_Misc.Form.Requery
5000        DoCmd.SelectObject acForm, strCallersCallingForm, False
5010      End If
5020      DoEvents
5030  On Error Resume Next
5040      Forms(strCallersCallingForm).frmJournal_Sub5_Misc.Form.Requery
5050      Forms(strCallersCallingForm).frmJournal_Sub5_Misc.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub5_Misc.
5060  On Error GoTo ERRH
5070    Case "frmJournal_Columns"
5080      gblnSetFocus = True
5090      gstrReturningForm = "frmMap_Return"
5100      If IsLoaded(strCallersCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5110        DoCmd.OpenForm strCallersCallingForm, , , , , , "frmMenu_Post"
5120      Else
5130        Forms(strCallersCallingForm).frmJournal_Columns_Sub.Form.Requery
5140        Forms(strCallersCallingForm).frmJournal_Columns_Sub.Form.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
5150        DoCmd.SelectObject acForm, strCallersCallingForm, False
5160        Forms(strCallersCallingForm).TimerInterval = 250&
5170      End If
5180    End Select

5190    DoCmd.Close acForm, THIS_NAME

EXITP:
5200    DoCmd.Hourglass False
5210    Set qdf = Nothing
5220    Set dbs = Nothing
5230    Exit Sub

ERRH:
5240    DoCmd.Hourglass False
5250    Select Case ERR.Number
        Case Else
5260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5270    End Select
5280    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5320    intShiftDown = (Shift And acShiftMask) > 0
5330    intAltDown = (Shift And acAltMask) > 0
5340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab
5370        With Me
5380          intRetVal = 0
5390          .cmdCancel.SetFocus
5400        End With
5410      End Select
5420    End If

        ' ** Shift keys.
5430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5440      Select Case intRetVal
          Case vbKeyTab
5450        With Me
5460          intRetVal = 0
5470          .cmdPrintReport.SetFocus
5480        End With
5490      End Select
5500    End If

EXITP:
5510    KeyCode = intRetVal
5520    Exit Sub

ERRH:
5530    intRetVal = 0
5540    Select Case ERR.Number
        Case Else
5550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5560    End Select
5570    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intTmp01 As Integer
        Dim intX As Integer

5610    DoCmd.Hourglass True
5620    DoEvents

5630    intTmp01 = 0
5640    Do While Forms.Count > 1
5650      intTmp01 = intTmp01 + 1
5660      For intX = (Forms.Count - 1) To 0 Step -1
5670        Select Case Forms(intX).Name
            Case THIS_NAME
              ' ** Skip.
5680        Case "frmJournal", "frmJournal_Columns", "frmSiteMap", "frmSiteMap_Journal"
              ' ** Skip.
5690        Case Else
5700          DoCmd.Close acForm, Forms(intX).Name
5710          Exit For  ' ** Do them one-at-a-time.
5720        End Select
5730      Next
5740      If intTmp01 >= 10 Then Exit Do  ' ** Don't get caught in a loop.
5750    Loop

5760    Set dbs = CurrentDb
5770    With dbs
5780      Select Case strCallersCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
5790        Set qdf = .QueryDefs("qryMap_01")
5800      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
5810        Set qdf = .QueryDefs("qryJournal_Columns_35_04")
5820      End Select
5830      qdf.Execute
5840      .Close
5850    End With

5860    gblnSetFocus = True
5870    Select Case strCallersCallingForm
        Case "frmJournal"
5880      If IsLoaded(strCallersCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5890        DoCmd.OpenForm strCallersCallingForm, , , , , , THIS_NAME
5900      Else
5910        gstrReturningForm = THIS_NAME
5920        Forms(strCallersCallingForm).frmJournal_Sub5_Misc.Form.Requery
5930        DoCmd.SelectObject acForm, strCallersCallingForm, False
5940        Forms(strCallersCallingForm).TimerInterval = 250&
5950      End If
5960    Case "frmJournal_Columns"
5970      If IsLoaded(strCallersCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5980        DoCmd.OpenForm strCallersCallingForm, , , , , , "frmMenu_Post"
5990      Else
6000        gstrReturningForm = THIS_NAME
6010        Forms(strCallersCallingForm).frmJournal_Columns_Sub.Form.Requery
6020        DoCmd.SelectObject acForm, strCallersCallingForm, False
6030        Forms(strCallersCallingForm).TimerInterval = 250&
6040      End If
6050    End Select

6060    DoCmd.Close acForm, THIS_NAME

EXITP:
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

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
6290          lngRecsCur = .frmMap_Misc_STCGL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_STCGL_Detail_Sub.
6300          If lngRecsCur > 0 Then
6310            .frmMap_Misc_STCGL_Detail_Sub.SetFocus
6320            .frmMap_Misc_STCGL_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMap_Misc_STCGL_Detail_Sub.
6330            .frmMap_Misc_STCGL_Detail_Sub.Form.FocusHolder.SetFocus
6340          Else
6350            .description.SetFocus
6360          End If
6370        End With
6380      End Select
6390    End If

        ' ** Shift keys.
6400    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6410      Select Case intRetVal
          Case vbKeyTab
6420        With Me
6430          intRetVal = 0
6440          .cmdOK.SetFocus
6450        End With
6460      End Select
6470    End If

EXITP:
6480    KeyCode = intRetVal
6490    Exit Sub

ERRH:
6500    intRetVal = 0
6510    Select Case ERR.Number
        Case Else
6520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6530    End Select
6540    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

6610    With Me
6620      blnPrintReport_Focus = True
6630      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
6640      .cmdPrintReport_raised_img.Visible = False
6650      .cmdPrintReport_raised_focus_img.Visible = False
6660      .cmdPrintReport_raised_focus_dots_img.Visible = False
6670      .cmdPrintReport_sunken_focus_dots_img.Visible = False
6680      .cmdPrintReport_raised_img_dis.Visible = False
6690    End With

EXITP:
6700    Exit Sub

ERRH:
6710    Select Case ERR.Number
        Case Else
6720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6730    End Select
6740    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

6810    With Me
6820      blnPrintReport_MouseDown = True
6830      .cmdPrintReport_sunken_focus_dots_img.Visible = True
6840      .cmdPrintReport_raised_img.Visible = False
6850      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
6860      .cmdPrintReport_raised_focus_img.Visible = False
6870      .cmdPrintReport_raised_focus_dots_img.Visible = False
6880      .cmdPrintReport_raised_img_dis.Visible = False
6890    End With

EXITP:
6900    Exit Sub

ERRH:
6910    Select Case ERR.Number
        Case Else
6920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6930    End Select
6940    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

7010    With Me
7020      If dblPerShare > 0# Then
7030        strDocName = "rptMap_Misc_STCG"
7040      Else
7050        strDocName = "rptMap_Misc_STCL"
7060      End If
7070      gstrReportCallingForm = THIS_NAME
7080      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
7090        .Modal = False
7100        DoCmd.OpenReport strDocName, acViewPreview, , , , strCallersCallingForm & "~" & CStr(acViewPreview)
7110        DoCmd.Maximize
7120        DoCmd.RunCommand acCmdFitToWindow
7130        Select Case strCallersCallingForm
            Case "frmJournal"
7140          Forms(strCallersCallingForm).Visible = False
7150          .Visible = False
7160        Case "frmJournal_Columns"
7170          .Visible = False
7180        End Select
7190      Else
7200        If dblPerShare > 0# Then
              '##GTR_Ref: rptMap_Misc_STCG
7210          DoCmd.OpenReport strDocName, acViewNormal, , , , strCallersCallingForm & "~" & CStr(acViewNormal)
7220        Else
              '##GTR_Ref: rptMap_Misc_STCL
7230          DoCmd.OpenReport strDocName, acViewNormal, , , , strCallersCallingForm & "~" & CStr(acViewNormal)
7240        End If
7250      End If
7260    End With

EXITP:
7270    Exit Sub

ERRH:
7280    Forms(strCallersCallingForm).Visible = True
7290    Me.Visible = True
7300    DoCmd.Restore
7310    Select Case ERR.Number
        Case Else
7320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7330    End Select
7340    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
7490          .cmdOK.SetFocus
7500        End With
7510      End Select
7520    End If

        ' ** Shift keys.
7530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7540      Select Case intRetVal
          Case vbKeyTab
7550        With Me
7560          intRetVal = 0
7570          If .cmbTaxCodes_Misc.Visible = True And .cmbTaxCodes_Misc.Enabled = True Then
7580            .cmbTaxCodes_Misc.SetFocus
7590          ElseIf .cmbRevenueCodes_Misc.Visible = True And .cmbRevenueCodes_Misc.Enabled = True Then
7600            .cmbRevenueCodes_Misc.SetFocus
7610          Else
7620            .description.SetFocus
7630          End If
7640        End With
7650      End Select
7660    End If

EXITP:
7670    KeyCode = intRetVal
7680    Exit Sub

ERRH:
7690    intRetVal = 0
7700    Select Case ERR.Number
        Case Else
7710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7720    End Select
7730    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

7810    With Me
7820      If blnPrintReport_MouseDown = False Then
7830        Select Case blnPrintReport_Focus
            Case True
7840          .cmdPrintReport_raised_focus_dots_img.Visible = True
7850          .cmdPrintReport_raised_focus_img.Visible = False
7860        Case False
7870          .cmdPrintReport_raised_focus_img.Visible = True
7880          .cmdPrintReport_raised_focus_dots_img.Visible = False
7890        End Select
7900        .cmdPrintReport_raised_img.Visible = False
7910        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7920        .cmdPrintReport_sunken_focus_dots_img.Visible = False
7930        .cmdPrintReport_raised_img_dis.Visible = False
7940      End If
7950    End With

EXITP:
7960    Exit Sub

ERRH:
7970    Select Case ERR.Number
        Case Else
7980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7990    End Select
8000    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

8110    With Me
8120      .cmdPrintReport_raised_focus_dots_img.Visible = True
8130      .cmdPrintReport_raised_img.Visible = False
8140      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8150      .cmdPrintReport_raised_focus_img.Visible = False
8160      .cmdPrintReport_sunken_focus_dots_img.Visible = False
8170      .cmdPrintReport_raised_img_dis.Visible = False
8180      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

8310    With Me
8320      .cmdPrintReport_raised_img.Visible = True
8330      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8340      .cmdPrintReport_raised_focus_img.Visible = False
8350      .cmdPrintReport_raised_focus_dots_img.Visible = False
8360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
8370      .cmdPrintReport_raised_img_dis.Visible = False
8380      blnPrintReport_Focus = False
8390    End With

EXITP:
8400    Exit Sub

ERRH:
8410    Select Case ERR.Number
        Case Else
8420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8430    End Select
8440    Resume EXITP

End Sub

Private Sub total_shareface_sum_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_sum_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
8570        With Me
8580          intRetVal = 0
8590          .description.SetFocus
8600        End With
8610      End Select
8620    End If

        ' ** Shift keys.
8630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8650        With Me
8660          intRetVal = 0
8670          lngRecsCur = .frmMap_Misc_STCGL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_STCGL_Detail_Sub.
8680          If lngRecsCur > 0& Then
8690            .frmMap_Misc_STCGL_Detail_Sub.SetFocus
8700            .frmMap_Misc_STCGL_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Misc_STCGL_Detail_Sub.
8710            .frmMap_Misc_STCGL_Detail_Sub.Form.FocusHolder.SetFocus
8720          Else
8730            .cmdCancel.SetFocus
8740          End If
8750        End With
8760      End Select
8770    End If

EXITP:
8780    KeyCode = intRetVal
8790    Exit Sub

ERRH:
8800    intRetVal = 0
8810    Select Case ERR.Number
        Case Else
8820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8830    End Select
8840    Resume EXITP

End Sub

Private Sub total_cash_sum_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "total_cash_sum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8920    intShiftDown = (Shift And acShiftMask) > 0
8930    intAltDown = (Shift And acAltMask) > 0
8940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8970        With Me
8980          intRetVal = 0
8990          .description.SetFocus
9000        End With
9010      End Select
9020    End If

        ' ** Shift keys.
9030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9050        With Me
9060          intRetVal = 0
9070          lngRecsCur = .frmMap_Misc_STCGL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_STCGL_Detail_Sub.
9080          If lngRecsCur > 0& Then
9090            .frmMap_Misc_STCGL_Detail_Sub.SetFocus
9100            .frmMap_Misc_STCGL_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Misc_STCGL_Detail_Sub.
9110            .frmMap_Misc_STCGL_Detail_Sub.Form.FocusHolder.SetFocus
9120          Else
9130            .cmdCancel.SetFocus
9140          End If
9150        End With
9160      End Select
9170    End If

EXITP:
9180    KeyCode = intRetVal
9190    Exit Sub

ERRH:
9200    intRetVal = 0
9210    Select Case ERR.Number
        Case Else
9220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9230    End Select
9240    Resume EXITP

End Sub

Private Sub Overshort_KeyDown(KeyCode As Integer, Shift As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "Overshort_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9320    intShiftDown = (Shift And acShiftMask) > 0
9330    intAltDown = (Shift And acAltMask) > 0
9340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9370        With Me
9380          intRetVal = 0
9390          .description.SetFocus
9400        End With
9410      End Select
9420    End If

        ' ** Shift keys.
9430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9450        With Me
9460          intRetVal = 0
9470          lngRecsCur = .frmMap_Misc_STCGL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_STCGL_Detail_Sub.
9480          If lngRecsCur > 0& Then
9490            .frmMap_Misc_STCGL_Detail_Sub.SetFocus
9500            .frmMap_Misc_STCGL_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Misc_STCGL_Detail_Sub.
9510            .frmMap_Misc_STCGL_Detail_Sub.Form.FocusHolder.SetFocus
9520          Else
9530            .cmdCancel.SetFocus
9540          End If
9550        End With
9560      End Select
9570    End If

EXITP:
9580    KeyCode = intRetVal
9590    Exit Sub

ERRH:
9600    intRetVal = 0
9610    Select Case ERR.Number
        Case Else
9620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9630    End Select
9640    Resume EXITP

End Sub

Private Sub cmbRecurringItems_Enter()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_Enter"

9710    With Me
9720      If IsNull(.cmbRecurringItems) = True Or .cmbRecurringItems.text = vbNullString Then
9730        ExpandCombo .cmbRecurringItems  ' ** Module Procedure: modUtilities.
9740      End If
9750    End With

EXITP:
9760    Exit Sub

ERRH:
9770    Select Case ERR.Number
        Case Else
9780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9790    End Select
9800    Resume EXITP

End Sub

Private Sub cmbRecurringItems_AfterUpdate()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_AfterUpdate"

9910    With Me
9920      If IsNull(.cmbRecurringItems) = True Then
9930        MsgBox "A Description/Recurring Item is required for a Misc. Journal Type.", vbInformation + vbOKOnly, "Entry Required"
9940        .cmbRecurringItems.SetFocus
9950      Else
9960        If InStr(.cmbRecurringItems.text, Chr(34)) > 0 Then
9970          MsgBox "Cannot have quote marks in text.", vbInformation + vbOKOnly, "Invalid Characters"
9980          .cmbRecurringItems.SetFocus
9990        End If
10000     End If
10010   End With

EXITP:
10020   Exit Sub

ERRH:
10030   Select Case ERR.Number
        Case Else
10040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10050   End Select
10060   Resume EXITP

End Sub

Private Sub cmbRecurringItems_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_KeyDown"

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
10270         lngRecsCur = .frmMap_Misc_STCGL_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Misc_STCGL_Detail_Sub.
10280         If lngRecsCur > 0& Then
10290           .frmMap_Misc_STCGL_Detail_Sub.SetFocus
10300           .frmMap_Misc_STCGL_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Misc_STCGL_Detail_Sub.
10310           .frmMap_Misc_STCGL_Detail_Sub.Form.FocusHolder.SetFocus
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

Private Sub cmbRecurringItems_NotInList(NewData As String, Response As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRecurringItems_NotInList"

10510   With Me
10520     .cmbRecurringItems.SelStart = 0
10530     .cmbRecurringItems.SelLength = Len(.cmbRecurringItems.text)
10540   End With

EXITP:
10550   Exit Sub

ERRH:
10560   Select Case ERR.Number
        Case Else
10570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10580   End Select
10590   Resume EXITP

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
11210         strFrmt2 = "Currency"
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
11450           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11460         Case 1  ' ** None have this currently.
11470           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
11480           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11490         Case 2
11500           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
11510           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11520         Case 3
11530           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
11540           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11550         End Select
11560         lngBkClr = CLR_VLTGRN
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
11770         .pershare_usd.Visible = True
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
12290         If .cmbRevenueCodes_Misc.Visible = True And .cmbRevenueCodes_Misc.Enabled = True Then
12300           .cmbRevenueCodes_Misc.SetFocus
12310         ElseIf .cmbTaxCodes_Misc.Visible = True And .cmbTaxCodes_Misc.Enabled = True Then
12320           .cmbTaxCodes_Misc.SetFocus
12330         Else
12340           .cmdOK.SetFocus
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
12430         .cmbRecurringItems.SetFocus
12440       End With
12450     End Select
12460   End If

EXITP:
12470   KeyCode = intRetVal
12480   Exit Sub

ERRH:
12490   intRetVal = 0
12500   Select Case ERR.Number
        Case Else
12510     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12520   End Select
12530   Resume EXITP

End Sub

Private Sub description_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "description_MouseMove"

12610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12620   Exit Sub

ERRH:
12630   Select Case ERR.Number
        Case Else
12640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12650   End Select
12660   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_display_GotFocus()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_display_GotFocus"

12710   With Me
12720     .cmbRevenueCodes_Misc.SetFocus
12730   End With

EXITP:
12740   Exit Sub

ERRH:
12750   Select Case ERR.Number
        Case Else
12760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12770   End Select
12780   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_display_KeyDown(KeyCode As Integer, Shift As Integer)

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12820   intShiftDown = (Shift And acShiftMask) > 0
12830   intAltDown = (Shift And acAltMask) > 0
12840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12870       With Me
12880         intRetVal = 0
12890         .cmbRevenueCodes_Misc.SetFocus
12900       End With
12910     End Select
12920   End If

        ' ** Shift keys.
12930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12950       With Me
12960         intRetVal = 0
12970         .description.SetFocus
12980       End With
12990     End Select
13000   End If

EXITP:
13010   KeyCode = intRetVal
13020   Exit Sub

ERRH:
13030   intRetVal = 0
13040   Select Case ERR.Number
        Case Else
13050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13060   End Select
13070   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_Enter()

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_Enter"

13110   With Me
13120     Select Case IsNull(.cmbRevenueCodes_Misc)
          Case True
13130       ExpandCombo .cmbRevenueCodes_Misc  ' ** Module Procedure: modUtilities.
13140     Case False
13150       If .cmbRevenueCodes_Misc = REVID_INC Or .cmbRevenueCodes_Misc = REVID_EXP Then
13160         ExpandCombo .cmbRevenueCodes_Misc  ' ** Module Procedure: modUtilities.
13170       End If
13180     End Select
13190   End With

EXITP:
13200   Exit Sub

ERRH:
13210   Select Case ERR.Number
        Case Else
13220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13230   End Select
13240   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_AfterUpdate()

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_AfterUpdate"

13310   With Me
13320     Select Case IsNull(.cmbRevenueCodes_Misc)
          Case True
13330       .cmbRevenueCodes_Misc = REVID_INC
13340     Case False
            ' ** Nothing at the moment.
13350     End Select
13360     If .cmbRevenueCodes_Misc = REVID_INC Or .cmbRevenueCodes_Misc = REVID_EXP Then
13370       .cmbRevenueCodes_Misc_display = Null
13380     Else
13390       .cmbRevenueCodes_Misc_display = .cmbRevenueCodes_Misc.Column(1)
13400     End If
13410   End With

EXITP:
13420   Exit Sub

ERRH:
13430   Select Case ERR.Number
        Case Else
13440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13450   End Select
13460   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Misc_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Misc_KeyDown"

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
13590         If .cmbTaxCodes_Misc.Visible = True And .cmbTaxCodes_Misc.Enabled = True Then
13600           .cmbTaxCodes_Misc.SetFocus
13610         Else
13620           .cmdOK.SetFocus
13630         End If
13640       End With
13650     End Select
13660   End If

        ' ** Shift keys.
13670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13680     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13690       With Me
13700         intRetVal = 0
13710         .description.SetFocus
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

Private Sub cmbTaxCodes_Misc_display_GotFocus()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_display_GotFocus"

13910   With Me
13920     .cmbTaxCodes_Misc.SetFocus
13930   End With

EXITP:
13940   Exit Sub

ERRH:
13950   Select Case ERR.Number
        Case Else
13960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13970   End Select
13980   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Misc_display_KeyDown(KeyCode As Integer, Shift As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14020   intShiftDown = (Shift And acShiftMask) > 0
14030   intAltDown = (Shift And acAltMask) > 0
14040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14070       With Me
14080         intRetVal = 0
14090         .cmbTaxCodes_Misc.SetFocus
14100       End With
14110     End Select
14120   End If

        ' ** Shift keys.
14130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14150       With Me
14160         intRetVal = 0
14170         If .cmbRevenueCodes_Misc.Visible = True And .cmbRevenueCodes_Misc.Enabled = True Then
14180           .cmbRevenueCodes_Misc.SetFocus
14190         Else
14200           .description.SetFocus
14210         End If
14220       End With
14230     End Select
14240   End If

EXITP:
14250   KeyCode = intRetVal
14260   Exit Sub

ERRH:
14270   intRetVal = 0
14280   Select Case ERR.Number
        Case Else
14290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14300   End Select
14310   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Misc_Enter()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_Enter"

14410   With Me
14420     Select Case IsNull(.cmbTaxCodes_Misc)
          Case True
14430       ExpandCombo .cmbTaxCodes_Misc  ' ** Module Procedure: modUtilities.
14440     Case False
14450       If .cmbTaxCodes_Misc = TAXID_INC Or .cmbTaxCodes_Misc = TAXID_DED Then
14460         ExpandCombo .cmbTaxCodes_Misc  ' ** Module Procedure: modUtilities.
14470       End If
14480     End Select
14490   End With

EXITP:
14500   Exit Sub

ERRH:
14510   Select Case ERR.Number
        Case Else
14520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14530   End Select
14540   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Misc_AfterUpdate()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_AfterUpdate"

14610   With Me
14620     Select Case IsNull(.cmbTaxCodes_Misc)
          Case True
14630       .cmbTaxCodes_Misc = TAXID_INC
14640     Case False
            ' ** Nothing at the moment.
14650     End Select
14660     If .cmbTaxCodes_Misc = TAXID_INC Or .cmbTaxCodes_Misc = TAXID_DED Then
14670       .cmbTaxCodes_Misc_display = Null
14680     Else
14690       .cmbTaxCodes_Misc_display = .cmbTaxCodes_Misc.Column(1)
14700     End If
14710   End With

EXITP:
14720   Exit Sub

ERRH:
14730   Select Case ERR.Number
        Case Else
14740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14750   End Select
14760   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Misc_KeyDown(KeyCode As Integer, Shift As Integer)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Misc_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14820   intShiftDown = (Shift And acShiftMask) > 0
14830   intAltDown = (Shift And acAltMask) > 0
14840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14870       With Me
14880         intRetVal = 0
14890         .cmdOK.SetFocus
14900       End With
14910     End Select
14920   End If

        ' ** Shift keys.
14930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14950       With Me
14960         intRetVal = 0
14970         If .cmbRevenueCodes_Misc.Visible = True And .cmbRevenueCodes_Misc.Enabled = True Then
14980           .cmbRevenueCodes_Misc.SetFocus
14990         Else
15000           .description.SetFocus
15010         End If
15020       End With
15030     End Select
15040   End If

EXITP:
15050   KeyCode = intRetVal
15060   Exit Sub

ERRH:
15070   intRetVal = 0
15080   Select Case ERR.Number
        Case Else
15090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15100   End Select
15110   Resume EXITP

End Sub

Public Function RecalcTots() As Boolean
' ** I give up trying to get it from the subform!

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim dblTmp01 As Double, dblTmp02 As Double
        Dim lngX As Long
        Dim blnRetVal As Boolean

15210   With Me
15220     blnRetVal = False
15230     Set rst = .frmMap_Misc_STCGL_Detail_Sub.Form.RecordsetClone
15240     With rst
15250       If .BOF = True And .EOF = True Then
              ' ** Ho-Hum...
15260       Else
15270         .MoveLast
15280         lngRecs = .RecordCount
15290         .MoveFirst
15300         dblTmp01 = 0#: dblTmp02 = 0#
15310         For lngX = 1& To lngRecs
15320           dblTmp01 = dblTmp01 + ![total_shareface]
15330           dblTmp02 = dblTmp02 + ![total_icash]
15340           If lngX < lngRecs Then .MoveNext
15350         Next
15360       End If
15370       .Close
15380     End With
15390     If dblTmp01 <> 0# Then
15400       blnRetVal = True
15410     End If
15420     .total_shareface_sum = dblTmp01
15430     .total_cash_sum = dblTmp02
15440     If dblTmp02 < 0 Then
15450       If blnCurrID = True And lngCurrID <> 150& Then
15460         If .total_cash_sum.BackColor <> CLR_VLTPRP Then
15470           .total_cash_sum.BackColor = CLR_VLTPRP
15480         End If
15490       Else
15500         If .total_cash_sum.BackColor <> CLR_VLTRED Then
15510           .total_cash_sum.BackColor = CLR_VLTRED
15520         End If
15530       End If
15540     Else
15550       If blnCurrID = True And lngCurrID <> 150& Then
15560         If .total_cash_sum.BackColor <> CLR_VLTGRN Then
15570           .total_cash_sum.BackColor = CLR_VLTGRN
15580         End If
15590       Else
15600         If .total_cash_sum.BackColor <> CLR_DISABLED_BG Then
15610           .total_cash_sum.BackColor = CLR_DISABLED_BG
15620         End If
15630       End If
15640     End If
15650     If .pershare < 0 Then
15660       If blnCurrID = True And lngCurrID <> 150& Then
15670         If .pershare.BackColor <> CLR_VLTPRP Then
15680           .pershare.BackColor = CLR_VLTPRP
15690         End If
15700       Else
15710         If .pershare.BackColor <> CLR_VLTRED Then
15720           .pershare.BackColor = CLR_VLTRED
15730         End If
15740       End If
15750     Else
15760       If blnCurrID = True And lngCurrID <> 150& Then
15770         If .pershare.BackColor <> CLR_VLTGRN Then
15780           .pershare.BackColor = CLR_VLTGRN
15790         End If
15800       Else
15810         If .pershare.BackColor <> CLR_DISABLED_BG Then
15820           .pershare.BackColor = CLR_DISABLED_BG
15830         End If
15840       End If
15850     End If
15860   End With

EXITP:
15870   Set rst = Nothing
15880   RecalcTots = blnRetVal
15890   Exit Function

ERRH:
15900   blnRetVal = False
15910   Select Case ERR.Number
        Case Else
15920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15930   End Select
15940   Resume EXITP

End Function

Public Sub IncludeCurrency(blnCalc As Boolean)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim dblRate2 As Double
        Dim blnFound As Boolean
        Dim dblTmp01 As Double, dblTmp02 As Double, dblTmp03 As Double

16010   With Me
16020     Select Case blnCalc
          Case True
16030       If blnCurrID = True Then
16040         blnFound = False
16050         Do While blnFound = False
16060           If IsNull(.total_cash_sum) = False Then
16070             If .total_cash_sum <> 0 Then
16080               blnFound = True
16090               If lngCurrID <> 150& Then
16100                 dblRate2 = .curr_id.Column(CBX_C_RATE2)
16110                 dblTmp01 = (.total_cash_sum * dblRate2)
16120                 dblTmp01 = Round(dblTmp01, 2)
16130                 .total_cash_sum_usd = dblTmp01
16140                 dblTmp02 = (.overshort * dblRate2)
16150                 dblTmp02 = Round(dblTmp02, 2)
16160                 .Overshort_usd = dblTmp02
16170                 dblTmp03 = (.pershare * dblRate2)
16180                 dblTmp03 = Round(dblTmp03, 5)
16190                 .pershare_usd = dblTmp03
16200                 If dblTmp01 >= 0 Then
16210                   .total_cash_sum.BackColor = CLR_VLTGRN
16220                   .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
16230                 Else
16240                   .total_cash_sum.BackColor = CLR_VLTPRP
16250                   .total_cash_sum_usd.BackColor = CLR_VLTRED
16260                 End If
16270                 If dblTmp02 >= 0 Then
16280                   .overshort.BackColor = CLR_VLTGRN
16290                   .Overshort_usd.BackColor = CLR_DISABLED_BG
16300                 Else
16310                   .overshort.BackColor = CLR_VLTPRP
16320                   .Overshort_usd.BackColor = CLR_VLTRED
16330                 End If
16340                 If dblTmp03 >= 0 Then
16350                   .pershare.BackColor = CLR_VLTGRN
16360                   .pershare_usd.BackColor = CLR_DISABLED_BG
16370                 Else
16380                   .pershare.BackColor = CLR_VLTPRP
16390                   .pershare_usd.BackColor = CLR_VLTRED
16400                 End If
16410                 .total_cash_sum_usd.Visible = True
16420                 .Overshort_usd.Visible = True
16430                 .pershare_usd.Visible = True
16440               Else
16450                 If .total_cash_sum >= 0 Then
16460                   .total_cash_sum.BackColor = CLR_DISABLED_BG
16470                 Else
16480                   .total_cash_sum.BackColor = CLR_VLTRED
16490                 End If
16500                 If .overshort >= 0 Then
16510                   .overshort.BackColor = CLR_DISABLED_BG
16520                 Else
16530                   .overshort.BackColor = CLR_VLTRED
16540                 End If
16550                 If .pershare >= 0 Then
16560                   .pershare.BackColor = CLR_DISABLED_BG
16570                 Else
16580                   .pershare.BackColor = CLR_VLTRED
16590                 End If
16600               End If
16610             End If
16620           End If
16630         Loop
16640       End If
16650     Case False
16660       .curr_id.Enabled = False
16670       .curr_id.Locked = True
16680       .curr_id.ForeColor = CLR_DISABLED_FG
16690       .curr_id.BackColor = CLR_VLTGRN
16700       If blnCurrID = True And lngCurrID <> 150& Then
16710         .curr_id.Visible = True
16720         .curr_id_cmd.Visible = True
16730         .curr_id_cmd.Enabled = True
16740         If .overshort >= 0 Then
16750           .overshort.BackColor = CLR_VLTGRN
16760           .Overshort_usd.BackColor = CLR_DISABLED_BG
16770         Else
16780           .overshort.BackColor = CLR_VLTPRP
16790           .Overshort_usd.BackColor = CLR_VLTRED
16800         End If
16810       Else
16820         .curr_id.Visible = False
16830         .curr_id_cmd.Visible = False
16840         .curr_id_cmd.Enabled = False
16850         .cmbRevenueCodes_Misc_vline05.Visible = False
16860         .cmbRevenueCodes_Misc_vline06.Visible = False
16870         .cmbRevenueCodes_Misc_box.Width = (.cmbRevenueCodes_Misc_vline06.Left - .cmbRevenueCodes_Misc_box.Left)
16880         .overshort.Top = .Overshort_alt_line.Top
16890         .Overshort_lbl.Top = (.Overshort_alt_line.Top + lngTpp)
16900         .pershare.Top = .pershare_alt_line.Top
16910         .pershare_lbl.Top = (.pershare_alt_line.Top + lngTpp)
16920         If IsNull(.total_cash_sum) = True Then
16930           .total_cash_sum.BackColor = CLR_DISABLED_BG
16940           .pershare.BackColor = CLR_DISABLED_BG
16950         Else
16960           If .total_cash_sum >= 0 Then
16970             .total_cash_sum.BackColor = CLR_DISABLED_BG
16980             .pershare.BackColor = CLR_DISABLED_BG
16990           Else
17000             .total_cash_sum.BackColor = CLR_VLTRED
17010             .pershare.BackColor = CLR_VLTRED
17020           End If
17030         End If
17040       End If
17050       .total_cash_sum_usd.Visible = False
17060       .Overshort_usd.Visible = False
17070       .pershare_usd.Visible = False
17080     End Select
17090   End With

EXITP:
17100   Exit Sub

ERRH:
17110   DoCmd.Hourglass False
17120   Select Case ERR.Number
        Case Else
17130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17140   End Select
17150   Resume EXITP

End Sub

Private Sub IncludeCurrency_Sub()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngTmp01 As Long, lngTmp02 As Long

17210   With Me
17220     If blnCurrID = True And lngCurrID <> 150& Then

17230       With .frmMap_Misc_STCGL_Detail_Sub.Form
17240         lngTmp01 = .Width - .total_icash_alt_line.Left
17250       End With

17260       lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
17270       lngMonitorNum = 1&: lngTmp02 = 0&
17280       EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
17290       If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

17300       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
17310       DoCmd.MoveSize (lngFrm_Left - (lngTmp01 / 2)), lngTmp02, (lngFrm_Width + lngTmp01), lngFrm_Height  'lngFrm_Top
17320       If lngMonitorNum > 1& Then
17330         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
17340       End If

17350       .Width = (.Width + lngTmp01)
17360       .frmMap_Misc_STCGL_Detail_Sub.Width = (.frmMap_Misc_STCGL_Detail_Sub.Width + lngTmp01)
17370       .frmMap_Misc_STCGL_Detail_Sub_box.Width = (.frmMap_Misc_STCGL_Detail_Sub_box.Width + lngTmp01)

17380       .Header_vline01.Left = .Width
17390       .Header_vline02.Left = .Width
17400       .Footer_vline01.Left = .Width
17410       .Footer_vline02.Left = .Width
17420       .Header_hline01.Width = .Width
17430       .Header_hline02.Width = .Width
17440       .Footer_hline01.Width = .Width
17450       .Footer_hline02.Width = .Width

17460       .total_cash_sum_usd.Left = ((.total_cash_sum.Left + .total_cash_sum.Width) + (4& * lngTpp))
17470       .total_cash_sum_usd.Top = .total_cash_sum.Top
17480       .Overshort_usd.Left = ((.overshort.Left + .overshort.Width) + (4& * lngTpp))
17490       .overshort.Top = .Overshort_alt_line.Top
17500       .Overshort_lbl.Top = (.Overshort_alt_line.Top + lngTpp)
17510       .Overshort_usd.Top = .Overshort_alt_line.Top
17520       .pershare_usd.Left = ((.total_cash_sum_usd.Left + .total_cash_sum_usd.Width) - .pershare_usd.Width)
17530       .pershare.Top = .pershare_alt_line.Top
17540       .pershare_lbl.Top = (.pershare_alt_line.Top + lngTpp)
17550       .pershare_usd.Top = .pershare_alt_line.Top

17560       Select Case IsNull(.total_cash_sum_usd)
            Case True
17570         .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
17580         .Overshort_usd.BackColor = CLR_DISABLED_BG
17590         .pershare_usd.BackColor = CLR_DISABLED_BG
17600       Case False
17610         If .total_cash_sum_usd >= 0 Then
17620           .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
17630           .pershare_usd.BackColor = CLR_DISABLED_BG
17640         Else
17650           .total_cash_sum_usd.BackColor = CLR_VLTRED
17660           .pershare_usd.BackColor = CLR_VLTRED
17670         End If
17680         If .overshort >= 0 Then
17690           .Overshort_usd.BackColor = CLR_DISABLED_BG
17700         Else
17710           .Overshort_usd.BackColor = CLR_VLTRED
17720         End If
17730       End Select
17740       .total_cash_sum_usd.ForeColor = CLR_DISABLED_FG
17750       .Overshort_usd.ForeColor = CLR_DISABLED_FG
17760       .pershare_usd.ForeColor = CLR_DISABLED_FG

17770       .maptype.Left = (.maptype.Left + lngTmp01)
17780       .cmdCancel.Left = (.cmdCancel.Left + lngTmp01)
17790       .cmdOK.Left = (.cmdOK.Left + lngTmp01)

17800     End If
17810   End With

EXITP:
17820   Exit Sub

ERRH:
17830   DoCmd.Hourglass False
17840   Select Case ERR.Number
        Case Else
17850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17860   End Select
17870   Resume EXITP

End Sub

Private Sub NoChar_Load()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

17910   Set dbs = CurrentDb
17920   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
17930     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
17940     Set rst = qdf.OpenRecordset
17950     With rst
17960       .MoveLast
17970       lngNoChars = .RecordCount
17980       .MoveFirst
17990       arr_varNoChar = .GetRows(lngNoChars)
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
18000       .Close
18010     End With
18020     .Close
18030   End With

EXITP:
18040   Set rst = Nothing
18050   Set qdf = Nothing
18060   Set dbs = Nothing
18070   Exit Sub

ERRH:
18080   DoCmd.Hourglass False
18090   Select Case ERR.Number
        Case Else
18100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18110   End Select
18120   Resume EXITP

End Sub
