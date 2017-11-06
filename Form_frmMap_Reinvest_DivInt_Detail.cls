VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Reinvest_DivInt_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Reinvest_DivInt_Detail"

'VGC 08/30/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   IncludeCurrency_Sub()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Show Location:    L {chkShowLocation}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Purchase}
' **   OK:               O {cmdOK}
' **   Print Report:     P {cmdPrintReport}
' **   Tax Codes:        T {cmbTaxCodes_Purchase}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots}

' ** MapType = "Purchase"

'rptMap_Reinvest_Div
'rptMap_Reinvest_Int
' ** tblJournal, by specified FormRef('Price').
' ** qryMapReport_05.

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

Private lngForm_Width As Long, lngSub_Width As Long, lngOK_Left As Long, lngCancel_Left As Long, lngMapType_Left As Long
Private strFrmt1 As String, strFrmt2 As String, intDec1 As Integer, intDec2 As Integer, lngBkClr As Long
Private blnGoingToReport As Boolean, blnWasGTR As Boolean
Private blnCurrID As Boolean, lngCurrID As Long
Private lngCusip_Offset As Long, blnRecalcTots As Boolean
Private strJType As String, dblPrice As Double, lngTaxcode As Long, datTransDate As Date, datAssetDate As Date
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
150           If CharCnt(strTmp01, "~") = 5 Then

160             blnIsOpen = True

                ' ** strCallingForm & "~" & strJType & "~" & CStr(dblPrice) & "~" & Format(datTransDate, "mm/dd/yyyy")
170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             strJType = Left(strTmp01, (intPos01 - 1))
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             dblPrice = Val(Left(strTmp01, (intPos01 - 1)))
250             strTmp01 = Mid(strTmp01, (intPos01 + 1))
260             intPos01 = InStr(strTmp01, "~")
270             datTransDate = CDate(Left(strTmp01, (intPos01 - 1)))
280             strTmp01 = Mid(strTmp01, (intPos01 + 1))
290             intPos01 = InStr(strTmp01, "~")
300             datAssetDate = CDate(Left(strTmp01, (intPos01 - 1)))
310             lngCurrID = Val(Mid(strTmp01, (intPos01 + 1)))

320             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
330             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

                ' ** Variables are fed empty, then populated ByRef.
340             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

350             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

360             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
370             If lngMonitorCnt > 1& Then
380               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
390               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
400               If lngMonitorNum > 1& Then
410                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
420               End If
430             End If

440             .form_width_line.Left = 0&  ' ** Just get it out of the way.
450             lngForm_Width = .Width
460             lngSub_Width = .frmMap_Reinvest_DivInt_Detail_Sub.Width
470             lngOK_Left = .cmdOK.Left
480             lngCancel_Left = .cmdCancel.Left
490             lngMapType_Left = .maptype.Left
500             lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

510             CLR_DISABLED_FG = CLR_DKGRY
520             CLR_DISABLED_BG = CLR_LTTEAL

530             blnCurrID = Pref_CurrID  ' ** Module Function: modPreferenceFuncs.
540             .curr_id = lngCurrID
550             IncludeCurrency_Sub  ' ** Procedure: Below.

560             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
570             blnWasGTR = False

580             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

590             chkShowLocation_AfterUpdate  ' ** Procedure: Below.

600             .Caption = "Map " & strJType & " Reinvest Detail"
610             lngCusip_Offset = (.cusip.Left - (.Header_lbl_Dividend.Left + .Header_lbl_Dividend.Width))

620             Set dbs = CurrentDb
630             With dbs
640               Select Case strCallingForm
                  Case "frmJournal"
                    ' ** Journal Map, linked to Account, grouped and summed by accountno.
650                 Set qdf = .QueryDefs("qryMap_Reinvest_02_04")
660               Case "frmJournal_Columns"
                    ' ** Journal Map, linked to Account, grouped and summed by accountno.
670                 Set qdf = .QueryDefs("qryJournal_Columns_30_08")
680               End Select
690               Set rst = qdf.OpenRecordset
700               With rst
710                 If .BOF = True And .EOF = True Then
720                   lngRecs = 0&
730                 Else
740                   .MoveLast
750                   lngRecs = .RecordCount
760                 End If
770                 .Close
780               End With
790               Set rst = Nothing
800               Set qdf = Nothing
810               .Close
820             End With
830             Set dbs = Nothing

840             Select Case strCallingForm
                Case "frmJournal"
850               Set dbs = CurrentDb
860               With dbs
                    ' ** Empty tblJournal_Map_Staging.
870                 Set qdf = .QueryDefs("qryMap_Reinvest_02_08")
880                 qdf.Execute
890                 Set qdf = Nothing
900                 DoEvents
                    ' ** Append qryMap_Reinvest_02_04 (Journal Map, linked to Account, qryMap_Reinvest_02_02
                    ' ** (qryMap_Reinvest_02_01 (Journal Map, grouped by accountno, assetno, Location_ID),
                    ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
                    ' ** grouped and summed by accountno) to tblJournal_Map_Staging.
910                 Set qdf = .QueryDefs("qryMap_Reinvest_02_09")
920                 qdf.Execute
930                 Set qdf = Nothing
940                 DoEvents
                    ' ** jmap_id in tblJournal_Map_Staging has Journal Map ID for reinvests,
                    ' ** and is the same as jmap_id2 in tblJournal_Map_Staging3.
                    ' **
                    ' ** Update qryMap_Reinvest_05_02 (tblJournal_Map_Staging, with DLookups() to
                    ' ** qryMap_Reinvest_05_01 (tblJournal_Map_Staging, linked to tblJournal_Map_Staging3,
                    ' ** with CheckNum_new (parent's Journal_ID))).
950                 Set qdf = .QueryDefs("qryMap_Reinvest_05_03")
960                 qdf.Execute
970                 Set qdf = Nothing
980                 .Close
990               End With
1000              Set dbs = Nothing
1010              DoEvents
1020              If .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecordSource <> "qryMap_Reinvest_02_10" Then  'qryMap_Reinvest_02_04
                    ' ** tblJournal_Map_Staging, all fields.
1030                .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecordSource = "qryMap_Reinvest_02_10"  'qryMap_Reinvest_02_04
1040              End If
1050            Case "frmJournal_Columns"
1060              Set dbs = CurrentDb
1070              With dbs
                    ' ** Empty tblJournal_Map_Staging.
1080                Set qdf = .QueryDefs("qryJournal_Columns_30_08a")
1090                qdf.Execute
1100                Set qdf = Nothing
1110                DoEvents
                    ' ** Append qryJournal_Columns_30_08 (Journal Map, linked to Account, qryJournal_Columns_30_05_02
                    ' ** (qryJournal_Columns_30_05_01 (Journal Map, grouped by accountno, assetno, Location_ID),
                    ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
                    ' ** grouped and summed by accountno) to tblJournal_Map_Staging.
1120                Set qdf = .QueryDefs("qryJournal_Columns_30_08b")
1130                qdf.Execute
1140                Set qdf = Nothing
1150                DoEvents
                    ' ** jmap_id in tblJournal_Map_Staging has Journal Map ID for reinvests,
                    ' ** and is the same as jmap_id2 in tblJournal_Map_Staging3.
                    ' **
                    ' ** Update qryJournal_Columns_30_12 (tblJournal_Map_Staging, with DLookups() to
                    ' ** qryJournal_Columns_30_11 (tblJournal_Map_Staging, linked to tblJournal_Map_Staging3,
                    ' ** with CheckNum_new (parent's JrnlCol_ID))).
1160                Set qdf = .QueryDefs("qryJournal_Columns_30_13")
1170                qdf.Execute
1180                Set qdf = Nothing
1190                .Close
1200              End With
1210              Set dbs = Nothing
1220              DoEvents
1230              If .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecordSource <> "qryJournal_Columns_30_08c" Then  'qryJournal_Columns_30_08
                    ' ** tblJournal_Map_Staging, all fields.
1240                .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecordSource = "qryJournal_Columns_30_08c"  'qryJournal_Columns_30_08
1250              End If
1260            End Select

1270            .maptype = "Purchase"
1280            .txtNumberRecords = "Number Records = " & CStr(lngRecs)
1290            .pershare = dblPrice
1300            .description = strJType & " Reinvestment"

1310            If .cmdOK.Enabled = True Then
1320              .cmdOK.SetFocus
1330            Else
1340              .description.SetFocus
1350            End If

1360            blnRecalcTots = True
1370            .TimerInterval = 500&

1380            DoEvents
1390            DoCmd.SelectObject acForm, THIS_NAME, False
1400            DoEvents

1410            DoCmd.Hourglass False

1420            blnIsOpen = False

1430          Else
1440            Cancel = -1
1450          End If
1460        Else
1470          Cancel = -1
1480        End If
1490      Else
1500        Cancel = -1
1510      End If
1520    End With

1530    If Cancel = -1 Then
1540      Beep
1550      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1560    End If

EXITP:
1570    Set rst = Nothing
1580    Set qdf = Nothing
1590    Set dbs = Nothing
1600    Exit Sub

ERRH:
1610    DoCmd.Hourglass False
1620    Select Case ERR.Number
        Case Else
1630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1640    End Select
1650    Resume EXITP

End Sub

Private Sub Form_Load()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1710    With Me

1720      Select Case strJType
          Case "Dividend"
1730        gstrCrtRpt_CashAssets_Beg = strJType
1740        gstrCrtRpt_CashAssets_End = "Shares"
1750        .Header_lbl_Dividend.Visible = True
1760        .Header_lbl_Interest.Visible = False
1770        .frmMap_Reinvest_DivInt_Detail_Sub.Form.total_shareface_lbl.Caption = gstrCrtRpt_CashAssets_End
1780        If blnCurrID = True And lngCurrID <> 150& Then
1790          .pershare_lbl.Visible = False
1800        Else
1810          .pershare_lbl.Visible = True
1820        End If
1830        .pershare_lbl1.Visible = False
            ' ** Borrowing these variables from the Court Reports for this report, below.
1840      Case "Interest"
1850        gstrCrtRpt_CashAssets_Beg = strJType
1860        gstrCrtRpt_CashAssets_End = "Units"
1870        .Header_lbl_Dividend.Visible = False
1880        .Header_lbl_Interest.Visible = True
1890        .cusip.Left = ((.Header_lbl_Interest.Left + .Header_lbl_Interest.Width) + lngCusip_Offset)
1900        .totdesc.Left = .cusip.Left
1910        .frmMap_Reinvest_DivInt_Detail_Sub.Form.total_shareface_lbl.Caption = gstrCrtRpt_CashAssets_End
1920        If blnCurrID = True And lngCurrID <> 150& Then
1930          .pershare_lbl1.Visible = False
1940        Else
1950          .pershare_lbl1.Visible = True
1960        End If
1970        .pershare_lbl.Visible = False
            ' ** Borrowing these variables from the Court Reports for this report, below.
1980      End Select

1990      .total_shareface_sum.ForeColor = CLR_DISABLED_FG
2000      .total_shareface_sum.BackColor = CLR_DISABLED_BG
2010      .total_cash_sum.ForeColor = CLR_DISABLED_FG
2020      .total_cash_sum.BackColor = CLR_DISABLED_BG
2030      .total_cash_sum_usd.ForeColor = CLR_DISABLED_FG
2040      .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
2050      .pershare.ForeColor = CLR_DISABLED_FG
2060      .pershare.BackColor = CLR_DISABLED_BG
2070      .pershare_usd.ForeColor = CLR_DISABLED_FG
2080      .pershare_usd.BackColor = CLR_DISABLED_BG

          ' ** Borrowing these variables from the Court Reports.
2090      If gstrCrtRpt_Version <> vbNullString Then
2100        .cusip = gstrCrtRpt_Ordinal
2110        .cusip.ForeColor = CLR_DISABLED_FG
2120        .cusip.BackColor = CLR_DISABLED_BG
2130        .cusip.Visible = True
2140        .totdesc = gstrCrtRpt_Version
2150        .totdesc.ForeColor = CLR_DISABLED_FG
2160        .totdesc.BackColor = CLR_DISABLED_BG
2170        .totdesc.Visible = True
2180      End If

2190      .assetdate = datAssetDate
2200      .transdate = datTransDate

2210      Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
2220        .cmbRevenueCodes_Purchase = REVID_INC
2230        .cmbRevenueCodes_Purchase_display = Null
2240      Case False
2250        If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
2260          .cmbRevenueCodes_Purchase_display = Null
2270        Else
2280          .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
2290        End If
2300      End Select

2310      Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
2320        .cmbTaxCodes_Purchase = TAXID_INC
2330        .cmbTaxCodes_Purchase_display = Null
2340      Case False
2350        If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
2360          .cmbTaxCodes_Purchase_display = Null
2370        Else
2380          .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
2390        End If
2400      End Select

2410      Select Case gblnRevenueExpenseTracking
          Case True
2420        .cmbRevenueCodes_Purchase.Visible = True
2430        .cmbRevenueCodes_Purchase_display.Visible = True
2440        .cmbRevenueCodes_Purchase_lbl2.Visible = False  ' ** Option off msg.
2450      Case False
2460        .cmbRevenueCodes_Purchase.Visible = False
2470        .cmbRevenueCodes_Purchase_display.Visible = False
2480        .cmbRevenueCodes_Purchase_lbl_box.Visible = False
2490        .cmbRevenueCodes_Purchase_lbl2.Visible = True  ' ** Option off msg.
2500      End Select

2510      Select Case gblnIncomeTaxCoding
          Case True
2520        .cmbTaxCodes_Purchase.Visible = True
2530        .cmbTaxCodes_Purchase_display.Visible = True
2540        .cmbTaxCodes_Purchase_lbl2.Visible = False  ' ** Option off msg.
2550      Case False
2560        .cmbTaxCodes_Purchase.Visible = False
2570        .cmbTaxCodes_Purchase_display.Visible = False
2580        .cmbTaxCodes_Purchase_lbl_box.Visible = False  ' ** Option off msg.
2590        .cmbTaxCodes_Purchase_lbl2.Visible = True
2600      End Select

2610      If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
2620        .cmbRevenueCodes_Purchase_box_lbl.ForeColor = WIN_CLR_DISF
2630        .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Visible = True
2640      End If

2650      curr_id_AfterUpdate  ' ** Procedure: Below.
2660      IncludeCurrency False  ' ** Procedure: Below
2670      DoEvents

2680      .frmMap_Reinvest_DivInt_Detail_Sub.Form.Requery
2690      DoEvents
2700      .frmMap_Reinvest_DivInt_Detail_Sub.Form.TimerInterval = 1000&

2710    End With

EXITP:
2720    Exit Sub

ERRH:
2730    DoCmd.Hourglass False
2740    Select Case ERR.Number
        Case Else
2750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2760    End Select
2770    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Comments:         M {description}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc Totals:    F5 {RecalcTots}

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyEscape
2870        intRetVal = 0
2880        cmdCancel_Click  ' ** Procedure: Below.
2890      Case vbKeyF5
2900        intRetVal = 0
2910        RecalcTots  ' ** Function: Below.
2920      End Select
2930    End If

        ' ** Ctrl keys.
2940    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2950      Select Case intRetVal
          Case vbKeyM
2960        With Me
2970          intRetVal = 0
2980          .description.SetFocus
2990        End With
3000      End Select
3010    End If

        ' ** Ctrl-Shift keys.
3020    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3030      Select Case intRetVal
          Case vbKeyF
3040        With Me
3050          intRetVal = 0
3060          .FocusHolder.SetFocus
3070        End With
3080      End Select
3090    End If

EXITP:
3100    KeyCode = intRetVal
3110    Exit Sub

ERRH:
3120    intRetVal = 0
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub Form_Timer()

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim blnRetVal As Boolean

3210    With Me
3220      .TimerInterval = 0&
3230      If gblnGoToReport = True Then
3240        Select Case blnGoingToReport
            Case True
3250          DoCmd.Hourglass False
3260          gblnGoToReport = False
3270          blnGoingToReport = False
3280          .GoToReport_arw_pr_img.Visible = False
3290          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
3300          blnWasGTR = True
3310          ReportList_DelJournal  ' ** Module Function: modReportFunctions.
3320          .cmdCancel.SetFocus
3330        Case False
3340          DoCmd.Hourglass True  ' ** Make sure it's still running.
3350          DoEvents
3360          blnGoingToReport = True
3370          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3380          .GoToReport_arw_pr_img.Visible = True
3390          If .cmdPrintReport.Enabled = True Then
3400            .cmdPrintReport.SetFocus
3410          End If
3420          DoEvents
3430          Beep
3440          DoCmd.Hourglass False
3450          DoEvents
3460          .TimerInterval = (GTR_WAIT * 2&)
3470        End Select
3480      End If
3490      If blnRecalcTots = True Then
3500        blnRecalcTots = False
3510        blnRetVal = RecalcTots  ' ** Function: Below.
3520        If blnRetVal = False Then
3530          blnRecalcTots = True
3540          .TimerInterval = 500&
3550        End If
3560      End If
3570    End With

EXITP:
3580    Exit Sub

ERRH:
3590    DoCmd.Hourglass False
3600    Select Case ERR.Number
        Case Else
3610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3620    End Select
3630    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3710    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3720    Set dbs = CurrentDb
        ' ** Empty tmpAccount.
3730    Set qdf = dbs.QueryDefs("qryMap_Div_08_01")
3740    qdf.Execute
3750    Set qdf = Nothing
3760    DoEvents
3770    dbs.Close
3780    Set dbs = Nothing
3790    DoEvents

3800    gstrCrtRpt_CashAssets_Beg = vbNullString
3810    gstrCrtRpt_CashAssets_End = vbNullString
3820    gstrReportCallingForm = vbNullString
3830    gstrReportQuerySpec = vbNullString

3840    If gstrReturningForm = vbNullString Then
3850      gstrReturningForm = THIS_NAME
3860    End If

EXITP:
3870    Set qdf = Nothing
3880    Set dbs = Nothing
3890    Exit Sub

ERRH:
3900    Select Case ERR.Number
        Case Else
3910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3920    End Select
3930    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4010    With Me
4020      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
4030        Select Case blnPrintReport_Focus
            Case True
4040          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
4050          .cmdPrintReport_raised_img.Visible = False
4060        Case False
4070          .cmdPrintReport_raised_img.Visible = True
4080          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4090        End Select
4100        .cmdPrintReport_raised_focus_img.Visible = False
4110        .cmdPrintReport_raised_focus_dots_img.Visible = False
4120        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4130        .cmdPrintReport_raised_img_dis.Visible = False
4140      End If
4150    End With

EXITP:
4160    Exit Sub

ERRH:
4170    Select Case ERR.Number
        Case Else
4180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4190    End Select
4200    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4320    Exit Sub

ERRH:
4330    Select Case ERR.Number
        Case Else
4340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4350    End Select
4360    Resume EXITP

End Sub

Private Sub cmdOK_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngRevCode As Long, lngRevType As Long, lngTaxType As Long
        Dim strRevDesc As String, strTaxDesc As String
        Dim lngAssetNo As Long
        Dim varComments As Variant
        Dim blnLocChanged As Boolean, blnSkip As Boolean
        Dim varTmp00 As Variant

4410    With Me

4420      lngRecsCur = .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecCnt
4430      If lngRecsCur > 0& Then

4440        DoCmd.Hourglass True
4450        DoEvents

4460        blnLocChanged = False

            ' ** EVERY RECORD HAS A DIFFERENT ACCOUNTNO!
            'strAccountNo = .frmMap_Reinvest_DivInt_Detail_Sub.Form.accountno
4470        lngAssetNo = .frmMap_Reinvest_DivInt_Detail_Sub.Form.assetno
4480        lngRevCode = .cmbRevenueCodes_Purchase
4490        strRevDesc = .cmbRevenueCodes_Purchase.Column(1)
4500        lngRevType = .cmbRevenueCodes_Purchase.Column(2)
4510        lngTaxcode = .cmbTaxCodes_Purchase
4520        strTaxDesc = .cmbTaxCodes_Purchase.Column(1)
4530        lngTaxType = .cmbTaxCodes_Purchase.Column(2)
4540        varComments = .description

4550        Set dbs = CurrentDb
4560        With dbs

              ' ** OK, HOW DO I GET ALL THE accountno'S, AND HOW DO I KNOW WHICH DIVIDEND ENTRIES THEY REFLECT?

              ' ** Update qryMap_Reinvest_02_11 (tblJournal_Map_Staging, linked
              ' ** to Journal, for icash, by specified [astno]), for Reinvested = True.
4570          Set qdf = .QueryDefs("qryMap_Reinvest_02_12")
4580          With qdf.Parameters
4590            ![astno] = lngAssetNo
4600          End With
4610          qdf.Execute
4620          Set qdf = Nothing
4630          DoEvents

4640          Select Case strCallingForm
              Case "frmJournal"

                ' ** Max(ID) prior to appending Dividends has been saved to tblMark.
                'varTmp00 = DLookup("[unique_id]", "tblMark")
                'Select Case IsNull(varTmp00)
                'Case True  ' ** Journal was empty.
                '  ' ** Update qryMap_Reinvest_03m (Journal, just 'Dividend' or 'Interest',
                '  ' ** by specified [actno], [astno], [jtyp]), for Reinvested = True.
                '  Set qdf = .QueryDefs("qryMap_Reinvest_03n")
                '  With qdf.Parameters
                '    ![actno] = strAccountNo  ' ** Because Journal is in TrustDta.mdb.
                '    ![astno] = lngAssetNo
                '    ![jtyp] = strJType
                '  End With
                '  qdf.Execute
                'Case False
                '  ' ** Update qryMap_Reinvest_03o (Journal, just 'Dividend' or 'Interest',
                '  ' ** by specified [actno], [astno], [jtyp], [maxid]), for Reinvested = True.
                '  Set qdf = .QueryDefs("qryMap_Reinvest_03p")
                '  With qdf.Parameters
                '    ![actno] = strAccountNo  ' ** Because Journal is in TrustDta.mdb.
                '    ![astno] = lngAssetNo
                '    ![jtyp] = strJType
                '    ![maxid] = CLng(varTmp00)
                '  End With
                '  qdf.Execute
                'End Select

                'WHY IS THIS LOOKING FOR PURCHASE'S IN THE JOURNAL
                'IF THESE REINVEST'S ARE ONLY IN THE TMP TABLE?
                ' ** Delete Journal, just 'Purchase', for shareface < 0.0001, cost < 0.01.
4650            Set qdf = .QueryDefs("qryMap_Reinvest_04")
4660            qdf.Execute
4670            Set qdf = Nothing
4680            DoEvents

                ' ** qryMap_Reinvest_03a (tblJournal_Map_Staging, just loc_changed = True), grouped, with cnt.
4690            varTmp00 = DLookup("[cnt]", "qryMap_Reinvest_03b")
4700            If IsNull(varTmp00) = False Then
4710              If varTmp00 > 0 Then
4720                blnLocChanged = True
4730              End If
4740            End If
4750            If blnLocChanged = True Then
                  ' ** Update qryMap_Reinvest_03c (Journal Map, linked to qryMap_Reinvest_03a
                  ' ** (tblJournal_Map_Staging, just loc_changed = True), with Location_ID_new).
4760              Set qdf = .QueryDefs("qryMap_Reinvest_03d")
4770              qdf.Execute
4780              Set qdf = Nothing
4790            End If
4800            DoEvents

                ' ** Append Journal Map, linked to tblJournal_Map_Staging, to Journal,
                ' ** by specified [pric], [astdat], [desc], [revcod], [taxcod].  #curr_id
4810            Set qdf = .QueryDefs("qryMap_Reinvest_03")
                ' ** The reinvests should already be coming over with their parent's Journal_ID in CheckNum.

4820          Case "frmJournal_Columns"

                ' ** Max(JrnlCol_ID) prior to appending Dividends has been saved to tblMark.
                'varTmp00 = DLookup("[unique_id]", "tblMark")
                'Select Case IsNull(varTmp00)
                'Case True  ' ** Journal was empty.
                '  ' ** Update qryJournal_Columns_30_09m (tblJournal_Column, just 'Dividend' or 'Interest',
                '  ' ** by specified [jtyp]), for Reinvested = True.
                '  Set qdf = .QueryDefs("qryJournal_Columns_30_09n")
                '  With qdf.Parameters
                '    ![jtyp] = strJType  ' ** Because this is local, accountno and assetno aren't needed.
                '  End With
                '  qdf.Execute
                'Case False
                '  ' ** Update qryJournal_Columns_30_09o (tblJournal_Column, just 'Dividend' or 'Interest',
                '  ' ** by specified [jtyp], [maxid]), for Reinvested = True.
                '  Set qdf = .QueryDefs("qryJournal_Columns_30_09p")
                '  With qdf.Parameters
                '    ![jtyp] = strJType  ' ** Because this is local, accountno and assetno aren't needed.
                '    ![maxid] = CLng(varTmp00)
                '  End With
                '  qdf.Execute
                'End Select

                ' ** Delete Journal, just 'Purchase', for shareface < 0.0001, cost < 0.01.
4830            Set qdf = .QueryDefs("qryJournal_Columns_30_10")
4840            qdf.Execute
4850            Set qdf = Nothing
4860            DoEvents

                ' ** qryJournal_Columns_30_09a (tblJournal_Map_Staging, just loc_changed = True), grouped, with cnt.
4870            varTmp00 = DLookup("[cnt]", "qryJournal_Columns_30_09b")
4880            If IsNull(varTmp00) = False Then
4890              If varTmp00 > 0 Then
4900                blnLocChanged = True
4910              End If
4920            End If
4930            If blnLocChanged = True Then
                  ' ** Update qryJournal_Columns_30_09c (Journal Map, linked to qryJournal_Columns_30_09a
                  ' ** (tblJournal_Map_Staging, just loc_changed = True), with Location_ID_new).
4940              Set qdf = .QueryDefs("qryJournal_Columns_30_09d")
4950              qdf.Execute
4960              Set qdf = Nothing
4970            End If
4980            DoEvents

                ' ** Append Journal Map, linked to tblJournal_Map_Staging, to tblJournal_Column, by specified
                ' ** [pric], [astdat], [desc], [revcod], [taxcod], [revdesc], [taxdesc], [revtyp], [taxtyp].
4990            Set qdf = .QueryDefs("qryJournal_Columns_30_09")
5000            With qdf.Parameters
5010              ![revdesc] = strRevDesc
5020              ![taxdesc] = strTaxDesc
5030              ![revtyp] = lngRevType
5040              ![taxtyp] = lngTaxType
5050            End With
                ' ** The reinvests should already be coming over with their parent's JrnlCol_ID in CheckNum.

5060          End Select
5070          With qdf.Parameters
5080            ![pric] = CCur(dblPrice)
5090            ![astdat] = datAssetDate  ' ** Query adds Time().
5100            ![Desc] = varComments
5110            ![revcod] = lngRevCode
5120            ![taxcod] = lngTaxcode
5130          End With
5140          qdf.Execute dbFailOnError
5150          Set qdf = Nothing
5160          DoEvents

              ' ** Use tblJournal_Map_Staging3 to populate tblJournal_MiscSold.
5170          Select Case strCallingForm
              Case "frmJournal"
                ' ** Journal_ID in .._Staging3 is parents' Journal ID.
                ' ** CheckNum in parents' Journal is children's .._Staging3 jmap_id1.
                ' ** CheckNum in children's Journal is parents' Journal ID.
                ' **
                ' ** Append tblJournal_Map_Staging3, linked to Journal, to
                ' ** tblJournal_MiscSold, by specified [jtyp], for parent record.
5180            Set qdf = .QueryDefs("qryMap_Reinvest_07_01")
5190            With qdf.Parameters
5200              ![jtyp] = strJType
5210            End With
5220            qdf.Execute
5230            Set qdf = Nothing
5240            DoEvents
                ' ** Append tblJournal_Map_Staging3, linked to Journal, to
                ' ** tblJournal_MiscSold, by specified [jtyp], for child record.
5250            Set qdf = .QueryDefs("qryMap_Reinvest_07_02")
5260            With qdf.Parameters
5270              ![jtyp] = strJType
5280            End With
5290            qdf.Execute
5300            Set qdf = Nothing
5310          Case "frmJournal_Columns"
                ' ** JrnlCol_ID in .._Staging3 is parents' JrnlCol_ID.
                ' ** CheckNum in parents' tblJournal_Column is children's .._Staging3 jmap_id1.
                ' ** CheckNum in children's tblJournal_Column is parents' JrnlCol_ID.
                ' **
                ' ** Empty tblJournal_MiscSold_Staging.
5320            Set qdf = .QueryDefs("qryJournal_Columns_32_06")
5330            qdf.Execute
5340            Set qdf = Nothing
5350            DoEvents
                ' ** Append tblJournal_Map_Staging3, linked to tblJournal_Column, to
                ' ** tblJournal_MiscSold_Staging, by specified [jtyp], for parent record.
5360            Set qdf = .QueryDefs("qryJournal_Columns_32_04")
5370            With qdf.Parameters
5380              ![jtyp] = strJType
5390            End With
5400            qdf.Execute
5410            Set qdf = Nothing
5420            DoEvents
                ' ** Append tblJournal_Map_Staging3, linked to tblJournal_Column, to
                ' ** tblJournal_MiscSold_Staging, by specified [jtyp], for child record.
5430            Set qdf = .QueryDefs("qryJournal_Columns_32_05")
5440            With qdf.Parameters
5450              ![jtyp] = strJType
5460            End With
5470            qdf.Execute
5480            Set qdf = Nothing
5490          End Select
5500          DoEvents

5510          blnSkip = False
5520          If blnSkip = False Then
5530            Select Case strCallingForm
                Case "frmJournal"
                  ' ** Empty Journal Map.
5540              Set qdf = .QueryDefs("qryMap_01")
5550              qdf.Execute
5560              Set qdf = Nothing
5570              DoEvents
                  ' ** Empty tblJournal_Map_Staging.
5580              Set qdf = .QueryDefs("qryMap_Reinvest_02_08")
5590              qdf.Execute
5600              Set qdf = Nothing
5610              DoEvents
                  ' ** Empty tblJournal_Map_Staging3.
5620              Set qdf = .QueryDefs("qryMap_Reinvest_05_07")
5630              qdf.Execute
5640              Set qdf = Nothing
5650            Case "frmJournal_Columns"
                  ' ** Empty Journal Map.
5660              Set qdf = .QueryDefs("qryJournal_Columns_30_04")
5670              qdf.Execute
5680              Set qdf = Nothing
5690              DoEvents
                  ' ** Empty tblJournal_Map_Staging.
5700              Set qdf = .QueryDefs("qryJournal_Columns_30_08a")
5710              qdf.Execute
5720              Set qdf = Nothing
5730              DoEvents
                  'WE CAN'T EMPTY THIS HERE! IT'S USED IN JC_Frm_Map_Return!
                  ' ** Empty tblJournal_Map_Staging3.
                  'Set qdf = .QueryDefs("qryJournal_Columns_32_03")
                  'qdf.Execute
                  'Set qdf = Nothing
5740            End Select
5750          End If  ' ** blnSkip.
5760          DoEvents

5770          .Close
5780        End With
5790        Set dbs = Nothing
5800        DoEvents

5810        Select Case strCallingForm
            Case "frmJournal"
5820          If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5830            DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
5840          Else
5850            DoCmd.SelectObject acForm, strCallingForm, False
5860          End If
5870          Select Case strJType
              Case "Dividend"
5880            Forms(strCallingForm).frmJournal_Sub1_Dividend.Form.Requery
5890            DoEvents
5900            Forms(strCallingForm).frmJournal_Sub1_Dividend.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub1_Dividend.
5910          Case "Interest"
5920            Forms(strCallingForm).frmJournal_Sub2_Interest.Form.Requery
5930            DoEvents
5940            Forms(strCallingForm).frmJournal_Sub2_Interest.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub2_Interest.
5950          End Select
5960        Case "frmJournal_Columns"
5970          gblnSetFocus = True
5980          gstrReturningForm = "frmMap_Return"
5990          If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
6000            DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
6010          Else
6020            Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
6030            Forms(strCallingForm).frmJournal_Columns_Sub.Form.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
6040            DoCmd.SelectObject acForm, strCallingForm, False
6050            Forms(strCallingForm).TimerInterval = 250&
6060          End If
6070        End Select

6080        DoCmd.Close acForm, THIS_NAME

6090      Else
6100        Beep
6110      End If  ' ** lngRecsCur.

6120      DoCmd.Hourglass False

6130    End With

EXITP:
6140    Set qdf = Nothing
6150    Set dbs = Nothing
6160    Exit Sub

ERRH:
6170    DoCmd.Hourglass False
6180    Select Case ERR.Number
        Case Else
6190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6200    End Select
6210    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6320    intShiftDown = (Shift And acShiftMask) > 0
6330    intAltDown = (Shift And acAltMask) > 0
6340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6360      Select Case intRetVal
          Case vbKeyTab
6370        With Me
6380          intRetVal = 0
6390          .cmdCancel.SetFocus
6400        End With
6410      End Select
6420    End If

        ' ** Shift keys.
6430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6440      Select Case intRetVal
          Case vbKeyTab
6450        With Me
6460          intRetVal = 0
6470          .cmdPrintReport.SetFocus
6480        End With
6490      End Select
6500    End If

EXITP:
6510    KeyCode = intRetVal
6520    Exit Sub

ERRH:
6530    intRetVal = 0
6540    Select Case ERR.Number
        Case Else
6550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6560    End Select
6570    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intTmp01 As Integer
        Dim intX As Integer

6610    DoCmd.Hourglass True
6620    DoEvents

6630    intTmp01 = 0
6640    Do While Forms.Count > 1
6650      intTmp01 = intTmp01 + 1
6660      For intX = (Forms.Count - 1) To 0 Step -1
6670        Select Case Forms(intX).Name
            Case THIS_NAME
              ' ** Skip.
6680        Case "frmJournal", "frmJournal_Columns"
              ' ** Skip.
6690        Case Else
6700          DoCmd.Close acForm, Forms(intX).Name
6710          Exit For  ' ** Do them one-at-a-time.
6720        End Select
6730      Next
6740      If intTmp01 >= 10 Then Exit Do  ' ** Don't get caught in a loop.
6750    Loop

6760    Set dbs = CurrentDb
6770    With dbs

          ' ** See that posted and reinvested are turned off,
          ' ** and Null-out CheckNum on parent transactions.
6780      Select Case strCallingForm
          Case "frmJournal"
            ' ** Update qryMap_Reinvest_06_01 (tblJournal_Map_Staging3, linked to
            ' ** Journal, with posted_new, Reinvested_new, Check_Num_new).
6790        Set qdf = .QueryDefs("qryMap_Reinvest_06_02")
6800        qdf.Execute
6810        Set qdf = Nothing
6820        DoEvents
            ' ** Empty Journal Map.
6830        Set qdf = .QueryDefs("qryMap_01")
6840        qdf.Execute
6850        Set qdf = Nothing
6860        DoEvents
            ' ** Empty tblJournal_Map_Staging.
6870        Set qdf = .QueryDefs("qryMap_Reinvest_02_08")
6880        qdf.Execute
6890        Set qdf = Nothing
6900        DoEvents
            ' ** Empty tblJournal_Map_Staging3
6910        Set qdf = .QueryDefs("qryMap_Reinvest_05_07")
6920        qdf.Execute
6930        Set qdf = Nothing
6940      Case "frmJournal_Columns"
            ' ** Update qryJournal_Columns_32_01 (tblJournal_Map_Staging3, linked to
            ' ** tblJournal_Column, with posted_new, Reinvested_new, Check_Num_new).
6950        Set qdf = .QueryDefs("qryJournal_Columns_32_02")
6960        qdf.Execute
6970        Set qdf = Nothing
6980        DoEvents
            ' ** Empty Journal Map.
6990        Set qdf = .QueryDefs("qryJournal_Columns_30_04")
7000        qdf.Execute
7010        Set qdf = Nothing
7020        DoEvents
            ' ** Empty tblJournal_Map_Staging.
7030        Set qdf = .QueryDefs("qryJournal_Columns_30_08a")
7040        qdf.Execute
7050        Set qdf = Nothing
7060        DoEvents
            ' ** Empty tblJournal_Map_Staging3.
7070        Set qdf = .QueryDefs("qryJournal_Columns_32_03")
7080        qdf.Execute
7090        Set qdf = Nothing
7100      End Select
7110      DoEvents
7120      .Close
7130    End With
7140    Set dbs = Nothing
7150    DoEvents

7160    gblnSetFocus = True
7170    Select Case strCallingForm
        Case "frmJournal"
7180      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
7190        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
7200      Else
7210        gstrReturningForm = THIS_NAME
7220        Select Case strJType
            Case "Dividend"
7230          Forms(strCallingForm).frmJournal_Sub1_Dividend.Form.Requery
7240          DoEvents
7250          Forms(strCallingForm).frmJournal_Sub1_Dividend.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub1_Dividend.
7260        Case "Interest"
7270          Forms(strCallingForm).frmJournal_Sub2_Interest.Form.Requery
7280          DoEvents
7290          If blnWasGTR = False Then
7300            Forms(strCallingForm).frmJournal_Sub2_Interest.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub2_Interest.
7310          End If
7320        End Select
7330        DoCmd.SelectObject acForm, strCallingForm, False
7340        Forms(strCallingForm).TimerInterval = 250&
7350      End If
7360    Case "frmJournal_Columns"
7370      gstrReturningForm = "frmMap_Return"
7380      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
7390        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
7400      Else
7410        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
7420        DoCmd.SelectObject acForm, strCallingForm, False
7430        Forms(strCallingForm).TimerInterval = 250&
7440      End If
7450    End Select

7460    DoCmd.Close acForm, THIS_NAME

EXITP:
7470    Set qdf = Nothing
7480    Set dbs = Nothing
7490    Exit Sub

ERRH:
7500    DoCmd.Hourglass False
7510    Select Case ERR.Number
        Case Else
7520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7530    End Select
7540    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7620    intShiftDown = (Shift And acShiftMask) > 0
7630    intAltDown = (Shift And acAltMask) > 0
7640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7660      Select Case intRetVal
          Case vbKeyTab
7670        With Me
7680          intRetVal = 0
7690          lngRecsCur = .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Reinvest_DivInt_Detail_Sub.
7700          If lngRecsCur > 0 Then
7710            .frmMap_Reinvest_DivInt_Detail_Sub.SetFocus
7720            .frmMap_Reinvest_DivInt_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMap_Reinvest_DivInt_Detail_Sub.
7730            .frmMap_Reinvest_DivInt_Detail_Sub.Form.FocusHolder.SetFocus
7740          Else
7750            .description.SetFocus
7760          End If
7770        End With
7780      End Select
7790    End If

        ' ** Shift keys.
7800    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7810      Select Case intRetVal
          Case vbKeyTab
7820        With Me
7830          intRetVal = 0
7840          .cmdOK.SetFocus
7850        End With
7860      End Select
7870    End If

EXITP:
7880    KeyCode = intRetVal
7890    Exit Sub

ERRH:
7900    intRetVal = 0
7910    Select Case ERR.Number
        Case Else
7920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7930    End Select
7940    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

8010    With Me
8020      blnPrintReport_Focus = True
8030      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
8040      .cmdPrintReport_raised_img.Visible = False
8050      .cmdPrintReport_raised_focus_img.Visible = False
8060      .cmdPrintReport_raised_focus_dots_img.Visible = False
8070      .cmdPrintReport_sunken_focus_dots_img.Visible = False
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

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

8210    With Me
8220      blnPrintReport_MouseDown = True
8230      .cmdPrintReport_sunken_focus_dots_img.Visible = True
8240      .cmdPrintReport_raised_img.Visible = False
8250      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8260      .cmdPrintReport_raised_focus_img.Visible = False
8270      .cmdPrintReport_raised_focus_dots_img.Visible = False
8280      .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_Click()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

        ' ** tblJournal, by specified FormRef('Price').
        ' ** qryMapReport_05.
        'strSQL = "SELECT DISTINCTROW account.accountno, account.shortname, Sum([journal map].icash/" & strPrice & ") AS total_shareface, " & _
        '  "Sum([journal map].icash) AS total_icash, " & CoInfo & ", " & _
        '  "CStr([masterasset].[Description]) & IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy')) AS totdesc " & _
        '  "FROM ([journal map] INNER JOIN account ON [journal map].accountno = account.accountno) " & _
        '  "INNER JOIN masterasset ON [journal map].assetno = masterasset.assetno " & _
        '  "GROUP BY account.accountno, account.shortname, " & _
        '  "CStr([masterasset].[Description]) & IIf([masterasset].[rate]>0,' ' & " & _
        '  "Format([masterasset].[rate],'0.000%')) & IIf([masterasset].[due] Is Not Null,'  Due ' & " & _
        '  "Format([masterasset].[due],'mm/dd/yyyy'));"

8410    With Me
          ' ** If the report is sent to a file, it takes the name of the report caption.
8420      Select Case strJType
          Case "Dividend"
8430        strDocName = "rptMap_Reinvest_Div"
8440      Case "Interest"
8450        strDocName = "rptMap_Reinvest_Int"
8460      End Select
8470      gstrReportCallingForm = THIS_NAME
8480      gstrReportQuerySpec = strDocName
8490      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
8500        .Modal = False
8510        DoCmd.OpenReport strDocName, acViewPreview
8520        DoCmd.Maximize
8530        DoCmd.RunCommand acCmdFitToWindow
8540        .Visible = False
8550      Else
            '##GTR_Ref: rptMap_Reinvest_Div
            '##GTR_Ref: rptMap_Reinvest_Int
8560        DoCmd.OpenReport strDocName, acViewNormal
8570      End If
8580    End With

EXITP:
8590    Exit Sub

ERRH:
8600    Me.Visible = True
8610    DoCmd.Restore
8620    Select Case ERR.Number
        Case Else
8630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8640    End Select
8650    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8720    intShiftDown = (Shift And acShiftMask) > 0
8730    intAltDown = (Shift And acAltMask) > 0
8740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8760      Select Case intRetVal
          Case vbKeyTab
8770        With Me
8780          intRetVal = 0
8790          .cmdOK.SetFocus
8800        End With
8810      End Select
8820    End If

        ' ** Shift keys.
8830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8840      Select Case intRetVal
          Case vbKeyTab
8850        With Me
8860          intRetVal = 0
8870          .chkShowLocation.SetFocus
8880        End With
8890      End Select
8900    End If

EXITP:
8910    KeyCode = intRetVal
8920    Exit Sub

ERRH:
8930    intRetVal = 0
8940    Select Case ERR.Number
        Case Else
8950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8960    End Select
8970    Resume EXITP

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

Private Sub total_cash_sum_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "total_cash_sum_KeyDown"

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
9870          lngRecsCur = .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Reinvest_DivInt_Detail_Sub.
9880          If lngRecsCur > 0& Then
9890            .frmMap_Reinvest_DivInt_Detail_Sub.SetFocus
9900            .frmMap_Reinvest_DivInt_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Reinvest_DivInt_Detail_Sub.
9910            .frmMap_Reinvest_DivInt_Detail_Sub.Form.FocusHolder.SetFocus
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

Private Sub total_shareface_sum_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_sum_KeyDown"

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
10270         lngRecsCur = .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Reinvest_DivInt_Detail_Sub.
10280         If lngRecsCur > 0& Then
10290           .frmMap_Reinvest_DivInt_Detail_Sub.SetFocus
10300           .frmMap_Reinvest_DivInt_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Reinvest_DivInt_Detail_Sub.
10310           .frmMap_Reinvest_DivInt_Detail_Sub.Form.FocusHolder.SetFocus
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

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

10510   With Me
10520     Select Case .curr_date.Visible
          Case True
10530       .curr_date.Visible = False
10540     Case False
10550       .curr_date = .curr_id.Column(CBX_C_DATE)
10560       .curr_date.Visible = True
10570     End Select
10580   End With

EXITP:
10590   Exit Sub

ERRH:
10600   Select Case ERR.Number
        Case Else
10610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
10620   End Select
10630   Resume EXITP

End Sub

Private Sub curr_id_AfterUpdate()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_AfterUpdate"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String
        Dim lngX As Long

10710   With Me

10720     If IsNull(.curr_id) = False Then

10730       lngCurrID = .curr_id.Column(CBX_C_CURRID)

10740       blnFound = False
10750       If lngCurrID <> 150& Then
10760         If lngNoChars = 0& Or IsEmpty(arr_varNoChar) = True Then
10770           NoChar_Load  ' ** Procedure: Below.
10780         End If
10790         For lngX = 0& To (lngNoChars - 1&)
10800           If arr_varNoChar(N_CID, lngX) = lngCurrID Then
                  ' ** This currency's symbol does not show in the Fixedsys font.
10810             blnFound = True
10820           End If
10830         Next
10840       End If

10850       Select Case blnFound
            Case True
              ' ** Fixed Width Fonts:
              ' **   Fixedsys
              ' **   Miriam Fixed
              ' **   Courier New
10860         .total_cash_sum.FontName = "Courier New"  ' ** Looks a little better than Miriam for our purposes.
10870         .total_cash_sum.FontSize = 10
10880         .total_cash_sum.FontBold = True
10890         .pershare.FontName = "Courier New"
10900         .pershare.FontSize = 10
10910         .pershare.FontBold = True
10920       Case False
10930         If .total_cash_sum.FontName <> "Arial" Then
10940           .total_cash_sum.FontName = "Arial"
10950           .total_cash_sum.FontSize = 10
10960           .total_cash_sum.FontBold = False
10970           .pershare.FontName = "Arial"
10980           .pershare.FontSize = 10
10990           .pershare.FontBold = False
11000         End If
11010       End Select

11020       If lngCurrID = 150& Then  ' ** USD.
11030         strFrmt1 = "Currency"
11040         intDec1 = 2
11050         strFrmt2 = "Currency"
11060         intDec2 = 5
11070         lngBkClr = CLR_DISABLED_BG
11080       Else
11090         strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
11100         If .curr_id.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
11110           intDec1 = 2
11120         Else
11130           intDec1 = .curr_id.Column(CBX_C_DEC)  ' ** Decimal places.
11140         End If
11150         intDec2 = 5
11160         intLen = Len(strTmp01)
11170         Select Case intLen
              Case 1
                'strTmp01 = strTmp01
11180         Case 2
11190           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
11200         Case 3
11210           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
11220         Case 4
11230           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
11240         Case 5
11250           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                  "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
11260         End Select
11270         Select Case intDec1
              Case 0
11280           strFrmt1 = strTmp01 & "#,##0;(" & strTmp01 & "#,##0);" & strTmp01 & "0;" & strTmp01 & "0"
11290           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11300         Case 1  ' ** None have this currently.
11310           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
11320           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11330         Case 2
11340           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
11350           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11360         Case 3
11370           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
11380           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
11390         End Select
11400         lngBkClr = CLR_LTGRN
11410       End If
11420       .total_cash_sum.Format = strFrmt1
11430       .total_cash_sum.DecimalPlaces = intDec1
11440       .total_cash_sum.BackColor = lngBkClr
11450       .pershare.Format = strFrmt2
11460       .pershare.DecimalPlaces = intDec2
11470       .pershare.BackColor = lngBkClr

11480       If lngCurrID = 150& Then
11490         .total_cash_sum_usd.Visible = False
11500         .pershare_usd.Visible = False
11510       Else
11520         .total_cash_sum_usd = 0
11530         .total_cash_sum_usd.Visible = True
11540         .pershare_usd = 0
11550         .pershare_usd.Visible = True
11560       End If

11570     End If

11580   End With

EXITP:
11590   Exit Sub

ERRH:
11600   Select Case ERR.Number
        Case Else
11610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11620   End Select
11630   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

11710   With Me
11720     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
11730     .curr_id.SetFocus
11740     Response = acDataErrContinue
11750   End With

EXITP:
11760   Exit Sub

ERRH:
11770   Select Case ERR.Number
        Case Else
11780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11790   End Select
11800   Resume EXITP

End Sub

Private Sub description_GotFocus()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

11910   With Me
11920     .description.SelLength = 0
11930     .description.SelStart = 999
11940   End With

EXITP:
11950   Exit Sub

ERRH:
11960   Select Case ERR.Number
        Case Else
11970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11980   End Select
11990   Resume EXITP

End Sub

Private Sub description_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12020   intShiftDown = (Shift And acShiftMask) > 0
12030   intAltDown = (Shift And acAltMask) > 0
12040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12070       With Me
12080         intRetVal = 0
12090         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
12100           .cmbRevenueCodes_Purchase.SetFocus
12110         ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
12120           .cmbTaxCodes_Purchase.SetFocus
12130         Else
12140           .cmdOK.SetFocus
12150         End If
12160       End With
12170     End Select
12180   End If

        ' ** Shift keys.
12190   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12200     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12210       With Me
12220         intRetVal = 0
12230         lngRecsCur = .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Reinvest_DivInt_Detail_Sub.
12240         If lngRecsCur > 0& Then
12250           .frmMap_Reinvest_DivInt_Detail_Sub.SetFocus
12260           .frmMap_Reinvest_DivInt_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Reinvest_DivInt_Detail_Sub.
12270           .frmMap_Reinvest_DivInt_Detail_Sub.Form.FocusHolder.SetFocus
12280         Else
12290           .cmdCancel.SetFocus
12300         End If
12310       End With
12320     End Select
12330   End If

EXITP:
12340   KeyCode = intRetVal
12350   Exit Sub

ERRH:
12360   intRetVal = 0
12370   Select Case ERR.Number
        Case Else
12380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12390   End Select
12400   Resume EXITP

End Sub

Private Sub description_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "description_MouseMove"

12510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
12520   Exit Sub

ERRH:
12530   Select Case ERR.Number
        Case Else
12540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12550   End Select
12560   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_GotFocus()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_GotFocus"

12610   With Me
12620     .cmbRevenueCodes_Purchase.SetFocus
12630   End With

EXITP:
12640   Exit Sub

ERRH:
12650   Select Case ERR.Number
        Case Else
12660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12670   End Select
12680   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12720   intShiftDown = (Shift And acShiftMask) > 0
12730   intAltDown = (Shift And acAltMask) > 0
12740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12770       With Me
12780         intRetVal = 0
12790         .cmbRevenueCodes_Purchase.SetFocus
12800       End With
12810     End Select
12820   End If

        ' ** Shift keys.
12830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12850       With Me
12860         intRetVal = 0
12870         .description.SetFocus
12880       End With
12890     End Select
12900   End If

EXITP:
12910   KeyCode = intRetVal
12920   Exit Sub

ERRH:
12930   intRetVal = 0
12940   Select Case ERR.Number
        Case Else
12950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12960   End Select
12970   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_Enter()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_Enter"

13010   With Me
13020     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
13030       ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
13040     Case False
13050       If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
13060         ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
13070       End If
13080     End Select
13090   End With

EXITP:
13100   Exit Sub

ERRH:
13110   Select Case ERR.Number
        Case Else
13120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13130   End Select
13140   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_AfterUpdate()

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_AfterUpdate"

13210   With Me
13220     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
13230       .cmbRevenueCodes_Purchase = REVID_INC
13240     Case False
            ' ** Nothing at the moment.
13250     End Select
13260     If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
13270       .cmbRevenueCodes_Purchase_display = Null
13280     Else
13290       .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
13300     End If
13310   End With

EXITP:
13320   Exit Sub

ERRH:
13330   Select Case ERR.Number
        Case Else
13340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13350   End Select
13360   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13420   intShiftDown = (Shift And acShiftMask) > 0
13430   intAltDown = (Shift And acAltMask) > 0
13440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13470       With Me
13480         intRetVal = 0
13490         If .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
13500           .cmbTaxCodes_Purchase.SetFocus
13510         Else
13520           .cmdOK.SetFocus
13530         End If
13540       End With
13550     End Select
13560   End If

        ' ** Shift keys.
13570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13590       With Me
13600         intRetVal = 0
13610         .description.SetFocus
13620       End With
13630     End Select
13640   End If

EXITP:
13650   KeyCode = intRetVal
13660   Exit Sub

ERRH:
13670   intRetVal = 0
13680   Select Case ERR.Number
        Case Else
13690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13700   End Select
13710   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_GotFocus()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_GotFocus"

13810   With Me
13820     .cmbTaxCodes_Purchase.SetFocus
13830   End With

EXITP:
13840   Exit Sub

ERRH:
13850   Select Case ERR.Number
        Case Else
13860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13870   End Select
13880   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_KeyDown"

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
13990         .cmbTaxCodes_Purchase.SetFocus
14000       End With
14010     End Select
14020   End If

        ' ** Shift keys.
14030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14050       With Me
14060         intRetVal = 0
14070         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
14080           .cmbRevenueCodes_Purchase.SetFocus
14090         Else
14100           .description.SetFocus
14110         End If
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

Private Sub cmbTaxCodes_Purchase_Enter()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_Enter"

14310   With Me
14320     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
14330       ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
14340     Case False
14350       If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
14360         ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
14370       End If
14380     End Select
14390   End With

EXITP:
14400   Exit Sub

ERRH:
14410   Select Case ERR.Number
        Case Else
14420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14430   End Select
14440   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_AfterUpdate()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_AfterUpdate"

14510   With Me
14520     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
14530       .cmbTaxCodes_Purchase = TAXID_INC
14540     Case False
            ' ** Nothing at the moment.
14550     End Select
14560     If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
14570       .cmbTaxCodes_Purchase_display = Null
14580     Else
14590       .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
14600     End If
14610   End With

EXITP:
14620   Exit Sub

ERRH:
14630   Select Case ERR.Number
        Case Else
14640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14650   End Select
14660   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14720   intShiftDown = (Shift And acShiftMask) > 0
14730   intAltDown = (Shift And acAltMask) > 0
14740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14770       With Me
14780         intRetVal = 0
14790         .cmdOK.SetFocus
14800       End With
14810     End Select
14820   End If

        ' ** Shift keys.
14830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14850       With Me
14860         intRetVal = 0
14870         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
14880           .cmbRevenueCodes_Purchase.SetFocus
14890         Else
14900           .description.SetFocus
14910         End If
14920       End With
14930     End Select
14940   End If

EXITP:
14950   KeyCode = intRetVal
14960   Exit Sub

ERRH:
14970   intRetVal = 0
14980   Select Case ERR.Number
        Case Else
14990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15000   End Select
15010   Resume EXITP

End Sub

Private Sub chkShowLocation_AfterUpdate()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowLocation_AfterUpdate"

15110   With Me
15120     Select Case .chkShowLocation
          Case True
15130       .chkShowLocation_lbl.FontBold = True
15140       .frmMap_Reinvest_DivInt_Detail_Sub.Form.ShowLocation True  ' ** Form Procedure: frmMap_Reinvest_DivInt_Detail_Sub.
15150       ShowLocation True  ' ** Procedure: Below
15160     Case False
15170       .chkShowLocation_lbl.FontBold = False
15180       .frmMap_Reinvest_DivInt_Detail_Sub.Form.ShowLocation False  ' ** Form Procedure: frmMap_Reinvest_DivInt_Detail_Sub.
15190       ShowLocation False  ' ** Procedure: Below.
15200     End Select
15210   End With

EXITP:
15220   Exit Sub

ERRH:
15230   Select Case ERR.Number
        Case Else
15240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15250   End Select
15260   Resume EXITP

End Sub

Private Sub chkShowLocation_KeyDown(KeyCode As Integer, Shift As Integer)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowLocation_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15320   intShiftDown = (Shift And acShiftMask) > 0
15330   intAltDown = (Shift And acAltMask) > 0
15340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15370       With Me
15380         intRetVal = 0
15390         If .cmdPrintReport.Enabled = True Then
15400           .cmdPrintReport.SetFocus
15410         Else
15420           .cmdOK.SetFocus
15430         End If
15440       End With
15450     End Select
15460   End If

        ' ** Shift keys.
15470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15490       With Me
15500         intRetVal = 0
15510         If .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
15520           .cmbTaxCodes_Purchase.SetFocus
15530         ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
15540           .cmbRevenueCodes_Purchase.SetFocus
15550         Else
15560           .description.SetFocus
15570         End If
15580       End With
15590     End Select
15600   End If

EXITP:
15610   KeyCode = intRetVal
15620   Exit Sub

ERRH:
15630   intRetVal = 0
15640   Select Case ERR.Number
        Case Else
15650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15660   End Select
15670   Resume EXITP

End Sub

Public Function RecalcTots() As Boolean
' ** I give up trying to get it from the subform!

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim dblTmp01 As Double, dblTmp02 As Double
        Dim lngX As Long
        Dim blnRetVal As Boolean

15710   With Me
15720     blnRetVal = False
15730     Set rst = .frmMap_Reinvest_DivInt_Detail_Sub.Form.RecordsetClone
15740     With rst
15750       If .BOF = True And .EOF = True Then
              ' ** Ho-Hum...
15760       Else
15770         .MoveLast
15780         lngRecs = .RecordCount
15790         .MoveFirst
15800         dblTmp01 = 0#: dblTmp02 = 0#
15810         For lngX = 1& To lngRecs
15820           dblTmp01 = dblTmp01 + ![total_shareface]
15830           dblTmp02 = dblTmp02 + ![total_icash]
15840           If lngX < lngRecs Then .MoveNext
15850         Next
15860       End If
15870       .Close
15880     End With
15890     If dblTmp01 <> 0# Then
15900       blnRetVal = True
15910     End If
15920     .total_shareface_sum = dblTmp01
15930     .total_cash_sum = dblTmp02
15940   End With

EXITP:
15950   Set rst = Nothing
15960   RecalcTots = blnRetVal
15970   Exit Function

ERRH:
15980   blnRetVal = False
15990   Select Case ERR.Number
        Case Else
16000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16010   End Select
16020   Resume EXITP

End Function

Public Sub IncludeCurrency(blnCalc As Boolean)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim dblRate2 As Double
        Dim blnFound As Boolean
        Dim dblTmp01 As Double

16110   With Me
16120     Select Case blnCalc
          Case True
16130       If blnCurrID = True Then
16140         blnFound = False
16150         Do While blnFound = False
16160           If IsNull(.total_cash_sum) = False Then
16170             If .total_cash_sum <> 0 Then
16180               blnFound = True
16190               dblRate2 = .curr_id.Column(CBX_C_RATE2)
16200               dblTmp01 = (.total_cash_sum * dblRate2)
16210               dblTmp01 = Round(dblTmp01, 2)
16220               .total_cash_sum_usd = dblTmp01
16230               dblTmp01 = (.pershare * dblRate2)
16240               dblTmp01 = Round(dblTmp01, 5)
16250               .pershare_usd = dblTmp01
16260               If lngCurrID <> 150& Then
16270                 .total_cash_sum_usd.Visible = True
16280                 .pershare_usd.Visible = True
16290               End If
16300             End If
16310           End If
16320         Loop
16330       End If
16340     Case False
16350       .curr_id.Enabled = False
16360       .curr_id_cmd.Visible = True
16370       .curr_id_cmd.Enabled = True
16380       .curr_id.Locked = True
16390       .curr_id.ForeColor = CLR_DISABLED_FG
16400       .curr_id.BackColor = CLR_VLTGRN
16410       If blnCurrID = True And lngCurrID <> 150& Then
16420         .curr_id.Visible = True
16430       Else
16440         .curr_id.Visible = False
16450         .curr_id_cmd.Visible = False
16460         .curr_id_cmd.Enabled = False
16470         .cmbRevenueCodes_Purchase_vline05.Visible = False
16480         .cmbRevenueCodes_Purchase_vline06.Visible = False
16490         .cmbRevenueCodes_Purchase_box.Width = (.cmbRevenueCodes_Purchase_vline06.Left - .cmbRevenueCodes_Purchase_box.Left)
16500         .pershare.Top = (.pershare_alt_line.Top + (2& * lngTpp))
16510         .pershare_lbl.Top = (.pershare_alt_line.Top + (3& * lngTpp))
16520         .pershare_lbl1.Top = (.pershare_alt_line.Top + (3& * lngTpp))
16530       End If
16540       .total_cash_sum_usd.Visible = False
16550       .pershare_usd.Visible = False
16560     End Select
16570   End With

EXITP:
16580   Exit Sub

ERRH:
16590   DoCmd.Hourglass False
16600   Select Case ERR.Number
        Case Else
16610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16620   End Select
16630   Resume EXITP

End Sub

Private Sub IncludeCurrency_Sub()

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngNewForm_Width As Long, lngWidth_Diff As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

16710   With Me
16720     If blnCurrID = True And lngCurrID <> 150& Then

16730       lngTmp01 = (.frmMap_Reinvest_DivInt_Detail_Sub.Form.Location_ID.Width + (4& * lngTpp))
16740       lngTmp02 = (.frmMap_Reinvest_DivInt_Detail_Sub.Form.total_icash_usd.Width + (4& * lngTpp))

16750       Select Case .chkShowLocation
            Case True
16760         lngNewForm_Width = (lngForm_Width + lngTmp01 + lngTmp02)
16770         lngWidth_Diff = (lngNewForm_Width - lngForm_Width)
16780       Case False
16790         lngNewForm_Width = (lngForm_Width + lngTmp02)
16800         lngWidth_Diff = (lngNewForm_Width - lngForm_Width)
16810       End Select

16820       lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
16830       lngMonitorNum = 1&: lngTmp04 = 0&
16840       EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
16850       If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

16860       If lngMonitorNum = 1& Then lngTmp04 = lngFrm_Top
16870       DoCmd.MoveSize (lngFrm_Left - (lngWidth_Diff / 2)), lngTmp04, (lngFrm_Width + lngWidth_Diff), lngFrm_Height  'lngFrm_Top
16880       If lngMonitorNum > 1& Then
16890         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
16900       End If

16910       .Width = lngNewForm_Width
16920       .frmMap_Reinvest_DivInt_Detail_Sub.Width = (lngSub_Width + lngWidth_Diff)
16930       .frmMap_Reinvest_DivInt_Detail_Sub_box.Width = ((lngSub_Width + lngWidth_Diff) + (2& * lngTpp))

16940       .Header_vline01.Left = lngNewForm_Width
16950       .Header_vline02.Left = lngNewForm_Width
16960       .Detail_vline01.Left = lngNewForm_Width
16970       .Detail_vline02.Left = lngNewForm_Width
16980       .Footer_vline01.Left = lngNewForm_Width
16990       .Footer_vline02.Left = lngNewForm_Width
17000       .Header_hline01.Width = lngNewForm_Width
17010       .Header_hline02.Width = lngNewForm_Width
17020       .Detail_hline01.Width = lngNewForm_Width
17030       .Detail_hline02.Width = lngNewForm_Width
17040       .Footer_hline01.Width = lngNewForm_Width
17050       .Footer_hline02.Width = lngNewForm_Width

17060       .total_cash_sum_usd.Left = ((.total_cash_sum.Left + .total_cash_sum.Width) + (4& * lngTpp))
17070       .total_cash_sum_usd.Top = .total_cash_sum.Top
17080       .pershare_usd.Left = ((.total_cash_sum_usd.Left + .total_cash_sum_usd.Width) - .pershare_usd.Width)
17090       .pershare.Top = (.pershare_alt_line.Top + (2& * lngTpp))
17100       .pershare_usd.Top = (.pershare_alt_line.Top + (2& * lngTpp))
17110       .pershare_lbl.Visible = False
17120       .pershare_lbl1.Visible = False
17130       Select Case strJType
            Case "Dividend"
17140         lngTmp03 = (.pershare_lbl3.Top - .pershare_lbl2.Top)
17150         .pershare_lbl3.Top = (.pershare_alt_line.Top + (3& * lngTpp))
17160         .pershare_lbl2.Top = .pershare_lbl3.Top - lngTmp03
17170         .pershare_lbl2.Visible = True
17180         .pershare_lbl3.Visible = True
17190       Case "Interest"
17200         lngTmp03 = (.pershare_lbl5.Top - .pershare_lbl4.Top)
17210         .pershare_lbl5.Top = (.pershare_alt_line.Top + (3& * lngTpp))
17220         .pershare_lbl4.Top = .pershare_lbl5.Top - lngTmp03
17230         .pershare_lbl4.Visible = True
17240         .pershare_lbl5.Visible = True
17250       End Select

17260       .total_cash_sum_usd.ForeColor = CLR_DISABLED_FG
17270       .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
17280       .pershare_usd.ForeColor = CLR_DISABLED_FG
17290       .pershare_usd.BackColor = CLR_DISABLED_BG

17300       .maptype.Left = (lngMapType_Left + lngWidth_Diff)
17310       .cmdCancel.Left = (lngCancel_Left + lngWidth_Diff)
17320       .cmdOK.Left = (lngOK_Left + lngWidth_Diff)

17330       .Width = lngNewForm_Width

17340     End If
17350   End With

EXITP:
17360   Exit Sub

ERRH:
17370   DoCmd.Hourglass False
17380   Select Case ERR.Number
        Case Else
17390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17400   End Select
17410   Resume EXITP

End Sub

Private Sub NoChar_Load()

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

17510   Set dbs = CurrentDb
17520   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
17530     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
17540     Set rst = qdf.OpenRecordset
17550     With rst
17560       .MoveLast
17570       lngNoChars = .RecordCount
17580       .MoveFirst
17590       arr_varNoChar = .GetRows(lngNoChars)
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
17600       .Close
17610     End With
17620     .Close
17630   End With

EXITP:
17640   Set rst = Nothing
17650   Set qdf = Nothing
17660   Set dbs = Nothing
17670   Exit Sub

ERRH:
17680   DoCmd.Hourglass False
17690   Select Case ERR.Number
        Case Else
17700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17710   End Select
17720   Resume EXITP

End Sub

Private Sub ShowLocation(blnShow)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowLocation"

        Dim lngNewForm_Width As Long, lngWidth_Diff As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long

17810   With Me

17820     lngTmp01 = (.frmMap_Reinvest_DivInt_Detail_Sub.Form.Location_ID.Width + (4& * lngTpp))
17830     lngTmp02 = (.frmMap_Reinvest_DivInt_Detail_Sub.Form.total_icash_usd.Width + (4& * lngTpp))

17840     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
17850     lngMonitorNum = 1&: lngTmp03 = 0&
17860     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
17870     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

17880     Select Case blnShow
          Case True
17890       If blnCurrID = True And lngCurrID <> 150& Then
17900         lngNewForm_Width = (lngForm_Width + lngTmp01 + lngTmp02)
17910         lngWidth_Diff = (lngNewForm_Width - lngForm_Width)
17920       Else
17930         lngNewForm_Width = (lngForm_Width + lngTmp01)
17940         lngWidth_Diff = ((lngForm_Width + lngTmp01) - lngForm_Width)
17950       End If

17960       If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
17970       DoCmd.MoveSize (lngFrm_Left - (lngWidth_Diff / 2)), lngTmp03, (lngFrm_Width + lngWidth_Diff), lngFrm_Height  'lngFrm_Top
17980       If lngMonitorNum > 1& Then
17990         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
18000       End If

18010       .Width = (lngForm_Width + lngWidth_Diff)
18020       .frmMap_Reinvest_DivInt_Detail_Sub.Width = (lngSub_Width + lngWidth_Diff)
18030       .frmMap_Reinvest_DivInt_Detail_Sub_box.Width = ((lngSub_Width + lngWidth_Diff) + (2& * lngTpp))
18040       .cmdOK.Left = (lngOK_Left + lngWidth_Diff)
18050       .cmdCancel.Left = (lngCancel_Left + lngWidth_Diff)
18060       .maptype.Left = (lngMapType_Left + lngWidth_Diff)
18070       .Header_vline01.Left = .Width
18080       .Header_vline02.Left = .Width
18090       .Detail_vline01.Left = .Width
18100       .Detail_vline02.Left = .Width
18110       .Footer_vline01.Left = .Width
18120       .Footer_vline02.Left = .Width
18130       .Header_hline01.Width = .Width
18140       .Header_hline02.Width = .Width
18150       .Detail_hline01.Width = .Width
18160       .Detail_hline02.Width = .Width
18170       .Footer_hline01.Width = .Width
18180       .Footer_hline02.Width = .Width
18190     Case False
18200       If blnCurrID = True And lngCurrID <> 150& Then
18210         lngNewForm_Width = (lngForm_Width + lngTmp02)
18220         lngWidth_Diff = (lngNewForm_Width - lngForm_Width)
18230       Else
18240         lngNewForm_Width = lngForm_Width
18250         lngWidth_Diff = 0&
18260       End If
18270       .frmMap_Reinvest_DivInt_Detail_Sub.Width = (lngSub_Width + lngWidth_Diff)
18280       .frmMap_Reinvest_DivInt_Detail_Sub_box.Width = ((lngSub_Width + lngWidth_Diff) + (2& * lngTpp))
18290       .cmdOK.Left = (lngOK_Left + lngWidth_Diff)
18300       .cmdCancel.Left = (lngCancel_Left + lngWidth_Diff)
18310       .maptype.Left = (lngMapType_Left + lngWidth_Diff)
18320       .Header_vline01.Left = lngNewForm_Width
18330       .Header_vline02.Left = lngNewForm_Width
18340       .Detail_vline01.Left = lngNewForm_Width
18350       .Detail_vline02.Left = lngNewForm_Width
18360       .Footer_vline01.Left = lngNewForm_Width
18370       .Footer_vline02.Left = lngNewForm_Width
18380       .Header_hline01.Width = lngNewForm_Width
18390       .Header_hline02.Width = lngNewForm_Width
18400       .Detail_hline01.Width = lngNewForm_Width
18410       .Detail_hline02.Width = lngNewForm_Width
18420       .Footer_hline01.Width = lngNewForm_Width
18430       .Footer_hline02.Width = lngNewForm_Width
18440       .Width = lngNewForm_Width
18450       If blnCurrID = True And lngCurrID <> 150& Then
18460         If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
18470         DoCmd.MoveSize (lngFrm_Left - (lngWidth_Diff / 2)), lngTmp03, (lngFrm_Width + lngWidth_Diff), lngFrm_Height  'lngFrm_Top
18480         If lngMonitorNum > 1& Then
18490           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
18500         End If
18510       Else
18520         If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
18530         DoCmd.MoveSize lngFrm_Left, lngTmp03, lngFrm_Width, lngFrm_Height  'lngFrm_Top
18540         If lngMonitorNum > 1& Then
18550           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
18560         End If
18570       End If
18580     End Select
18590   End With

EXITP:
18600   Exit Sub

ERRH:
18610   Select Case ERR.Number
        Case Else
18620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18630   End Select
18640   Resume EXITP

End Sub
