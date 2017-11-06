VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Reinvest_Rec_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Reinvest_Rec_Detail"

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

' ***********************************************************************
' ** NOTE: This subform's RecordSource uses FormRef() for dblPrice,
' **       used as ShareFace, via borrowed gdblCrtRpt_PrinTot variable.
' ***********************************************************************

' ** MapType = "Purchase"

'rptMap_Reinvest_Rec
' ** qryMapReport_06.

'1. ActiveAssets to Journal Map:
'   qryMap_Rec_01
'2. Assign amounts:
'   qryMap_Rec_14_05
'3. Rounding

'4. Delete insignificant amounts

'5. Sum distributions:
'   qryMap_Rec_02_07
'6. Rounding...

'OK REINVEST:
'1. Append Journal Map to tblJournal_Map_Staging3:
'   qryMap_Reinvest_05_04
'PRICE:
'2. Append ActiveAssets to Journal Map:
'   qryMap_Received_01
'   THIS TIME, IT'S TOTAL SHAREFACE X PERSHARE!
'REINVEST DETAIL:
'3. Append Journal Map to tblJournal_Map_Staging:
'   qryMap_Received_02_09
'4. Detail Source:
'   qryMap_Received_02_10
'THE AMT GOING TO THE JOURNAL IS ONLY THE 1ST REC OF JOURNAL MAP!!!!
'5. Append to Journal:
'   qryMap_Received_03

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
Private lngForm_Width As Long, lngSub_Width As Long, lngOK_Left As Long, lngCancel_Left As Long, lngMapType_Left As Long
Private lngPrintReportLbl_Left As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private strFrmt1 As String, strFrmt2 As String, intDec1 As Integer, intDec2 As Integer, lngBkClr As Long
Private blnGoingToReport As Boolean, blnWasGTR As Boolean
Private blnCurrID As Boolean, lngCurrID As Long
Private blnRecalcTots As Boolean
Private dblPrice As Double, lngTaxcode As Long, datTransDate As Date, datAssetDate As Date
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
150           If CharCnt(strTmp01, "~") = 4 Then

160             blnIsOpen = True

                ' ** strCallingForm & "~" & CStr(dblPrice) & "~" & Format(datAssetDate, "mm/dd/yyyy hh:nn:ss")
170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             dblPrice = Val(Left(strTmp01, (intPos01 - 1)))
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             datTransDate = CDate(Left(strTmp01, (intPos01 - 1)))
250             strTmp01 = Mid(strTmp01, (intPos01 + 1))
260             intPos01 = InStr(strTmp01, "~")
270             datAssetDate = CDate(Left(strTmp01, (intPos01 - 1)))
280             lngCurrID = Val(Mid(strTmp01, (intPos01 + 1)))

290             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
300             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

                ' ** Variables are fed empty, then populated ByRef.
310             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

320             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

330             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
340             If lngMonitorCnt > 1& Then
350               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
360               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
370               If lngMonitorNum > 1& Then
380                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
390               End If
400             End If

410             .form_width_line.Left = 0&  ' ** Just get it out of the way.
420             lngForm_Width = .Width
430             lngSub_Width = .frmMap_Reinvest_Rec_Detail_Sub.Width
440             lngOK_Left = .cmdOK.Left
450             lngCancel_Left = .cmdCancel.Left
460             lngMapType_Left = .maptype.Left
470             lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

480             CLR_DISABLED_FG = CLR_DKGRY
490             CLR_DISABLED_BG = CLR_LTTEAL

500             blnCurrID = Pref_CurrID  ' ** Module Function: modPreferenceFuncs.
510             .curr_id = lngCurrID
520             IncludeCurrency_Sub  ' ** Procedure: Below.

530             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
540             blnWasGTR = False

550             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

560             chkShowLocation_AfterUpdate  ' ** Procedure: Below.

570             Set dbs = CurrentDb
580             With dbs
590               Select Case strCallingForm
                  Case "frmJournal"
                    ' ** Journal Map, grouped and summed, by accountno, by specified FormRef().
600                 Set qdf = .QueryDefs("qryMap_Received_02_04")
610               Case "frmJournal_Columns"
                    ' ** Journal Map, grouped and summed, by accountno, by specified FormRef().
620                 Set qdf = .QueryDefs("qryJournal_Columns_35_08")
630               End Select
640               Set rst = qdf.OpenRecordset
650               With rst
660                 If .BOF = True And .EOF = True Then
670                   lngRecs = 0&
680                 Else
690                   .MoveLast
700                   lngRecs = .RecordCount
710                 End If
720                 .Close
730               End With
740               Set rst = Nothing
750               Set qdf = Nothing
760               .Close
770             End With
780             Set dbs = Nothing
790             DoEvents

800             Select Case strCallingForm
                Case "frmJournal"
810               Set dbs = CurrentDb
820               With dbs
                    ' ** Empty tblJournal_Map_Staging.
830                 Set qdf = .QueryDefs("qryMap_Received_02_08")
840                 qdf.Execute
850                 Set qdf = Nothing
860                 DoEvents
                    ' ** Append qryMap_Received_02_04 (Journal Map, linked to Account, qryMap_Received_02_02
                    ' ** (qryMap_Received_02_01 (Journal Map, grouped by accountno, assetno, Location_ID),
                    ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
                    ' ** grouped and summed, by accountno, by specified FormRef()) to tblJournal_Map_Staging.
870                 Set qdf = .QueryDefs("qryMap_Received_02_09")
880                 qdf.Execute
890                 Set qdf = Nothing
900                 DoEvents
                    ' ** jmap_id in tblJournal_Map_Staging has Journal Map ID for reinvests,
                    ' ** and is the same as jmap_id2 in tblJournal_Map_Staging3.
                    ' **
                    ' ** Update qryMap_Reinvest_05_02 (tblJournal_Map_Staging, with DLookups() to
                    ' ** qryMap_Reinvest_05_01 (tblJournal_Map_Staging, linked to tblJournal_Map_Staging3,
                    ' ** with CheckNum_new (parent's Journal_ID))).
910                 Set qdf = .QueryDefs("qryMap_Received_05_03")
920                 qdf.Execute
930                 Set qdf = Nothing
940                 .Close
950               End With
960               Set dbs = Nothing
970               DoEvents
980               If .frmMap_Reinvest_Rec_Detail_Sub.Form.RecordSource <> "qryMap_Received_02_10" Then  'qryMap_Received_02_04
                    ' ** tblJournal_Map_Staging, all fields.
990                 .frmMap_Reinvest_Rec_Detail_Sub.Form.RecordSource = "qryMap_Received_02_10"  'qryMap_Received_02_04
1000              End If
1010            Case "frmJournal_Columns"
1020              Set dbs = CurrentDb
1030              With dbs
                    ' ** Empty tblJournal_Map_Staging.
1040                Set qdf = .QueryDefs("qryJournal_Columns_35_08a")
1050                qdf.Execute
1060                Set qdf = Nothing
1070                DoEvents
                    ' ** Append qryJournal_Columns_35_08 (Journal Map, linked to Account, qryJournal_Columns_30_05_02
                    ' ** (qryJournal_Columns_30_05_01 (Journal Map, grouped by accountno, assetno, Location_ID),
                    ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max),
                    ' ** grouped and summed, by accountno, by specified FormRef()) to tblJournal_Map_Staging.
1080                Set qdf = .QueryDefs("qryJournal_Columns_35_08b")
1090                qdf.Execute
1100                Set qdf = Nothing
                    ' ** jmap_id in tblJournal_Map_Staging has Journal Map ID for reinvests,
                    ' ** and is the same as jmap_id2 in tblJournal_Map_Staging3.
                    ' **
                    ' ** Update qryJournal_Columns_35_12 (tblJournal_Map_Staging, with DLookups() to
                    ' ** qryJournal_Columns_35_11 (tblJournal_Map_Staging, linked to tblJournal_Map_Staging3,
                    ' ** with CheckNum_new (parent's JrnlCol_ID))).
1110                Set qdf = .QueryDefs("qryJournal_Columns_35_13")
1120                qdf.Execute
1130                Set qdf = Nothing
1140                .Close
1150              End With
1160              Set dbs = Nothing
1170              DoEvents
1180              If .frmMap_Reinvest_Rec_Detail_Sub.Form.RecordSource <> "qryJournal_Columns_35_08c" Then  'qryJournal_Columns_35_08
                    ' ** tblJournal_Map_Staging, all fields.
1190                .frmMap_Reinvest_Rec_Detail_Sub.Form.RecordSource = "qryJournal_Columns_35_08c"  'qryJournal_Columns_35_08
1200              End If
1210            End Select
1220            .frmMap_Reinvest_Rec_Detail_Sub.Form.Requery
1230            DoEvents

1240            .maptype = "Purchase"
1250            .txtNumberRecords = "Number Records = " & CStr(lngRecs)
1260            .pershare = dblPrice
1270            .description = "Long Term Capital Gain Reinvestment"

1280            If .cmdOK.Enabled = True Then
1290              .cmdOK.SetFocus
1300            Else
1310              .description.SetFocus
1320            End If

1330            blnRecalcTots = True
1340            .TimerInterval = 500&

1350            DoEvents
1360            DoCmd.SelectObject acForm, THIS_NAME, False
1370            DoEvents

1380            DoCmd.Hourglass False

1390            blnIsOpen = False

1400          Else
1410            Cancel = -1
1420          End If
1430        Else
1440          Cancel = -1
1450        End If
1460      Else
1470        Cancel = -1
1480      End If
1490    End With

1500    If Cancel = -1 Then
1510      Beep
1520      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1530    End If

EXITP:
1540    Set rst = Nothing
1550    Set qdf = Nothing
1560    Set dbs = Nothing
1570    Exit Sub

ERRH:
1580    DoCmd.Hourglass False
1590    Select Case ERR.Number
        Case Else
1600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1610    End Select
1620    Resume EXITP

End Sub

Private Sub Form_Load()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1710    With Me

1720      .total_shareface_sum.ForeColor = CLR_DISABLED_FG
1730      .total_shareface_sum.BackColor = CLR_DISABLED_BG
1740      .total_cash_sum.ForeColor = CLR_DISABLED_FG
1750      .total_cash_sum.BackColor = CLR_DISABLED_BG
1760      .total_cash_sum_usd.ForeColor = CLR_DISABLED_FG
1770      .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
1780      .pershare.ForeColor = CLR_DISABLED_FG
1790      .pershare.BackColor = CLR_DISABLED_BG
1800      .pershare_usd.ForeColor = CLR_DISABLED_FG
1810      .pershare_usd.BackColor = CLR_DISABLED_BG

          ' ** Borrowing these variables from the Court Reports.
1820      If gstrCrtRpt_Version <> vbNullString Then
1830        .cusip = gstrCrtRpt_Ordinal
1840        .cusip.ForeColor = CLR_DISABLED_FG
1850        .cusip.BackColor = CLR_DISABLED_BG
1860        .cusip.Visible = True
1870        .totdesc = gstrCrtRpt_Version
1880        .totdesc.ForeColor = CLR_DISABLED_FG
1890        .totdesc.BackColor = CLR_DISABLED_BG
1900        .totdesc.Visible = True
1910      End If

1920      If gstrAccountNo <> vbNullString Then
            ' ** An account was specified.
1930        If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1940          lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1950        End If
1960        .cusip.Top = .totdesc.Top
1970        If blnCurrID = True Or .chkShowLocation = True Then
1980          .shortname.Width = .totdesc.Width
1990        Else
2000          .totdesc.Width = .shortname.Width
2010        End If
2020        .totdesc.Left = ((.cusip.Left + .cusip.Width) + (4& * lngTpp))
2030        .accountno = gstrAccountNo
2040        .shortname = gstrAccountName
2050        .accountno.Visible = True
2060        .shortname.Visible = True
2070      End If

2080      .transdate = datTransDate
2090      .assetdate = datAssetDate

2100      Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
2110        .cmbRevenueCodes_Purchase = REVID_INC
2120        .cmbRevenueCodes_Purchase_display = Null
2130      Case False
2140        If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
2150          .cmbRevenueCodes_Purchase_display = Null
2160        Else
2170          .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
2180        End If
2190      End Select

2200      Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
2210        .cmbTaxCodes_Purchase = TAXID_INC
2220        .cmbTaxCodes_Purchase_display = Null
2230      Case False
2240        If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
2250          .cmbTaxCodes_Purchase_display = Null
2260        Else
2270          .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
2280        End If
2290      End Select

2300      Select Case gblnRevenueExpenseTracking
          Case True
2310        .cmbRevenueCodes_Purchase.Visible = True
2320        .cmbRevenueCodes_Purchase_display.Visible = True
2330        .cmbRevenueCodes_Purchase_lbl2.Visible = False  ' ** Option off msg.
2340      Case False
2350        .cmbRevenueCodes_Purchase.Visible = False
2360        .cmbRevenueCodes_Purchase_display.Visible = False
2370        .cmbRevenueCodes_Purchase_lbl_box.Visible = False
2380        .cmbRevenueCodes_Purchase_lbl2.Visible = True  ' ** Option off msg.
2390      End Select

2400      Select Case gblnIncomeTaxCoding
          Case True
2410        .cmbTaxCodes_Purchase.Visible = True
2420        .cmbTaxCodes_Purchase_display.Visible = True
2430        .cmbTaxCodes_Purchase_lbl2.Visible = False  ' ** Option off msg.
2440      Case False
2450        .cmbTaxCodes_Purchase.Visible = False
2460        .cmbTaxCodes_Purchase_display.Visible = False
2470        .cmbTaxCodes_Purchase_lbl_box.Visible = False  ' ** Option off msg.
2480        .cmbTaxCodes_Purchase_lbl2.Visible = True
2490      End Select

2500      If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
2510        .cmbRevenueCodes_Purchase_box_lbl.ForeColor = WIN_CLR_DISF
2520        .cmbRevenueCodes_Purchase_box_lbl_dim_hi.Visible = True
2530      End If

2540      curr_id_AfterUpdate  ' ** Procedure: Below.
2550      IncludeCurrency False  ' ** Procedure: Below
2560      DoEvents

2570      .frmMap_Reinvest_Rec_Detail_Sub.Form.Requery
2580      DoEvents
2590      .frmMap_Reinvest_Rec_Detail_Sub.Form.TimerInterval = 1000&

2600      DoCmd.Hourglass False

2610    End With

EXITP:
2620    Exit Sub

ERRH:
2630    DoCmd.Hourglass False
2640    Select Case ERR.Number
        Case Else
2650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2660    End Select
2670    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Comments:         M {description}

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Recalc Totals:    F5 {RecalcTots}

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyEscape
2770        intRetVal = 0
2780        cmdCancel_Click  ' ** Procedure: Below.
2790      Case vbKeyF5
2800        intRetVal = 0
2810        RecalcTots  ' ** Function: Below.
2820      End Select
2830    End If

        ' ** Ctrl keys.
2840    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2850      Select Case intRetVal
          Case vbKeyM
2860        With Me
2870          intRetVal = 0
2880          .description.SetFocus
2890        End With
2900      End Select
2910    End If

        ' ** Ctrl-Shift keys.
2920    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2930      Select Case intRetVal
          Case vbKeyF
2940        With Me
2950          intRetVal = 0
2960          .FocusHolder.SetFocus
2970        End With
2980      End Select
2990    End If

EXITP:
3000    KeyCode = intRetVal
3010    Exit Sub

ERRH:
3020    intRetVal = 0
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub Form_Timer()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim blnRetVal As Boolean

3110    With Me
3120      .TimerInterval = 0&
3130      If gblnGoToReport = True Then
3140        Select Case blnGoingToReport
            Case True
3150          DoCmd.Hourglass False
3160          gblnGoToReport = False
3170          blnGoingToReport = False
3180          .GoToReport_arw_pr_img.Visible = False
3190          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
3200          blnWasGTR = True
3210          ReportList_DelJournal  ' ** Module Function: modReportFunctions.
3220          .cmdCancel.SetFocus
3230          DoEvents
3240        Case False
3250          DoCmd.Hourglass True  ' ** Make sure it's still running.
3260          DoEvents
3270          blnGoingToReport = True
3280          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3290          .GoToReport_arw_pr_img.Visible = True
3300          If .cmdPrintReport.Enabled = True Then
3310            .cmdPrintReport.SetFocus
3320          End If
3330          DoEvents
3340          Beep
3350          DoCmd.Hourglass False
3360          DoEvents
3370          .TimerInterval = (GTR_WAIT * 2&)
3380        End Select
3390      End If
3400      If blnRecalcTots = True Then
3410        blnRecalcTots = False
3420        blnRetVal = RecalcTots  ' ** Function: Below.
3430        If blnRetVal = False Then
3440          blnRecalcTots = True
3450          .TimerInterval = 500&
3460        End If
3470      End If
3480    End With

EXITP:
3490    Exit Sub

ERRH:
3500    DoCmd.Hourglass False
3510    Select Case ERR.Number
        Case Else
3520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3530    End Select
3540    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3610    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3620    Set dbs = CurrentDb
        ' ** Empty tmpAccount.
3630    Set qdf = dbs.QueryDefs("qryMap_Rec_08_01")
3640    qdf.Execute
3650    Set qdf = Nothing
3660    DoEvents
3670    dbs.Close
3680    Set dbs = Nothing
3690    DoEvents

3700    gstrReportCallingForm = vbNullString
3710    gstrReportQuerySpec = vbNullString

3720    If gstrReturningForm = vbNullString Then
3730      gstrReturningForm = THIS_NAME
3740    End If

EXITP:
3750    Set qdf = Nothing
3760    Set dbs = Nothing
3770    Exit Sub

ERRH:
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
        Dim lngAssetNo As Long, lngRevCode As Long, lngRevType As Long, lngTaxType As Long
        Dim strRevDesc As String, strTaxDesc As String
        Dim varComments As Variant
        Dim blnLocChanged As Boolean
        Dim varTmp00 As Variant

4310    With Me

4320      lngRecsCur = .frmMap_Reinvest_Rec_Detail_Sub.Form.RecCnt
4330      If lngRecsCur > 0& Then

4340        DoCmd.Hourglass True
4350        DoEvents

4360        blnLocChanged = False

4370        lngAssetNo = .frmMap_Reinvest_Rec_Detail_Sub.Form.assetno
4380        lngRevCode = .cmbRevenueCodes_Purchase
4390        strRevDesc = .cmbRevenueCodes_Purchase.Column(1)
4400        lngRevType = .cmbRevenueCodes_Purchase.Column(2)
4410        lngTaxcode = .cmbTaxCodes_Purchase
4420        strTaxDesc = .cmbTaxCodes_Purchase.Column(1)
4430        lngTaxType = .cmbTaxCodes_Purchase.Column(2)
4440        varComments = .description

4450        Set dbs = CurrentDb
4460        With dbs

              ' ** Update qryMap_Reinvest_02_13 (tblJournal_Map_Staging, linked
              ' ** to Journal, for pcash, by specified [astno]), for Reinvested = True.
4470          Set qdf = .QueryDefs("qryMap_Reinvest_02_14")
4480          With qdf.Parameters
4490            ![astno] = lngAssetNo
4500          End With
4510          qdf.Execute
4520          Set qdf = Nothing
4530          DoEvents

4540          Select Case strCallingForm
              Case "frmJournal"

                ' ** qryMap_Received_03a (tblJournal_Map_Staging, just loc_changed = True), grouped, with cnt.
4550            varTmp00 = DLookup("[cnt]", "qryMap_Received_03b")
4560            If IsNull(varTmp00) = False Then
4570              If varTmp00 > 0 Then
4580                blnLocChanged = True
4590              End If
4600            End If
4610            If blnLocChanged = True Then
                  ' ** Update qryMap_Received_03c (Journal Map, linked to qryMap_Received_03a
                  ' ** (tblJournal_Map_Staging, just loc_changed = True), with Location_ID_new).
4620              Set qdf = .QueryDefs("qryMap_Received_03d")
4630              qdf.Execute
4640              Set qdf = Nothing
4650            End If
4660            DoEvents

                ' ** Append Journal Map to Journal, by specified [pric], [astdat], [desc], [revcod], [taxcod].
4670            Set qdf = .QueryDefs("qryMap_Received_03")

4680          Case "frmJournal_Columns"

                ' ** qryJournal_Columns_35_09a (tblJournal_Map_Staging, just loc_changed = True), grouped, with cnt.
4690            varTmp00 = DLookup("[cnt]", "qryJournal_Columns_35_09b")
4700            If IsNull(varTmp00) = False Then
4710              If varTmp00 > 0 Then
4720                blnLocChanged = True
4730              End If
4740            End If
4750            If blnLocChanged = True Then
                  ' ** Update qryJournal_Columns_35_09c (Journal Map, linked to qryJournal_Columns_35_09a
                  ' ** (tblJournal_Map_Staging, just loc_changed = True), with Location_ID_new).
4760              Set qdf = .QueryDefs("qryJournal_Columns_35_09d")
4770              qdf.Execute
4780              Set qdf = Nothing
4790            End If
4800            DoEvents

                ' ** Append Journal Map to tblJournal_Column, by specified [pric], [astdat],
                ' ** [desc], [revcod], [taxcod], [revdesc], [taxdesc], [revtyp], [taxtyp].
4810            Set qdf = .QueryDefs("qryJournal_Columns_35_09")
4820            With qdf.Parameters
4830              ![revdesc] = strRevDesc
4840              ![taxdesc] = strTaxDesc
4850              ![revtyp] = lngRevType
4860              ![taxtyp] = lngTaxType
4870            End With

4880          End Select
4890          With qdf.Parameters
4900            ![pric] = CCur(dblPrice)
4910            ![astdat] = CDate(Format(datAssetDate, "mm/dd/yyyy"))  ' ** Query adds Time().
4920            ![Desc] = varComments
4930            ![revcod] = lngRevCode
4940            ![taxcod] = lngTaxcode
4950          End With
4960          qdf.Execute dbFailOnError
4970          Set qdf = Nothing
4980          DoEvents

              ' ** Use tblJournal_Map_Staging3 to populate tblJournal_MiscSold.
4990          Select Case strCallingForm
              Case "frmJournal"
                ' ** Journal_ID in .._Staging3 is parents' Journal ID.
                ' ** CheckNum in parents' Journal is children's .._Staging3 jmap_id1.
                ' ** CheckNum in children's Journal is parents' Journal ID.
                ' **
                ' ** Append tblJournal_Map_Staging3, linked to Journal, to
                ' ** tblJournal_MiscSold, by specified [jtyp], for parent record.
5000            Set qdf = .QueryDefs("qryMap_Reinvest_07_01")
5010            With qdf.Parameters
5020              ![jtyp] = "Received"
5030            End With
5040            qdf.Execute
5050            Set qdf = Nothing
5060            DoEvents
                ' ** Append tblJournal_Map_Staging3, linked to Journal, to
                ' ** tblJournal_MiscSold, by specified [jtyp], for child record.
5070            Set qdf = .QueryDefs("qryMap_Reinvest_07_02")
5080            With qdf.Parameters
5090              ![jtyp] = "Received"
5100            End With
5110            qdf.Execute
5120            Set qdf = Nothing
5130          Case "frmJournal_Columns"
                ' ** JrnlCol_ID in .._Staging3 is parents' JrnlCol_ID.
                ' ** CheckNum in parents' tblJournal_Column is children's .._Staging3 jmap_id1.
                ' ** CheckNum in children's tblJournal_Column is parents' JrnlCol_ID.
                ' **
                ' ** Empty tblJournal_MiscSold_Staging.
5140            Set qdf = .QueryDefs("qryJournal_Columns_32_06")
5150            qdf.Execute
5160            Set qdf = Nothing
5170            DoEvents
                ' ** Append tblJournal_Map_Staging3, linked to tblJournal_Column, to
                ' ** tblJournal_MiscSold_Staging, by specified [jtyp], for parent record.
5180            Set qdf = .QueryDefs("qryJournal_Columns_32_04")
5190            With qdf.Parameters
5200              ![jtyp] = "Received"
5210            End With
5220            qdf.Execute
5230            Set qdf = Nothing
5240            DoEvents
                ' ** Append tblJournal_Map_Staging3, linked to tblJournal_Column, to
                ' ** tblJournal_MiscSold_Staging, by specified [jtyp], for child record.
5250            Set qdf = .QueryDefs("qryJournal_Columns_32_05")
5260            With qdf.Parameters
5270              ![jtyp] = "Received"
5280            End With
5290            qdf.Execute
5300            Set qdf = Nothing
5310          End Select
5320          DoEvents

5330          Select Case strCallingForm
              Case "frmJournal"
                ' ** Empty Journal Map.
5340            Set qdf = .QueryDefs("qryMap_01")
5350            qdf.Execute
5360            Set qdf = Nothing
5370            DoEvents
                ' ** Empty tblJournal_Map_Staging.
5380            Set qdf = .QueryDefs("qryMap_Received_02_08")
5390            qdf.Execute
5400            Set qdf = Nothing
5410            DoEvents
                ' ** Empty tblJournal_Map_Staging3.
5420            Set qdf = .QueryDefs("qryMap_Reinvest_05_07")
5430            qdf.Execute
5440            Set qdf = Nothing
5450          Case "frmJournal_Columns"
                ' ** Empty Journal Map.
5460            Set qdf = .QueryDefs("qryJournal_Columns_35_04")
5470            qdf.Execute
5480            Set qdf = Nothing
5490            DoEvents
                ' ** Empty tblJournal_Map_Staging.
5500            Set qdf = .QueryDefs("qryJournal_Columns_35_08a")
5510            qdf.Execute
5520            Set qdf = Nothing
5530            DoEvents
                'WE CAN'T EMPTY THIS HERE! IT'S USED IN JC_Frm_Map_Return!
                ' ** Empty tblJournal_Map_Staging3.
                'Set qdf = .QueryDefs("qryJournal_Columns_32_03")
                'qdf.Execute
                'Set qdf = Nothing
5540          End Select
5550          DoEvents

5560          .Close
5570        End With
5580        Set dbs = Nothing
5590        DoEvents

5600        Select Case strCallingForm
            Case "frmJournal"
5610          If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5620            DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
5630          Else
5640            DoCmd.SelectObject acForm, strCallingForm, False
5650          End If
5660          DoEvents
5670          Forms(strCallingForm).frmJournal_Sub5_Misc.Form.Requery
5680          DoEvents
5690          Forms(strCallingForm).frmJournal_Sub5_Misc.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub5_Misc.
5700          DoEvents
5710        Case "frmJournal_Columns"
5720          gblnSetFocus = True
5730          gstrReturningForm = "frmMap_Return"
5740          If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5750            DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
5760          Else
5770            Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
5780            Forms(strCallingForm).frmJournal_Columns_Sub.Form.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
5790            DoCmd.SelectObject acForm, strCallingForm, False
5800            Forms(strCallingForm).TimerInterval = 250&
5810          End If
5820        End Select

5830        DoCmd.Close acForm, THIS_NAME

5840      Else
5850        Beep
5860      End If  ' ** lngRecsCur.

5870      DoCmd.Hourglass False

5880    End With

EXITP:
5890    Set qdf = Nothing
5900    Set dbs = Nothing
5910    Exit Sub

ERRH:
5920    DoCmd.Hourglass False
5930    Select Case ERR.Number
        Case Else
5940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5950    End Select
5960    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6020    intShiftDown = (Shift And acShiftMask) > 0
6030    intAltDown = (Shift And acAltMask) > 0
6040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6060      Select Case intRetVal
          Case vbKeyTab
6070        With Me
6080          intRetVal = 0
6090          .cmdCancel.SetFocus
6100        End With
6110      End Select
6120    End If

        ' ** Shift keys.
6130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6140      Select Case intRetVal
          Case vbKeyTab
6150        With Me
6160          intRetVal = 0
6170          .cmdPrintReport.SetFocus
6180        End With
6190      End Select
6200    End If

EXITP:
6210    KeyCode = intRetVal
6220    Exit Sub

ERRH:
6230    intRetVal = 0
6240    Select Case ERR.Number
        Case Else
6250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6260    End Select
6270    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intTmp01 As Integer
        Dim intX As Integer

6310    DoCmd.Hourglass True
6320    DoEvents

6330    intTmp01 = 0
6340    Do While Forms.Count > 1
6350      intTmp01 = intTmp01 + 1
6360      For intX = (Forms.Count - 1) To 0 Step -1
6370        Select Case Forms(intX).Name
            Case THIS_NAME
              ' ** Skip.
6380        Case "frmJournal", "frmJournal_Columns"
              ' ** Skip.
6390        Case Else
6400          DoCmd.Close acForm, Forms(intX).Name
6410          Exit For  ' ** Do them one-at-a-time.
6420        End Select
6430      Next
6440      If intTmp01 >= 10 Then Exit Do  ' ** Don't get caught in a loop.
6450    Loop

6460    Set dbs = CurrentDb
6470    With dbs

          ' ** See that posted and reinvested are turned off,
          ' ** and Null-out CheckNum on parent transactions.
6480      Select Case strCallingForm
          Case "frmJournal"
            ' ** Update qryJournal_Columns_32_01 (tblJournal_Map_Staging3, linked to
            ' ** Journal, with posted_new, Reinvested_new, Check_Num_new).
6490        Set qdf = .QueryDefs("qryJournal_Columns_32_02")
6500        qdf.Execute
6510        Set qdf = Nothing
6520        DoEvents
            ' ** Empty Journal Map.
6530        Set qdf = .QueryDefs("qryMap_01")
6540        qdf.Execute
6550        Set qdf = Nothing
6560        DoEvents
            ' ** Empty tblJournal_Map_Staging.
6570        Set qdf = .QueryDefs("qryMap_Received_02_08")
6580        qdf.Execute
6590        Set qdf = Nothing
6600        DoEvents
            ' ** Empty tblJournal_Map_Staging3.
6610        Set qdf = .QueryDefs("qryMap_Reinvest_05_07")
6620        qdf.Execute
6630        Set qdf = Nothing
6640      Case "frmJournal_Columns"
            ' ** Update qryJournal_Columns_32_01 (tblJournal_Map_Staging3, linked to
            ' ** tblJournal_Column, with posted_new, Reinvested_new, Check_Num_new).
6650        Set qdf = .QueryDefs("qryJournal_Columns_32_02")
6660        qdf.Execute
6670        Set qdf = Nothing
6680        DoEvents
            ' ** Empty Journal Map.
6690        Set qdf = .QueryDefs("qryJournal_Columns_30_04")
6700        qdf.Execute
6710        Set qdf = Nothing
6720        DoEvents
            ' ** Empty tblJournal_Map_Staging.
6730        Set qdf = .QueryDefs("qryJournal_Columns_35_08a")
6740        qdf.Execute
6750        Set qdf = Nothing
6760        DoEvents
            ' ** Empty tblJournal_Map_Staging3.
6770        Set qdf = .QueryDefs("qryJournal_Columns_32_03")
6780        qdf.Execute
6790        Set qdf = Nothing
6800      End Select
6810      DoEvents

6820      .Close
6830    End With
6840    Set dbs = Nothing
6850    DoEvents

6860    gblnSetFocus = True
6870    Select Case strCallingForm
        Case "frmJournal"
6880      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
6890        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
6900      Else
6910        gstrReturningForm = THIS_NAME
6920        Forms(strCallingForm).frmJournal_Sub5_Misc.Form.Requery
6930        DoEvents
6940        DoCmd.SelectObject acForm, strCallingForm, False
6950        Forms(strCallingForm).TimerInterval = 250&
6960      End If
6970    Case "frmJournal_Columns"
6980      gstrReturningForm = "frmMap_Return"
6990      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
7000        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
7010      Else
7020        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
7030        DoCmd.SelectObject acForm, strCallingForm, False
7040        Forms(strCallingForm).TimerInterval = 250&
7050      End If
7060    End Select

7070    DoCmd.Close acForm, THIS_NAME

EXITP:
7080    Set qdf = Nothing
7090    Set dbs = Nothing
7100    Exit Sub

ERRH:
7110    DoCmd.Hourglass False
7120    Select Case ERR.Number
        Case Else
7130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7140    End Select
7150    Resume EXITP

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
7290          lngRecsCur = .frmMap_Reinvest_Rec_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Reinvest_Rec_Detail_Sub.
7300          If lngRecsCur > 0 Then
7310            .frmMap_Reinvest_Rec_Detail_Sub.SetFocus
7320            .frmMap_Reinvest_Rec_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMap_Reinvest_Rec_Detail_Sub.
7330            .frmMap_Reinvest_Rec_Detail_Sub.Form.FocusHolder.SetFocus
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
7440          .cmdOK.SetFocus
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

        ' ** Journal Map, by specified FormRef('Price').
        ' ** qryMapReport_06.
        'strSQL = "SELECT DISTINCTROW account.accountno, account.shortname, Sum([journal map].shareface) AS total_shareface, " & _
        '  "Sum([journal map].pcash) AS total_pcash, " & CoInfo & ", " & _
        '  "CStr([masterasset].[Description]) & IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy')) AS totdesc " & _
        '  "FROM ([journal map] INNER JOIN account ON [journal map].accountno = account.accountno) " & _
        '  "INNER JOIN masterasset ON [journal map].assetno = masterasset.assetno " & _
        '  "GROUP BY account.accountno, account.shortname, CStr([masterasset].[Description]) & " & _
        '  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'));"

8010    With Me
8020      strDocName = "rptMap_Reinvest_Rec"
8030      gstrReportCallingForm = THIS_NAME
8040      gstrReportQuerySpec = strDocName
8050      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
8060        .Modal = False
8070        DoCmd.OpenReport strDocName, acViewPreview, , , , acViewPreview
8080        DoCmd.Maximize
8090        DoCmd.RunCommand acCmdFitToWindow
8100        .Visible = False
8110      Else
            '##GTR_Ref: rptMap_Reinvest_Rec
8120        DoCmd.OpenReport strDocName, acViewNormal, , , , acViewNormal
8130      End If
8140    End With

EXITP:
8150    Exit Sub

ERRH:
8160    Me.Visible = True
8170    DoCmd.Restore
8180    Select Case ERR.Number
        Case Else
8190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8200    End Select
8210    Resume EXITP

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
8470          .chkShowLocation.SetFocus
8480        End With
8490      End Select
8500    End If

EXITP:
8510    KeyCode = intRetVal
8520    Exit Sub

ERRH:
8530    intRetVal = 0
8540    Select Case ERR.Number
        Case Else
8550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8560    End Select
8570    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

8610    With Me
8620      If blnPrintReport_MouseDown = False Then
8630        Select Case blnPrintReport_Focus
            Case True
8640          .cmdPrintReport_raised_focus_dots_img.Visible = True
8650          .cmdPrintReport_raised_focus_img.Visible = False
8660        Case False
8670          .cmdPrintReport_raised_focus_img.Visible = True
8680          .cmdPrintReport_raised_focus_dots_img.Visible = False
8690        End Select
8700        .cmdPrintReport_raised_img.Visible = False
8710        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8720        .cmdPrintReport_sunken_focus_dots_img.Visible = False
8730        .cmdPrintReport_raised_img_dis.Visible = False
8740      End If
8750    End With

EXITP:
8760    Exit Sub

ERRH:
8770    Select Case ERR.Number
        Case Else
8780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8790    End Select
8800    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

8910    With Me
8920      .cmdPrintReport_raised_focus_dots_img.Visible = True
8930      .cmdPrintReport_raised_img.Visible = False
8940      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
8950      .cmdPrintReport_raised_focus_img.Visible = False
8960      .cmdPrintReport_sunken_focus_dots_img.Visible = False
8970      .cmdPrintReport_raised_img_dis.Visible = False
8980      blnPrintReport_MouseDown = False
8990    End With

EXITP:
9000    Exit Sub

ERRH:
9010    Select Case ERR.Number
        Case Else
9020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9030    End Select
9040    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

9110    With Me
9120      .cmdPrintReport_raised_img.Visible = True
9130      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
9140      .cmdPrintReport_raised_focus_img.Visible = False
9150      .cmdPrintReport_raised_focus_dots_img.Visible = False
9160      .cmdPrintReport_sunken_focus_dots_img.Visible = False
9170      .cmdPrintReport_raised_img_dis.Visible = False
9180      blnPrintReport_Focus = False
9190    End With

EXITP:
9200    Exit Sub

ERRH:
9210    Select Case ERR.Number
        Case Else
9220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9230    End Select
9240    Resume EXITP

End Sub

Private Sub total_cash_sum_KeyDown(KeyCode As Integer, Shift As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "total_cash_sum_KeyDown"

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
9470          lngRecsCur = .frmMap_Reinvest_Rec_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Reinvest_Rec_Detail_Sub.
9480          If lngRecsCur > 0& Then
9490            .frmMap_Reinvest_Rec_Detail_Sub.SetFocus
9500            .frmMap_Reinvest_Rec_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Reinvest_Rec_Detail_Sub.
9510            .frmMap_Reinvest_Rec_Detail_Sub.Form.FocusHolder.SetFocus
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
9870          lngRecsCur = .frmMap_Reinvest_Rec_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Reinvest_Rec_Detail_Sub.
9880          If lngRecsCur > 0& Then
9890            .frmMap_Reinvest_Rec_Detail_Sub.SetFocus
9900            .frmMap_Reinvest_Rec_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Reinvest_Rec_Detail_Sub.
9910            .frmMap_Reinvest_Rec_Detail_Sub.Form.FocusHolder.SetFocus
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
10490         .pershare.FontName = "Courier New"
10500         .pershare.FontSize = 10
10510         .pershare.FontBold = True
10520       Case False
10530         If .total_cash_sum.FontName <> "Arial" Then
10540           .total_cash_sum.FontName = "Arial"
10550           .total_cash_sum.FontSize = 10
10560           .total_cash_sum.FontBold = False
10570           .pershare.FontName = "Arial"
10580           .pershare.FontSize = 10
10590           .pershare.FontBold = False
10600         End If
10610       End Select

10620       If lngCurrID = 150& Then  ' ** USD.
10630         strFrmt1 = "Currency"
10640         intDec1 = 2
10650         strFrmt2 = "Currency"
10660         intDec2 = 5
10670         lngBkClr = CLR_DISABLED_BG
10680       Else
10690         strTmp01 = .curr_id.Column(CBX_C_SYM)  ' ** Symbol.
10700         If .curr_id.Column(CBX_C_DEC) = "." Then
                ' ** These are odd ones, and I don't know how they're displayed.
10710           intDec1 = 2
10720         Else
10730           intDec1 = .curr_id.Column(CBX_C_DEC)  ' ** Decimal places.
10740         End If
10750         intDec2 = 5
10760         intLen = Len(strTmp01)
10770         Select Case intLen
              Case 1
                'strTmp01 = strTmp01
10780         Case 2
10790           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2)
10800         Case 3
10810           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3)
10820         Case 4
10830           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & "\" & Mid(strTmp01, 4)
10840         Case 5
10850           strTmp01 = "\" & Left(strTmp01, 1) & "\" & Mid(strTmp01, 2, 1) & "\" & Mid(strTmp01, 3, 1) & _
                  "\" & Mid(strTmp01, 4, 1) & "\" & Mid(strTmp01, 5)
10860         End Select
10870         Select Case intDec1
              Case 0
10880           strFrmt1 = strTmp01 & "#,##0;(" & strTmp01 & "#,##0);" & strTmp01 & "0;" & strTmp01 & "0"
10890           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
10900         Case 1  ' ** None have this currently.
10910           strFrmt1 = strTmp01 & "#,##0.0;(" & strTmp01 & "#,##0.0);" & strTmp01 & "0.0;" & strTmp01 & "0.0"
10920           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
10930         Case 2
10940           strFrmt1 = strTmp01 & "#,##0.00;(" & strTmp01 & "#,##0.00);" & strTmp01 & "0.00;" & strTmp01 & "0.00"
10950           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
10960         Case 3
10970           strFrmt1 = strTmp01 & "#,##0.000;(" & strTmp01 & "#,##0.000);" & strTmp01 & "0.000;" & strTmp01 & "0.000"
10980           strFrmt2 = strTmp01 & "#,##0.00000;(" & strTmp01 & "#,##0.00000);" & strTmp01 & "0.00000;" & strTmp01 & "0.00000"
10990         End Select
11000         lngBkClr = CLR_LTGRN
11010       End If
11020       .total_cash_sum.Format = strFrmt1
11030       .total_cash_sum.DecimalPlaces = intDec1
11040       .total_cash_sum.BackColor = lngBkClr
11050       .pershare.Format = strFrmt2
11060       .pershare.DecimalPlaces = intDec2
11070       .pershare.BackColor = lngBkClr

11080       If lngCurrID = 150& Then
11090         .total_cash_sum_usd.Visible = False
11100         .pershare_usd.Visible = False
11110       Else
11120         .total_cash_sum_usd = 0
11130         .total_cash_sum_usd.Visible = True
11140         .pershare_usd = 0
11150         .pershare_usd.Visible = True
11160       End If

11170     End If

11180   End With

EXITP:
11190   Exit Sub

ERRH:
11200   Select Case ERR.Number
        Case Else
11210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11220   End Select
11230   Resume EXITP

End Sub

Private Sub curr_id_NotInList(NewData As String, Response As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_NotInList"

11310   With Me
11320     MsgBox "Please select a valid currency code.", vbInformation + vbOKOnly, "Invalid Entry"
11330     .curr_id.SetFocus
11340     Response = acDataErrContinue
11350   End With

EXITP:
11360   Exit Sub

ERRH:
11370   Select Case ERR.Number
        Case Else
11380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11390   End Select
11400   Resume EXITP

End Sub

Private Sub description_GotFocus()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

11510   With Me
11520     .description.SelLength = 0
11530     .description.SelStart = 999
11540   End With

EXITP:
11550   Exit Sub

ERRH:
11560   Select Case ERR.Number
        Case Else
11570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11580   End Select
11590   Resume EXITP

End Sub

Private Sub description_KeyDown(KeyCode As Integer, Shift As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "description_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11620   intShiftDown = (Shift And acShiftMask) > 0
11630   intAltDown = (Shift And acAltMask) > 0
11640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11670       With Me
11680         intRetVal = 0
11690         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
11700           .cmbRevenueCodes_Purchase.SetFocus
11710         ElseIf .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
11720           .cmbTaxCodes_Purchase.SetFocus
11730         Else
11740           .cmdOK.SetFocus
11750         End If
11760       End With
11770     End Select
11780   End If

        ' ** Shift keys.
11790   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11800     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11810       With Me
11820         intRetVal = 0
11830         lngRecsCur = .frmMap_Reinvest_Rec_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Reinvest_Rec_Detail_Sub.
11840         If lngRecsCur > 0& Then
11850           .frmMap_Reinvest_Rec_Detail_Sub.SetFocus
11860           .frmMap_Reinvest_Rec_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Reinvest_Rec_Detail_Sub.
11870           .frmMap_Reinvest_Rec_Detail_Sub.Form.FocusHolder.SetFocus
11880         Else
11890           .cmdCancel.SetFocus
11900         End If
11910       End With
11920     End Select
11930   End If

EXITP:
11940   KeyCode = intRetVal
11950   Exit Sub

ERRH:
11960   intRetVal = 0
11970   Select Case ERR.Number
        Case Else
11980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11990   End Select
12000   Resume EXITP

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

Private Sub cmbRevenueCodes_Purchase_display_GotFocus()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_GotFocus"

12210   With Me
12220     .cmbRevenueCodes_Purchase.SetFocus
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

Private Sub cmbRevenueCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_display_KeyDown"

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
12390         .cmbRevenueCodes_Purchase.SetFocus
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

Private Sub cmbRevenueCodes_Purchase_Enter()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_Enter"

12610   With Me
12620     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
12630       ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
12640     Case False
12650       If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
12660         ExpandCombo .cmbRevenueCodes_Purchase  ' ** Module Procedure: modUtilities.
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

Private Sub cmbRevenueCodes_Purchase_AfterUpdate()

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_AfterUpdate"

12810   With Me
12820     Select Case IsNull(.cmbRevenueCodes_Purchase)
          Case True
12830       .cmbRevenueCodes_Purchase = REVID_INC
12840     Case False
            ' ** Nothing at the moment.
12850     End Select
12860     If .cmbRevenueCodes_Purchase = REVID_INC Or .cmbRevenueCodes_Purchase = REVID_EXP Then
12870       .cmbRevenueCodes_Purchase_display = Null
12880     Else
12890       .cmbRevenueCodes_Purchase_display = .cmbRevenueCodes_Purchase.Column(1)
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

Private Sub cmbRevenueCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Purchase_KeyDown"

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
13090         If .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
13100           .cmbTaxCodes_Purchase.SetFocus
13110         Else
13120           .cmdOK.SetFocus
13130         End If
13140       End With
13150     End Select
13160   End If

        ' ** Shift keys.
13170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13180     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13190       With Me
13200         intRetVal = 0
13210         .description.SetFocus
13220       End With
13230     End Select
13240   End If

EXITP:
13250   KeyCode = intRetVal
13260   Exit Sub

ERRH:
13270   intRetVal = 0
13280   Select Case ERR.Number
        Case Else
13290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13300   End Select
13310   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Purchase_display_GotFocus()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_GotFocus"

13410   With Me
13420     .cmbTaxCodes_Purchase.SetFocus
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

Private Sub cmbTaxCodes_Purchase_display_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_display_KeyDown"

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
13590         .cmbTaxCodes_Purchase.SetFocus
13600       End With
13610     End Select
13620   End If

        ' ** Shift keys.
13630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13650       With Me
13660         intRetVal = 0
13670         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
13680           .cmbRevenueCodes_Purchase.SetFocus
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

Private Sub cmbTaxCodes_Purchase_Enter()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_Enter"

13910   With Me
13920     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
13930       ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
13940     Case False
13950       If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
13960         ExpandCombo .cmbTaxCodes_Purchase  ' ** Module Procedure: modUtilities.
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

Private Sub cmbTaxCodes_Purchase_AfterUpdate()

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_AfterUpdate"

14110   With Me
14120     Select Case IsNull(.cmbTaxCodes_Purchase)
          Case True
14130       .cmbTaxCodes_Purchase = TAXID_INC
14140     Case False
            ' ** Nothing at the moment.
14150     End Select
14160     If .cmbTaxCodes_Purchase = TAXID_INC Or .cmbTaxCodes_Purchase = TAXID_DED Then
14170       .cmbTaxCodes_Purchase_display = Null
14180     Else
14190       .cmbTaxCodes_Purchase_display = .cmbTaxCodes_Purchase.Column(1)
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

Private Sub cmbTaxCodes_Purchase_KeyDown(KeyCode As Integer, Shift As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Purchase_KeyDown"

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
14390         .cmdOK.SetFocus
14400       End With
14410     End Select
14420   End If

        ' ** Shift keys.
14430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14450       With Me
14460         intRetVal = 0
14470         If .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
14480           .cmbRevenueCodes_Purchase.SetFocus
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

Private Sub chkShowLocation_AfterUpdate()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowLocation_AfterUpdate"

14710   With Me
14720     Select Case .chkShowLocation
          Case True
14730       .chkShowLocation_lbl.FontBold = True
14740       .frmMap_Reinvest_Rec_Detail_Sub.Form.ShowLocation True  ' ** Form Procedure: frmMap_Reinvest_Rec_Detail_Sub.
14750       ShowLocation True  ' ** Procedure: Below
14760     Case False
14770       .chkShowLocation_lbl.FontBold = False
14780       .frmMap_Reinvest_Rec_Detail_Sub.Form.ShowLocation False  ' ** Form Procedure: frmMap_Reinvest_Rec_Detail_Sub.
14790       ShowLocation False  ' ** Procedure: Below.
14800     End Select
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

Private Sub chkShowLocation_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowLocation_KeyDown"

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
14990         If .cmdPrintReport.Enabled = True Then
15000           .cmdPrintReport.SetFocus
15010         Else
15020           .cmdOK.SetFocus
15030         End If
15040       End With
15050     End Select
15060   End If

        ' ** Shift keys.
15070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15080     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15090       With Me
15100         intRetVal = 0
15110         If .cmbTaxCodes_Purchase.Visible = True And .cmbTaxCodes_Purchase.Enabled = True Then
15120           .cmbTaxCodes_Purchase.SetFocus
15130         ElseIf .cmbRevenueCodes_Purchase.Visible = True And .cmbRevenueCodes_Purchase.Enabled = True Then
15140           .cmbRevenueCodes_Purchase.SetFocus
15150         Else
15160           .description.SetFocus
15170         End If
15180       End With
15190     End Select
15200   End If

EXITP:
15210   KeyCode = intRetVal
15220   Exit Sub

ERRH:
15230   intRetVal = 0
15240   Select Case ERR.Number
        Case Else
15250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15260   End Select
15270   Resume EXITP

End Sub

Private Function RecalcTots()
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
15330     Set rst = .frmMap_Reinvest_Rec_Detail_Sub.Form.RecordsetClone
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
15430           dblTmp02 = dblTmp02 + ![total_pcash]
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
15830               dblTmp01 = (.pershare * dblRate2)
15840               dblTmp01 = Round(dblTmp01, 5)
15850               .pershare_usd = dblTmp01
15860               If lngCurrID <> 150& Then
15870                 .total_cash_sum_usd.Visible = True
15880                 .pershare_usd.Visible = True
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
16070         .cmbRevenueCodes_Purchase_vline05.Visible = False
16080         .cmbRevenueCodes_Purchase_vline06.Visible = False
16090         .cmbRevenueCodes_Purchase_box.Width = (.cmbRevenueCodes_Purchase_vline06.Left - .cmbRevenueCodes_Purchase_box.Left)
16100         .pershare.Top = (.pershare_alt_line.Top + (2& * lngTpp))
16110         .pershare_lbl.Top = (.pershare_alt_line.Top + (3& * lngTpp))
16120       End If
16130       .total_cash_sum_usd.Visible = False
16140       .pershare_usd.Visible = False
16150     End Select
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

Private Sub IncludeCurrency_Sub()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngNewForm_Width As Long, lngWidth_Diff As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

16310   With Me
16320     If blnCurrID = True And lngCurrID <> 150& Then

16330       lngTmp01 = (.frmMap_Reinvest_Rec_Detail_Sub.Form.Location_ID.Width + (4& * lngTpp))
16340       lngTmp02 = (.frmMap_Reinvest_Rec_Detail_Sub.Form.total_pcash_usd.Width + (4& * lngTpp))

16350       Select Case .chkShowLocation
            Case True
16360         lngNewForm_Width = (lngForm_Width + lngTmp01 + lngTmp02)
16370         lngWidth_Diff = (lngNewForm_Width - lngForm_Width)
16380       Case False
16390         lngNewForm_Width = (lngForm_Width + lngTmp02)
16400         lngWidth_Diff = (lngNewForm_Width - lngForm_Width)
16410       End Select

16420       lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
16430       lngMonitorNum = 1&: lngTmp04 = 0&
16440       EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
16450       If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

16460       If lngMonitorNum = 1& Then lngTmp04 = lngFrm_Top
16470       DoCmd.MoveSize (lngFrm_Left - (lngWidth_Diff / 2)), lngTmp04, (lngFrm_Width + lngWidth_Diff), lngFrm_Height  'lngFrm_Top
16480       If lngMonitorNum > 1& Then
16490         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
16500       End If

16510       .Width = lngNewForm_Width
16520       .frmMap_Reinvest_Rec_Detail_Sub.Width = (lngSub_Width + lngWidth_Diff)
16530       .frmMap_Reinvest_Rec_Detail_Sub_box.Width = ((lngSub_Width + lngWidth_Diff) + (2& * lngTpp))

16540       .Header_vline01.Left = lngNewForm_Width
16550       .Header_vline02.Left = lngNewForm_Width
16560       .Detail_vline01.Left = lngNewForm_Width
16570       .Detail_vline02.Left = lngNewForm_Width
16580       .Footer_vline01.Left = lngNewForm_Width
16590       .Footer_vline02.Left = lngNewForm_Width
16600       .Header_hline01.Width = lngNewForm_Width
16610       .Header_hline02.Width = lngNewForm_Width
16620       .Detail_hline01.Width = lngNewForm_Width
16630       .Detail_hline02.Width = lngNewForm_Width
16640       .Footer_hline01.Width = lngNewForm_Width
16650       .Footer_hline02.Width = lngNewForm_Width

16660       .total_cash_sum_usd.Left = ((.total_cash_sum.Left + .total_cash_sum.Width) + (4& * lngTpp))
16670       .total_cash_sum_usd.Top = .total_cash_sum.Top
16680       .pershare_usd.Left = ((.total_cash_sum_usd.Left + .total_cash_sum_usd.Width) - .pershare_usd.Width)
16690       .pershare.Top = (.pershare_alt_line.Top + (2& * lngTpp))
16700       .pershare_usd.Top = (.pershare_alt_line.Top + (2& * lngTpp))
16710       .pershare_lbl.Visible = False

16720       lngTmp03 = (.pershare_lbl3.Top - .pershare_lbl2.Top)
16730       .pershare_lbl3.Top = (.pershare_alt_line.Top + (3& * lngTpp))
16740       .pershare_lbl2.Top = .pershare_lbl3.Top - lngTmp03
16750       .pershare_lbl2.Visible = True
16760       .pershare_lbl3.Visible = True

16770       .total_cash_sum_usd.ForeColor = CLR_DISABLED_FG
16780       .total_cash_sum_usd.BackColor = CLR_DISABLED_BG
16790       .pershare_usd.ForeColor = CLR_DISABLED_FG
16800       .pershare_usd.BackColor = CLR_DISABLED_BG

16810       .maptype.Left = (lngMapType_Left + lngWidth_Diff)
16820       .cmdCancel.Left = (lngCancel_Left + lngWidth_Diff)
16830       .cmdOK.Left = (lngOK_Left + lngWidth_Diff)

16840       .Width = lngNewForm_Width

16850       If lngMonitorCnt > 1& And blnIsOpen = False Then
16860         EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
16870         lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
16880         If lngMonitorNum > 1& Then
16890           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
16900         End If
16910       End If

16920     End If
16930   End With

EXITP:
16940   Exit Sub

ERRH:
16950   DoCmd.Hourglass False
16960   Select Case ERR.Number
        Case Else
16970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16980   End Select
16990   Resume EXITP

End Sub

Private Sub NoChar_Load()

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

17010   Set dbs = CurrentDb
17020   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
17030     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
17040     Set rst = qdf.OpenRecordset
17050     With rst
17060       .MoveLast
17070       lngNoChars = .RecordCount
17080       .MoveFirst
17090       arr_varNoChar = .GetRows(lngNoChars)
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
17100       .Close
17110     End With
17120     .Close
17130   End With

EXITP:
17140   Set rst = Nothing
17150   Set qdf = Nothing
17160   Set dbs = Nothing
17170   Exit Sub

ERRH:
17180   DoCmd.Hourglass False
17190   Select Case ERR.Number
        Case Else
17200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17210   End Select
17220   Resume EXITP

End Sub

Private Sub ShowLocation(blnShow)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowLocation"

        Dim lngNewForm_Width As Long, lngWidth_Diff As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long

17310   With Me

17320     lngTmp01 = (.frmMap_Reinvest_Rec_Detail_Sub.Form.Location_ID.Width + (4& * lngTpp))
17330     lngTmp02 = (.frmMap_Reinvest_Rec_Detail_Sub.Form.total_pcash_usd.Width + (4& * lngTpp))

17340     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
17350     lngMonitorNum = 1&: lngTmp03 = 0&
17360     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
17370     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

17380     Select Case blnShow
          Case True

17390       If blnCurrID = True And lngCurrID <> 150& Then
17400         lngNewForm_Width = (lngForm_Width + lngTmp01 + lngTmp02)
17410         lngWidth_Diff = (lngNewForm_Width - lngForm_Width)
17420       Else
17430         lngNewForm_Width = (lngForm_Width + lngTmp01)
17440         lngWidth_Diff = ((lngForm_Width + lngTmp01) - lngForm_Width)
17450       End If

17460       If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
17470       DoCmd.MoveSize (lngFrm_Left - (lngWidth_Diff / 2)), lngTmp03, (lngFrm_Width + lngWidth_Diff), lngFrm_Height  'lngFrm_Top
17480       If lngMonitorNum > 1& Then
17490         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
17500       End If

17510       .Width = (lngForm_Width + lngWidth_Diff)
17520       .frmMap_Reinvest_Rec_Detail_Sub.Width = (lngSub_Width + lngWidth_Diff)
17530       .frmMap_Reinvest_Rec_Detail_Sub_box.Width = ((lngSub_Width + lngWidth_Diff) + (2& * lngTpp))
17540       .cmdOK.Left = (lngOK_Left + lngWidth_Diff)
17550       .cmdCancel.Left = (lngCancel_Left + lngWidth_Diff)
17560       .maptype.Left = (lngMapType_Left + lngWidth_Diff)
17570       .Header_vline01.Left = .Width
17580       .Header_vline02.Left = .Width
17590       .Detail_vline01.Left = .Width
17600       .Detail_vline02.Left = .Width
17610       .Footer_vline01.Left = .Width
17620       .Footer_vline02.Left = .Width
17630       .Header_hline01.Width = .Width
17640       .Header_hline02.Width = .Width
17650       .Detail_hline01.Width = .Width
17660       .Detail_hline02.Width = .Width
17670       .Footer_hline01.Width = .Width
17680       .Footer_hline02.Width = .Width

17690     Case False

17700       If blnCurrID = True And lngCurrID <> 150& Then
17710         lngNewForm_Width = (lngForm_Width + lngTmp02)
17720         lngWidth_Diff = (lngNewForm_Width - lngForm_Width)
17730       Else
17740         lngNewForm_Width = lngForm_Width
17750         lngWidth_Diff = 0&
17760       End If

17770       .frmMap_Reinvest_Rec_Detail_Sub.Width = (lngSub_Width + lngWidth_Diff)
17780       .frmMap_Reinvest_Rec_Detail_Sub_box.Width = ((lngSub_Width + lngWidth_Diff) + (2& * lngTpp))
17790       .cmdOK.Left = (lngOK_Left + lngWidth_Diff)
17800       .cmdCancel.Left = (lngCancel_Left + lngWidth_Diff)
17810       .maptype.Left = (lngMapType_Left + lngWidth_Diff)
17820       .Header_vline01.Left = lngNewForm_Width
17830       .Header_vline02.Left = lngNewForm_Width
17840       .Detail_vline01.Left = lngNewForm_Width
17850       .Detail_vline02.Left = lngNewForm_Width
17860       .Footer_vline01.Left = lngNewForm_Width
17870       .Footer_vline02.Left = lngNewForm_Width
17880       .Header_hline01.Width = lngNewForm_Width
17890       .Header_hline02.Width = lngNewForm_Width
17900       .Detail_hline01.Width = lngNewForm_Width
17910       .Detail_hline02.Width = lngNewForm_Width
17920       .Footer_hline01.Width = lngNewForm_Width
17930       .Footer_hline02.Width = lngNewForm_Width
17940       .Width = lngNewForm_Width

17950       If blnCurrID = True And lngCurrID <> 150& Then
17960         If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
17970         DoCmd.MoveSize (lngFrm_Left - (lngWidth_Diff / 2)), lngTmp03, (lngFrm_Width + lngWidth_Diff), lngFrm_Height  'lngFrm_Top
17980         If lngMonitorNum > 1& Then
17990           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
18000         End If
18010       Else
18020         If lngMonitorNum = 1& Then lngTmp03 = lngFrm_Top
18030         DoCmd.MoveSize lngFrm_Left, lngTmp03, lngFrm_Width, lngFrm_Height  'lngFrm_Top
18040         If lngMonitorNum > 1& Then
18050           LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
18060         End If
18070       End If

18080     End Select
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
