VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMap_Split_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmMap_Split_Detail"

'VGC 08/30/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   IncludeCurrency_Sub()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Show Location:    L {chkShowLocation}
' **   Inc/Exp Codes:    N {cmbRevenueCodes_Deposit}
' **   OK:               O {cmdOK}
' **   Print Report:     P {cmdPrintReport}
' **   Tax Codes:        T {cmbTaxCodes_Deposit}

' ** Shortcut Ctrl keys responsive from this form:
' **   Comments:         M {description}

' ** Shortcut F-keys responsive from this form:
' **   Recalc Totals:    F5 {RecalcTots}

' ** MapType = "Deposit", except...
' ** MapType = "Withdrawn" for Reverse Split

'rptMap_Split
'qryMapReport_04

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

' ** Combo box column constants: curr_id.
'Private Const CBX_C_CURRID As Integer = 0  'curr_id
'Private Const CBX_C_CODE   As Integer = 1  'curr_code
'Private Const CBX_C_NAME   As Integer = 2  'curr_name
'Private Const CBX_C_SYM    As Integer = 3  'currsym_symbol
'Private Const CBX_C_DEC    As Integer = 4  'curr_decimal
'Private Const CBX_C_RATE1  As Integer = 5  'curr_rate1
'Private Const CBX_C_RATE2  As Integer = 6  'curr_rate2
Private Const CBX_C_DATE   As Integer = 7  'curr_date

' ** Array: arr_varNoChar().
Private lngNoChars As Long, arr_varNoChar As Variant
'Private Const N_CID   As Integer = 0
'Private Const N_CNAM  As Integer = 1
'Private Const N_FID   As Integer = 2
'Private Const N_FNAM  As Integer = 3
'Private Const N_NOCHR As Integer = 4

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngSub_Width As Long, lngOK_Left As Long, lngCancel_Left As Long, lngMapType_Left As Long
Private lngPrintReportLbl_Left As Long

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private strRateMsg As String
Private blnGoingToReport As Boolean
Private blnCurrID As Boolean, lngCurrID As Long
Private blnRecalcTots As Boolean
Private dblPerShare As Double, datTransDate As Date, lngTaxcode As Long, datAssetDate As Date
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long, lngOffset As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = Trim(.OpenArgs)
150           If CharCnt(strTmp01, "~") = 5& Then  ' ** Module Function: modStringFuncs.

160             blnIsOpen = True

170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             dblPerShare = Val(Left(strTmp01, (intPos01 - 1)))
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             datTransDate = CDate(Left(strTmp01, (intPos01 - 1)))
250             strTmp01 = Mid(strTmp01, (intPos01 + 1))
260             intPos01 = InStr(strTmp01, "~")
270             datAssetDate = CDate(Left(strTmp01, (intPos01 - 1)))
280             strTmp01 = Mid(strTmp01, (intPos01 + 1))
290             intPos01 = InStr(strTmp01, "~")
300             lngCurrID = Val(Left(strTmp01, (intPos01 - 1)))
310             strRateMsg = Mid(strTmp01, (intPos01 + 1))
                ' ** gstrAccountNo has specified account.

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
460             lngSub_Width = .frmMap_Split_Detail_Sub.Width
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

570             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

580             chkShowLocation_AfterUpdate  ' ** Procedure: Below.

590             Set dbs = CurrentDb
600             With dbs
610               Select Case strCallingForm
                  Case "frmJournal"
                    ' ** Journal Map, linked to Account, grouped and summed, by accountno.
620                 Set qdf = .QueryDefs("qryMap_Split_02")
630               Case "frmJournal_Columns"
                    ' ** Journal Map, linked to Account, grouped and summed, by accountno.
640                 Set qdf = .QueryDefs("qryJournal_Columns_34_05")
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
760               .Close
770             End With

780             Select Case strCallingForm
                Case "frmJournal"
790               If .frmMap_Split_Detail_Sub.Form.RecordSource <> "qryMap_Split_02" Then .frmMap_Split_Detail_Sub.Form.RecordSource = "qryMap_Split_02"
800             Case "frmJournal_Columns"
810               If .frmMap_Split_Detail_Sub.Form.RecordSource <> "qryJournal_Columns_34_05" Then .frmMap_Split_Detail_Sub.Form.RecordSource = "qryJournal_Columns_34_05"
820             End Select

830             Select Case .maptype
                Case "Deposit"
                  ' ** Heading stays as is.
840             Case "Withdrawn"
850               lngOffset = (18& * lngTpp)
860               .cmbRevenueCodes_Deposit_box_lbl.Caption = "Withdrawn Entries"
870               .cmbRevenueCodes_Deposit_box_lbl.LeftMargin = lngTpp
880               .cmbRevenueCodes_Deposit_box_lbl_dim_hi.Caption = "Withdrawn Entries"
890               .cmbRevenueCodes_Deposit_box_lbl_dim_hi.LeftMargin = lngTpp
900               .cmbRevenueCodes_Deposit_box_lbl.Width = .cmbRevenueCodes_Deposit_box_lbl.Width + lngOffset
910               .cmbRevenueCodes_Deposit_box_lbl_dim_hi.Width = .cmbRevenueCodes_Deposit_box_lbl_dim_hi.Width + lngOffset
920               .cmbRevenueCodes_Deposit_box2.Width = .cmbRevenueCodes_Deposit_box2.Width + lngOffset
930               .cmbRevenueCodes_Deposit_hline01.Width = .cmbRevenueCodes_Deposit_hline01.Width + lngOffset
940               .cmbRevenueCodes_Deposit_hline02.Width = .cmbRevenueCodes_Deposit_hline02.Width + lngOffset
950               .cmbRevenueCodes_Deposit_hline03.Width = .cmbRevenueCodes_Deposit_hline03.Width + lngOffset
960               .cmbRevenueCodes_Deposit_vline03.Left = .cmbRevenueCodes_Deposit_vline03.Left + lngOffset
970               .cmbRevenueCodes_Deposit_vline04.Left = .cmbRevenueCodes_Deposit_vline04.Left + lngOffset
980             End Select

990             .txtNumberRecords = "Number of Records = " & CStr(lngRecs)
1000            .pershare = Format(dblPerShare, "#,##0.000")    '3 dec.
1010            If .pershare > 0 Then
1020              .maptype = "Deposit"
1030              .description = CStr(dblPerShare) & "% Stock Split" & _
                    IIf(strRateMsg = vbNullString, strRateMsg, " (" & strRateMsg & ")") & ", Payable " & Format(datAssetDate, "mm/dd/yyyy")
1040            Else
1050              .maptype = "Withdrawn"
1060              .description = CStr(dblPerShare) & "% Reverse Stock Split" & ", Payable " & Format(datAssetDate, "mm/dd/yyyy")
1070            End If

1080            If .cmdOK.Enabled = True Then
1090              .cmdOK.SetFocus
1100            Else
1110              .description.SetFocus
1120            End If

1130            blnRecalcTots = True
1140            .TimerInterval = 500&

1150            DoEvents
1160            DoCmd.SelectObject acForm, THIS_NAME, False
1170            DoEvents

1180            blnIsOpen = False

1190          Else
1200            Cancel = -1
1210          End If
1220        Else
1230          Cancel = -1
1240        End If
1250      Else
1260        Cancel = -1
1270      End If
1280    End With

1290    If Cancel = -1 Then
1300      Beep
1310      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1320    End If

EXITP:
1330    Set rst = Nothing
1340    Set qdf = Nothing
1350    Set dbs = Nothing
1360    Exit Sub

ERRH:
1370    DoCmd.Hourglass False
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub Form_Load()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1510    With Me

1520      .total_shareface_sum.ForeColor = CLR_DISABLED_FG
1530      .total_shareface_sum.BackColor = CLR_DISABLED_BG
1540      .total_cash_sum.ForeColor = CLR_DISABLED_FG
1550      .total_cash_sum.BackColor = CLR_DISABLED_BG
1560      If .maptype = "Withdrawn" Then
1570        .pershare.ForeColor = CLR_DISABLED_FG
1580        .pershare.BackColor = CLR_LTRED
1590        .frmMap_Split_Detail_Sub.Form.total_icash_rev.Visible = True
1600        .frmMap_Split_Detail_Sub.Form.total_icash.Visible = False
1610      Else
1620        .pershare.ForeColor = CLR_DISABLED_FG
1630        .pershare.BackColor = CLR_DISABLED_BG
1640      End If

          ' ** Borrowing these variables from the Court Reports.
1650      If gstrCrtRpt_Version <> vbNullString Then
1660        .cusip = gstrCrtRpt_Ordinal
1670        .cusip.ForeColor = CLR_DISABLED_FG
1680        .cusip.BackColor = CLR_DISABLED_BG
1690        .cusip.Visible = True
1700        .totdesc = gstrCrtRpt_Version
1710        .totdesc.ForeColor = CLR_DISABLED_FG
1720        .totdesc.BackColor = CLR_DISABLED_BG
1730        .totdesc.Visible = True
1740      End If

1750      If gstrAccountNo <> vbNullString Then
            ' ** An account was specified.
1760        If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1770          lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1780        End If
1790        .cusip.Top = .totdesc.Top
            '.totdesc.Width = .shortname.Width
1800        .totdesc.Left = ((.cusip.Left + .cusip.Width) + (4& * lngTpp))
1810        .accountno = gstrAccountNo
1820        .shortname = gstrAccountName
1830        .accountno.Visible = True
1840        .shortname.Visible = True
1850      End If

1860      .transdate = datTransDate
1870      .assetdate = datAssetDate

1880      Select Case IsNull(.cmbRevenueCodes_Deposit)
          Case True
1890        .cmbRevenueCodes_Deposit = REVID_INC
1900        .cmbRevenueCodes_Deposit_display = Null
1910      Case False
1920        If .cmbRevenueCodes_Deposit = REVID_INC Or .cmbRevenueCodes_Deposit = REVID_EXP Then
1930          .cmbRevenueCodes_Deposit_display = Null
1940        Else
1950          .cmbRevenueCodes_Deposit_display = .cmbRevenueCodes_Deposit.Column(1)
1960        End If
1970      End Select

1980      Select Case IsNull(.cmbTaxCodes_Deposit)
          Case True
1990        .cmbTaxCodes_Deposit = TAXID_INC
2000        .cmbTaxCodes_Deposit_display = Null
2010      Case False
2020        If .cmbTaxCodes_Deposit = TAXID_INC Or .cmbTaxCodes_Deposit = TAXID_DED Then
2030          .cmbTaxCodes_Deposit_display = Null
2040        Else
2050          .cmbTaxCodes_Deposit_display = .cmbTaxCodes_Deposit.Column(1)
2060        End If
2070      End Select

2080      Select Case gblnRevenueExpenseTracking
          Case True
2090        .cmbRevenueCodes_Deposit.Visible = True
2100        .cmbRevenueCodes_Deposit_display.Visible = True
2110        .cmbRevenueCodes_Deposit_lbl2.Visible = False  ' ** Option off msg.
2120      Case False
2130        .cmbRevenueCodes_Deposit.Visible = False
2140        .cmbRevenueCodes_Deposit_display.Visible = False
2150        .cmbRevenueCodes_Deposit_lbl_box.Visible = False
2160        .cmbRevenueCodes_Deposit_lbl2.Visible = True  ' ** Option off msg.
2170      End Select

2180      Select Case gblnIncomeTaxCoding
          Case True
2190        .cmbTaxCodes_Deposit.Visible = True
2200        .cmbTaxCodes_Deposit_display.Visible = True
2210        .cmbTaxCodes_Deposit_lbl2.Visible = False  ' ** Option off msg.
2220      Case False
2230        .cmbTaxCodes_Deposit.Visible = False
2240        .cmbTaxCodes_Deposit_display.Visible = False
2250        .cmbTaxCodes_Deposit_lbl_box.Visible = False  ' ** Option off msg.
2260        .cmbTaxCodes_Deposit_lbl2.Visible = True
2270      End Select

2280      If gblnRevenueExpenseTracking = False And gblnIncomeTaxCoding = False Then
2290        .cmbRevenueCodes_Deposit_box_lbl.ForeColor = WIN_CLR_DISF
2300        .cmbRevenueCodes_Deposit_box_lbl_dim_hi.Visible = True
2310      End If

2320      IncludeCurrency False  ' ** Procedure: Below.

2330      DoCmd.Hourglass False
2340      .frmMap_Split_Detail_Sub.Form.TimerInterval = 500&
2350      blnRecalcTots = True
2360      .TimerInterval = 600&

2370    End With

EXITP:
2380    Exit Sub

ERRH:
2390    DoCmd.Hourglass False
2400    Select Case ERR.Number
        Case Else
2410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2420    End Select
2430    Resume EXITP

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

        ' ** Ctrl keys.
2640    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2650      Select Case intRetVal
          Case vbKeyM
2660        With Me
2670          intRetVal = 0
2680          .description.SetFocus
2690        End With
2700      End Select
2710    End If

        ' ** Ctrl-Shift keys.
2720    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2730      Select Case intRetVal
          Case vbKeyF
2740        With Me
2750          intRetVal = 0
2760          .FocusHolder.SetFocus
2770        End With
2780      End Select
2790    End If

EXITP:
2800    KeyCode = intRetVal
2810    Exit Sub

ERRH:
2820    intRetVal = 0
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub Form_Timer()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim blnRetVal As Boolean

2910    With Me
2920      .TimerInterval = 0&
2930      If gblnGoToReport = True Then
2940        Select Case blnGoingToReport
            Case True
2950          DoCmd.Hourglass False
2960          gblnGoToReport = False
2970          blnGoingToReport = False
2980          .GoToReport_arw_pr_img.Visible = False
2990          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
3000          .cmdCancel.SetFocus
3010          DoEvents
3020        Case False
3030          DoCmd.Hourglass True  ' ** Make sure it's still running.
3040          DoEvents
3050          blnGoingToReport = True
3060          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3070          .GoToReport_arw_pr_img.Visible = True
3080          If .cmdPrintReport.Enabled = True Then
3090            .cmdPrintReport.SetFocus
3100          End If
3110          DoEvents
3120          Beep
3130          DoCmd.Hourglass False
3140          DoEvents
3150          .TimerInterval = (GTR_WAIT * 2&)
3160        End Select
3170      End If
3180      If blnRecalcTots = True Then
3190        blnRecalcTots = False
3200        RecalcTots  ' ** Procedure: Below.
3210      End If
3220      If blnRecalcTots = True Then
3230        blnRecalcTots = False
3240        blnRetVal = RecalcTots  ' ** Function: Below.
3250        If blnRetVal = False Then
3260          blnRecalcTots = True
3270          .TimerInterval = 500&
3280        End If
3290      End If
3300    End With

EXITP:
3310    Exit Sub

ERRH:
3320    DoCmd.Hourglass False
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3410    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

3420    Set dbs = CurrentDb
        ' ** Empty tmpAccount.
3430    Set qdf = dbs.QueryDefs("qryMap_Split_08_01")
3440    qdf.Execute
3450    Set qdf = Nothing
3460    DoEvents
3470    dbs.Close
3480    Set dbs = Nothing
3490    DoEvents

        ' ** Assure all reports are closed.
3500    Do While Reports.Count > 0
3510      DoCmd.Close acReport, Reports(0).Name
3520    Loop

3530    gstrReportCallingForm = vbNullString

EXITP:
3540    Set qdf = Nothing
3550    Set dbs = Nothing
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
4140      lngRevCode = .cmbRevenueCodes_Deposit
4150      strRevDesc = .cmbRevenueCodes_Deposit.Column(1)
4160      lngRevType = .cmbRevenueCodes_Deposit.Column(2)
4170      lngTaxcode = .cmbTaxCodes_Deposit
4180      strTaxDesc = .cmbTaxCodes_Deposit.Column(1)
4190      lngTaxType = .cmbTaxCodes_Deposit.Column(2)
4200      varComments = .description
4210    End With

4220    Set dbs = CurrentDb
4230    With dbs

4240      Select Case strCallingForm
          Case "frmJournal"
            ' ** Append Journal Map to Journal, by specified [desc], [jsub], [revcod], [taxcod].
4250        Set qdf = .QueryDefs("qryMap_Split_03")
4260      Case "frmJournal_Columns"
            ' ** Append Journal Map to tblJournal_Column, by specified
            ' ** [desc], [jsub], [revcod], [taxcod], [revdesc], [taxdesc], [revtyp], [taxtyp].
4270        Set qdf = .QueryDefs("qryJournal_Columns_34_06")
4280        With qdf.Parameters
4290          ![revdesc] = strRevDesc
4300          ![taxdesc] = strTaxDesc
4310          ![revtyp] = lngRevType
4320          ![taxtyp] = lngTaxType
4330        End With
4340      End Select
4350      With qdf.Parameters
4360        ![Desc] = varComments
4370        ![jsub] = "StockSplit"  ' ** journalSubtype (VGC: 02/15/09).
4380        ![revcod] = lngRevCode
4390        ![taxcod] = lngTaxcode
4400      End With
4410      qdf.Execute

4420      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
4430        Set qdf = .QueryDefs("qryMap_01")
4440      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
4450        Set qdf = .QueryDefs("qryJournal_Columns_34_04")
4460      End Select
4470      qdf.Execute

4480      .Close
4490    End With

4500    Select Case strCallingForm
        Case "frmJournal"
4510      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4520        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
4530      Else
4540        Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.Requery
4550        DoCmd.SelectObject acForm, strCallingForm, False
4560      End If
4570      DoEvents
4580      Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.Map_NewRec  ' ** Form Procedure: frmJournal_Sub3_Purchase.
4590    Case "frmJournal_Columns"
4600      gblnSetFocus = True
4610      gstrReturningForm = "frmMap_Return"
4620      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
4630        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
4640      Else
4650        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
4660        Forms(strCallingForm).frmJournal_Columns_Sub.Form.RecalcTots  ' ** Form Procedure: frmJournal_Columns_Sub.
4670        DoCmd.SelectObject acForm, strCallingForm, False
4680        Forms(strCallingForm).TimerInterval = 250&
4690      End If
4700    End Select

4710    DoCmd.Close acForm, THIS_NAME

EXITP:
4720    DoCmd.Hourglass False
4730    Set qdf = Nothing
4740    Set dbs = Nothing
4750    Exit Sub

ERRH:
4760    DoCmd.Hourglass False
4770    Select Case ERR.Number
        Case Else
4780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4790    End Select
4800    Resume EXITP

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
5280        Case "frmJournal", "frmJournal_Columns"
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
5380      Select Case strCallingForm
          Case "frmJournal"
            ' ** Empty Journal Map.
5390        Set qdf = .QueryDefs("qryMap_01")
5400      Case "frmJournal_Columns"
            ' ** Empty Journal Map.
5410        Set qdf = .QueryDefs("qryJournal_Columns_34_04")
5420      End Select
5430      qdf.Execute
5440      .Close
5450    End With

5460    gblnSetFocus = True
5470    Select Case strCallingForm
        Case "frmJournal"
5480      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5490        DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
5500      Else
5510        gstrReturningForm = THIS_NAME
5520        Forms(strCallingForm).frmJournal_Sub3_Purchase.Form.Requery
5530        DoCmd.SelectObject acForm, strCallingForm, False
5540        Forms(strCallingForm).TimerInterval = 250&
5550      End If
5560    Case "frmJournal_Columns"
5570      If IsLoaded(strCallingForm, acForm) = False Then  ' ** Module Function: modFileUtilities.
5580        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
5590      Else
5600        gstrReturningForm = THIS_NAME
5610        Forms(strCallingForm).frmJournal_Columns_Sub.Form.Requery
5620        DoCmd.SelectObject acForm, strCallingForm, False
5630        Forms(strCallingForm).TimerInterval = 250&
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
5890          lngRecsCur = .frmMap_Split_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Split_Detail_Sub.
5900          If lngRecsCur > 0 Then
5910            .frmMap_Split_Detail_Sub.SetFocus
5920            .frmMap_Split_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmMap_Split_Detail_Sub.
5930            .frmMap_Split_Detail_Sub.Form.FocusHolder.SetFocus
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

        ' ** qryMapReport_04.
        'strSQL = "SELECT DISTINCTROW account.accountno, account.shortname, Sum([journal map].shareface) AS total_shareface, " & _
        '  "Sum([journal map].icash) AS total_icash, " & CoInfo & ", CStr([masterasset].[Description]) & " & _
        '  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy')) AS totdesc " & _
        '  "FROM ([journal map] INNER JOIN account ON [journal map].accountno = account.accountno) " & _
        '  "INNER JOIN masterasset ON [journal map].assetno = masterasset.assetno " & _
        '  "GROUP BY account.accountno, account.shortname, CStr([masterasset].[Description]) & " & _
        '  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
        '  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'));"

6610    With Me
6620      strDocName = "rptMap_Split"
6630      gstrReportCallingForm = THIS_NAME
6640      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
6650        .Modal = False
6660        DoCmd.OpenReport strDocName, acViewPreview, , , , acViewPreview
6670        DoCmd.Maximize
6680        DoCmd.RunCommand acCmdFitToWindow
6690        .Visible = False
6700      Else
            '##GTR_Ref: rptMap_Split
6710        DoCmd.OpenReport strDocName, acViewNormal, , , , acViewNormal
6720      End If
6730    End With

EXITP:
6740    Exit Sub

ERRH:
6750    Me.Visible = True
6760    DoCmd.Restore
6770    Select Case ERR.Number
        Case Else
6780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6790    End Select
6800    Resume EXITP

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
7070          .chkShowLocation.SetFocus
7080        End With
7090      End Select
7100    End If

EXITP:
7110    KeyCode = intRetVal
7120    Exit Sub

ERRH:
7130    intRetVal = 0
7140    Select Case ERR.Number
        Case Else
7150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7160    End Select
7170    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

7210    With Me
7220      If blnPrintReport_MouseDown = False Then
7230        Select Case blnPrintReport_Focus
            Case True
7240          .cmdPrintReport_raised_focus_dots_img.Visible = True
7250          .cmdPrintReport_raised_focus_img.Visible = False
7260        Case False
7270          .cmdPrintReport_raised_focus_img.Visible = True
7280          .cmdPrintReport_raised_focus_dots_img.Visible = False
7290        End Select
7300        .cmdPrintReport_raised_img.Visible = False
7310        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7320        .cmdPrintReport_sunken_focus_dots_img.Visible = False
7330        .cmdPrintReport_raised_img_dis.Visible = False
7340      End If
7350    End With

EXITP:
7360    Exit Sub

ERRH:
7370    Select Case ERR.Number
        Case Else
7380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7390    End Select
7400    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

7510    With Me
7520      .cmdPrintReport_raised_focus_dots_img.Visible = True
7530      .cmdPrintReport_raised_img.Visible = False
7540      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7550      .cmdPrintReport_raised_focus_img.Visible = False
7560      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7570      .cmdPrintReport_raised_img_dis.Visible = False
7580      blnPrintReport_MouseDown = False
7590    End With

EXITP:
7600    Exit Sub

ERRH:
7610    Select Case ERR.Number
        Case Else
7620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7630    End Select
7640    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

7710    With Me
7720      .cmdPrintReport_raised_img.Visible = True
7730      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
7740      .cmdPrintReport_raised_focus_img.Visible = False
7750      .cmdPrintReport_raised_focus_dots_img.Visible = False
7760      .cmdPrintReport_sunken_focus_dots_img.Visible = False
7770      .cmdPrintReport_raised_img_dis.Visible = False
7780      blnPrintReport_Focus = False
7790    End With

EXITP:
7800    Exit Sub

ERRH:
7810    Select Case ERR.Number
        Case Else
7820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7830    End Select
7840    Resume EXITP

End Sub

Private Sub total_shareface_sum_KeyDown(KeyCode As Integer, Shift As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "total_shareface_sum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7920    intShiftDown = (Shift And acShiftMask) > 0
7930    intAltDown = (Shift And acAltMask) > 0
7940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7970        With Me
7980          intRetVal = 0
7990          .description.SetFocus
8000        End With
8010      End Select
8020    End If

        ' ** Shift keys.
8030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8050        With Me
8060          intRetVal = 0
8070          lngRecsCur = .frmMap_Split_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Split_Detail_Sub.
8080          If lngRecsCur > 0& Then
8090            .frmMap_Split_Detail_Sub.SetFocus
8100            .frmMap_Split_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Split_Detail_Sub.
8110            .frmMap_Split_Detail_Sub.Form.FocusHolder.SetFocus
8120          Else
8130            .cmdCancel.SetFocus
8140          End If
8150        End With
8160      End Select
8170    End If

EXITP:
8180    KeyCode = intRetVal
8190    Exit Sub

ERRH:
8200    intRetVal = 0
8210    Select Case ERR.Number
        Case Else
8220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8230    End Select
8240    Resume EXITP

End Sub

Private Sub total_cash_sum_KeyDown(KeyCode As Integer, Shift As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "total_cash_sum_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
8370        With Me
8380          intRetVal = 0
8390          .description.SetFocus
8400        End With
8410      End Select
8420    End If

        ' ** Shift keys.
8430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8450        With Me
8460          intRetVal = 0
8470          lngRecsCur = .frmMap_Split_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Split_Detail_Sub.
8480          If lngRecsCur > 0& Then
8490            .frmMap_Split_Detail_Sub.SetFocus
8500            .frmMap_Split_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Split_Detail_Sub.
8510            .frmMap_Split_Detail_Sub.Form.FocusHolder.SetFocus
8520          Else
8530            .cmdCancel.SetFocus
8540          End If
8550        End With
8560      End Select
8570    End If

EXITP:
8580    KeyCode = intRetVal
8590    Exit Sub

ERRH:
8600    intRetVal = 0
8610    Select Case ERR.Number
        Case Else
8620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8630    End Select
8640    Resume EXITP

End Sub

Private Sub description_GotFocus()

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "description_GotFocus"

8710    With Me
8720      .description.SelLength = 0
8730      .description.SelStart = 999
8740    End With

EXITP:
8750    Exit Sub

ERRH:
8760    Select Case ERR.Number
        Case Else
8770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8780    End Select
8790    Resume EXITP

End Sub

Private Sub description_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "description_KeyDown"

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
8890          If .cmbRevenueCodes_Deposit.Visible = True And .cmbRevenueCodes_Deposit.Enabled = True Then
8900            .cmbRevenueCodes_Deposit.SetFocus
8910          ElseIf .cmbTaxCodes_Deposit.Visible = True And .cmbTaxCodes_Deposit.Enabled = True Then
8920            .cmbTaxCodes_Deposit.SetFocus
8930          Else
8940            .cmdOK.SetFocus
8950          End If
8960        End With
8970      End Select
8980    End If

        ' ** Shift keys.
8990    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9000      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9010        With Me
9020          intRetVal = 0
9030          lngRecsCur = .frmMap_Split_Detail_Sub.Form.RecCnt  ' ** Form Function: frmMap_Split_Detail_Sub.
9040          If lngRecsCur > 0& Then
9050            .frmMap_Split_Detail_Sub.SetFocus
9060            .frmMap_Split_Detail_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmMap_Split_Detail_Sub.
9070            .frmMap_Split_Detail_Sub.Form.FocusHolder.SetFocus
9080          Else
9090            .cmdCancel.SetFocus
9100          End If
9110        End With
9120      End Select
9130    End If

EXITP:
9140    KeyCode = intRetVal
9150    Exit Sub

ERRH:
9160    intRetVal = 0
9170    Select Case ERR.Number
        Case Else
9180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9190    End Select
9200    Resume EXITP

End Sub

Private Sub description_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "description_MouseMove"

9310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
9320    Exit Sub

ERRH:
9330    Select Case ERR.Number
        Case Else
9340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9350    End Select
9360    Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Deposit_display_GotFocus()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Deposit_display_GotFocus"

9410    With Me
9420      .cmbRevenueCodes_Deposit.SetFocus
9430    End With

EXITP:
9440    Exit Sub

ERRH:
9450    Select Case ERR.Number
        Case Else
9460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9470    End Select
9480    Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Deposit_display_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Deposit_display_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
9570        With Me
9580          intRetVal = 0
9590          .cmbRevenueCodes_Deposit.SetFocus
9600        End With
9610      End Select
9620    End If

        ' ** Shift keys.
9630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9650        With Me
9660          intRetVal = 0
9670          .description.SetFocus
9680        End With
9690      End Select
9700    End If

EXITP:
9710    KeyCode = intRetVal
9720    Exit Sub

ERRH:
9730    intRetVal = 0
9740    Select Case ERR.Number
        Case Else
9750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9760    End Select
9770    Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Deposit_Enter()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Deposit_Enter"

9810    With Me
9820      Select Case IsNull(.cmbRevenueCodes_Deposit)
          Case True
9830        ExpandCombo .cmbRevenueCodes_Deposit  ' ** Module Procedure: modUtilities.
9840      Case False
9850        If .cmbRevenueCodes_Deposit = REVID_INC Or .cmbRevenueCodes_Deposit = REVID_EXP Then
9860          ExpandCombo .cmbRevenueCodes_Deposit  ' ** Module Procedure: modUtilities.
9870        End If
9880      End Select
9890    End With

EXITP:
9900    Exit Sub

ERRH:
9910    Select Case ERR.Number
        Case Else
9920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9930    End Select
9940    Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Deposit_AfterUpdate()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Deposit_AfterUpdate"

10010   With Me
10020     Select Case IsNull(.cmbRevenueCodes_Deposit)
          Case True
10030       .cmbRevenueCodes_Deposit = REVID_INC
10040     Case False
            ' ** Nothing at the moment.
10050     End Select
10060     If .cmbRevenueCodes_Deposit = REVID_INC Or .cmbRevenueCodes_Deposit = REVID_EXP Then
10070       .cmbRevenueCodes_Deposit_display = Null
10080     Else
10090       .cmbRevenueCodes_Deposit_display = .cmbRevenueCodes_Deposit.Column(1)
10100     End If
10110   End With

EXITP:
10120   Exit Sub

ERRH:
10130   Select Case ERR.Number
        Case Else
10140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10150   End Select
10160   Resume EXITP

End Sub

Private Sub cmbRevenueCodes_Deposit_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbRevenueCodes_Deposit_KeyDown"

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
10290         If .cmbTaxCodes_Deposit.Visible = True And .cmbTaxCodes_Deposit.Enabled = True Then
10300           .cmbTaxCodes_Deposit.SetFocus
10310         Else
10320           .cmdOK.SetFocus
10330         End If
10340       End With
10350     End Select
10360   End If

        ' ** Shift keys.
10370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10390       With Me
10400         intRetVal = 0
10410         .description.SetFocus
10420       End With
10430     End Select
10440   End If

EXITP:
10450   KeyCode = intRetVal
10460   Exit Sub

ERRH:
10470   intRetVal = 0
10480   Select Case ERR.Number
        Case Else
10490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10500   End Select
10510   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Deposit_display_GotFocus()

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Deposit_display_GotFocus"

10610   With Me
10620     .cmbTaxCodes_Deposit.SetFocus
10630   End With

EXITP:
10640   Exit Sub

ERRH:
10650   Select Case ERR.Number
        Case Else
10660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10670   End Select
10680   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Deposit_display_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Deposit_display_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10720   intShiftDown = (Shift And acShiftMask) > 0
10730   intAltDown = (Shift And acAltMask) > 0
10740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10770       With Me
10780         intRetVal = 0
10790         .cmbTaxCodes_Deposit.SetFocus
10800       End With
10810     End Select
10820   End If

        ' ** Shift keys.
10830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10850       With Me
10860         intRetVal = 0
10870         If .cmbRevenueCodes_Deposit.Visible = True And .cmbRevenueCodes_Deposit.Enabled = True Then
10880           .cmbRevenueCodes_Deposit.SetFocus
10890         Else
10900           .description.SetFocus
10910         End If
10920       End With
10930     End Select
10940   End If

EXITP:
10950   KeyCode = intRetVal
10960   Exit Sub

ERRH:
10970   intRetVal = 0
10980   Select Case ERR.Number
        Case Else
10990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11000   End Select
11010   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Deposit_Enter()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Deposit_Enter"

11110   With Me
11120     Select Case IsNull(.cmbTaxCodes_Deposit)
          Case True
11130       ExpandCombo .cmbTaxCodes_Deposit  ' ** Module Procedure: modUtilities.
11140     Case False
11150       If .cmbTaxCodes_Deposit = TAXID_INC Or .cmbTaxCodes_Deposit = TAXID_DED Then
11160         ExpandCombo .cmbTaxCodes_Deposit  ' ** Module Procedure: modUtilities.
11170       End If
11180     End Select
11190   End With

EXITP:
11200   Exit Sub

ERRH:
11210   Select Case ERR.Number
        Case Else
11220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11230   End Select
11240   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Deposit_AfterUpdate()

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Deposit_AfterUpdate"

11310   With Me
11320     Select Case IsNull(.cmbTaxCodes_Deposit)
          Case True
11330       .cmbTaxCodes_Deposit = TAXID_INC
11340     Case False
            ' ** Nothing at the moment.
11350     End Select
11360     If .cmbTaxCodes_Deposit = TAXID_INC Or .cmbTaxCodes_Deposit = TAXID_DED Then
11370       .cmbTaxCodes_Deposit_display = Null
11380     Else
11390       .cmbTaxCodes_Deposit_display = .cmbTaxCodes_Deposit.Column(1)
11400     End If
11410   End With

EXITP:
11420   Exit Sub

ERRH:
11430   Select Case ERR.Number
        Case Else
11440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11450   End Select
11460   Resume EXITP

End Sub

Private Sub cmbTaxCodes_Deposit_KeyDown(KeyCode As Integer, Shift As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbTaxCodes_Deposit_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11520   intShiftDown = (Shift And acShiftMask) > 0
11530   intAltDown = (Shift And acAltMask) > 0
11540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11570       With Me
11580         intRetVal = 0
11590         .cmdOK.SetFocus
11600       End With
11610     End Select
11620   End If

        ' ** Shift keys.
11630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11650       With Me
11660         intRetVal = 0
11670         If .cmbRevenueCodes_Deposit.Visible = True And .cmbRevenueCodes_Deposit.Enabled = True Then
11680           .cmbRevenueCodes_Deposit.SetFocus
11690         Else
11700           .description.SetFocus
11710         End If
11720       End With
11730     End Select
11740   End If

EXITP:
11750   KeyCode = intRetVal
11760   Exit Sub

ERRH:
11770   intRetVal = 0
11780   Select Case ERR.Number
        Case Else
11790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11800   End Select
11810   Resume EXITP

End Sub

Private Sub curr_id_cmd_DblClick(Cancel As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "curr_id_cmd_DblClick"

11910   With Me
11920     Select Case .curr_date.Visible
          Case True
11930       .curr_date.Visible = False
11940     Case False
11950       .curr_date = .curr_id.Column(CBX_C_DATE)
11960       .curr_date.Visible = True
11970     End Select
11980   End With

EXITP:
11990   Exit Sub

ERRH:
12000   Select Case ERR.Number
        Case Else
12010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
12020   End Select
12030   Resume EXITP

End Sub

Private Sub chkShowLocation_AfterUpdate()

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowLocation_AfterUpdate"

12110   With Me
12120     Select Case .chkShowLocation
          Case True
12130       .chkShowLocation_lbl.FontBold = True
12140       .frmMap_Split_Detail_Sub.Form.ShowLocation True  ' ** Form Procedure: frmMap_Split_Detail_Sub.
12150       ShowLocation True  ' ** Procedure: Below
12160     Case False
12170       .chkShowLocation_lbl.FontBold = False
12180       .frmMap_Split_Detail_Sub.Form.ShowLocation False  ' ** Form Procedure: frmMap_Split_Detail_Sub.
12190       ShowLocation False  ' ** Procedure: Below.
12200     End Select
12210   End With

EXITP:
12220   Exit Sub

ERRH:
12230   Select Case ERR.Number
        Case Else
12240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12250   End Select
12260   Resume EXITP

End Sub

Private Sub chkShowLocation_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowLocation_KeyDown"

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
12390         If .cmdPrintReport.Enabled = True Then
12400           .cmdPrintReport.SetFocus
12410         Else
12420           .cmdOK.SetFocus
12430         End If
12440       End With
12450     End Select
12460   End If

        ' ** Shift keys.
12470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12490       With Me
12500         intRetVal = 0
12510         If .cmbTaxCodes_Deposit.Visible = True And .cmbTaxCodes_Deposit.Enabled = True Then
12520           .cmbTaxCodes_Deposit.SetFocus
12530         ElseIf .cmbRevenueCodes_Deposit.Visible = True And .cmbRevenueCodes_Deposit.Enabled = True Then
12540           .cmbRevenueCodes_Deposit.SetFocus
12550         Else
12560           .description.SetFocus
12570         End If
12580       End With
12590     End Select
12600   End If

EXITP:
12610   KeyCode = intRetVal
12620   Exit Sub

ERRH:
12630   intRetVal = 0
12640   Select Case ERR.Number
        Case Else
12650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12660   End Select
12670   Resume EXITP

End Sub

Public Function RecalcTots() As Boolean
' ** I give up trying to get it from the subform!

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots"

        Dim rst As DAO.Recordset
        Dim lngRecs As Long
        Dim dblTmp01 As Double, dblTmp02 As Double
        Dim lngX As Long
        Dim blnRetVal As Boolean

12710   With Me
12720     blnRetVal = False
12730     Set rst = .frmMap_Split_Detail_Sub.Form.RecordsetClone
12740     With rst
12750       If .BOF = True And .EOF = True Then
              ' ** Ho-Hum...
12760       Else
12770         .MoveLast
12780         lngRecs = .RecordCount
12790         .MoveFirst
12800         dblTmp01 = 0#: dblTmp02 = 0#
12810         For lngX = 1& To lngRecs
12820           dblTmp01 = dblTmp01 + ![total_shareface]
12830           dblTmp02 = dblTmp02 + ![total_icash]
12840           If lngX < lngRecs Then .MoveNext
12850         Next
12860       End If
12870       .Close
12880     End With
12890     If dblTmp01 <> 0# Then
12900       blnRetVal = True
12910     End If
12920     .total_shareface_sum = dblTmp01
12930     .total_cash_sum = Abs(dblTmp02)
12940   End With

EXITP:
12950   Set rst = Nothing
12960   RecalcTots = blnRetVal
12970   Exit Function

ERRH:
12980   blnRetVal = False
12990   Select Case ERR.Number
        Case Else
13000     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13010   End Select
13020   Resume EXITP

End Function

Public Sub IncludeCurrency(blnCalc As Boolean)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

13110   With Me
13120     Select Case blnCalc
          Case True
            ' ** Nothing.
13130     Case False
13140       .curr_id.Enabled = False
13150       .curr_id.Locked = True
13160       .curr_id.ForeColor = CLR_DISABLED_FG
13170       .curr_id.BackColor = CLR_VLTGRN
13180       If blnCurrID = True And lngCurrID <> 150& Then
13190         .curr_id.Visible = True
13200         .curr_id_cmd.Visible = True
13210         .curr_id_cmd.Enabled = True
13220       Else
13230         .curr_id.Visible = False
13240         .curr_id_cmd.Visible = False
13250         .curr_id_cmd.Enabled = False
13260       End If
13270     End Select
13280   End With

EXITP:
13290   Exit Sub

ERRH:
13300   DoCmd.Hourglass False
13310   Select Case ERR.Number
        Case Else
13320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13330   End Select
13340   Resume EXITP

End Sub

Private Sub IncludeCurrency_Sub()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency_Sub"

        Dim lngTmp01 As Long, lngTmp02 As Long

13410   With Me

13420     With .frmMap_Split_Detail_Sub.Form
13430       lngTmp01 = .Width - .Location_ID_alt_line.Left
13440     End With

13450     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
13460     lngMonitorNum = 1&: lngTmp02 = 0&
13470     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
13480     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

13490     If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
13500     DoCmd.MoveSize (lngFrm_Left - (lngTmp01 / 2)), lngTmp02, (lngFrm_Width + lngTmp01), lngFrm_Height  'lngFrm_Top
13510     If lngMonitorNum > 1& Then
13520       LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
13530     End If

13540     .Width = (.Width + lngTmp01)
13550     .frmMap_Split_Detail_Sub.Width = (.frmMap_Split_Detail_Sub.Width + lngTmp01)
13560     .frmMap_Split_Detail_Sub_box.Width = (.frmMap_Split_Detail_Sub_box.Width + lngTmp01)

13570     .Header_vline01.Left = .Width
13580     .Header_vline02.Left = .Width
13590     .Detail_vline01.Left = .Width
13600     .Detail_vline02.Left = .Width
13610     .Footer_vline01.Left = .Width
13620     .Footer_vline02.Left = .Width
13630     .Header_hline01.Width = .Width
13640     .Header_hline02.Width = .Width
13650     .Detail_hline01.Width = .Width
13660     .Detail_hline02.Width = .Width
13670     .Footer_hline01.Width = .Width
13680     .Footer_hline02.Width = .Width

13690     .maptype.Left = (.maptype.Left + lngTmp01)
13700     .cmdCancel.Left = (.cmdCancel.Left + lngTmp01)
13710     .cmdOK.Left = (.cmdOK.Left + lngTmp01)

13720     If blnCurrID = True And lngCurrID <> 150& Then
13730       .curr_id.Visible = True
13740     Else
13750       .curr_id.Visible = False
13760       .cmbRevenueCodes_Deposit_vline05.Visible = False
13770       .cmbRevenueCodes_Deposit_vline06.Visible = False
13780       .cmbRevenueCodes_Deposit_box.Width = (.cmbRevenueCodes_Deposit_vline06.Left - .cmbRevenueCodes_Deposit_box.Left)
13790     End If

13800   End With

EXITP:
13810   Exit Sub

ERRH:
13820   DoCmd.Hourglass False
13830   Select Case ERR.Number
        Case Else
13840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13850   End Select
13860   Resume EXITP

End Sub

Private Sub NoChar_Load()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "NoChar_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

13910   Set dbs = CurrentDb
13920   With dbs
          ' ** tblCurrency_Symbol_Font3, linked to tblCurrency_Symbol, tblFontName,
          ' ** just 'Fixedsys', currfont3_nochar = True.
13930     Set qdf = .QueryDefs("qryCurrency_Font_Symbol_04")
13940     Set rst = qdf.OpenRecordset
13950     With rst
13960       .MoveLast
13970       lngNoChars = .RecordCount
13980       .MoveFirst
13990       arr_varNoChar = .GetRows(lngNoChars)
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
14000       .Close
14010     End With
14020     .Close
14030   End With

EXITP:
14040   Set rst = Nothing
14050   Set qdf = Nothing
14060   Set dbs = Nothing
14070   Exit Sub

ERRH:
14080   DoCmd.Hourglass False
14090   Select Case ERR.Number
        Case Else
14100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14110   End Select
14120   Resume EXITP

End Sub

Private Sub ShowLocation(blnShow)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "ShowLocation"

        Dim lngNewForm_Width As Long, lngWidth_Diff As Long
        Dim lngTmp01 As Long, lngTmp02 As Long

14210   With Me

14220     lngTmp01 = (.frmMap_Split_Detail_Sub.Form.Location_ID.Width + (4& * lngTpp))

14230     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
14240     lngMonitorNum = 1&: lngTmp02 = 0&
14250     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
14260     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

14270     Select Case blnShow
          Case True

14280       lngNewForm_Width = (lngForm_Width + lngTmp01)
14290       lngWidth_Diff = ((lngForm_Width + lngTmp01) - lngForm_Width)

14300       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
14310       DoCmd.MoveSize (lngFrm_Left - (lngWidth_Diff / 2)), lngTmp02, (lngFrm_Width + lngWidth_Diff), lngFrm_Height  'lngFrm_Top
14320       If lngMonitorNum > 1& Then
14330         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
14340       End If

14350       .Width = (lngForm_Width + lngWidth_Diff)
14360       .frmMap_Split_Detail_Sub.Width = (lngSub_Width + lngWidth_Diff)
14370       .frmMap_Split_Detail_Sub_box.Width = ((lngSub_Width + lngWidth_Diff) + (2& * lngTpp))
14380       .cmdOK.Left = (lngOK_Left + lngWidth_Diff)
14390       .cmdCancel.Left = (lngCancel_Left + lngWidth_Diff)
14400       .maptype.Left = (lngMapType_Left + lngWidth_Diff)
14410       .Header_vline01.Left = .Width
14420       .Header_vline02.Left = .Width
14430       .Detail_vline01.Left = .Width
14440       .Detail_vline02.Left = .Width
14450       .Footer_vline01.Left = .Width
14460       .Footer_vline02.Left = .Width
14470       .Header_hline01.Width = .Width
14480       .Header_hline02.Width = .Width
14490       .Detail_hline01.Width = .Width
14500       .Detail_hline02.Width = .Width
14510       .Footer_hline01.Width = .Width
14520       .Footer_hline02.Width = .Width

14530     Case False

14540       lngNewForm_Width = lngForm_Width
14550       lngWidth_Diff = 0&
14560       .frmMap_Split_Detail_Sub.Width = (lngSub_Width + lngWidth_Diff)
14570       .frmMap_Split_Detail_Sub_box.Width = ((lngSub_Width + lngWidth_Diff) + (2& * lngTpp))
14580       .cmdOK.Left = (lngOK_Left + lngWidth_Diff)
14590       .cmdCancel.Left = (lngCancel_Left + lngWidth_Diff)
14600       .maptype.Left = (lngMapType_Left + lngWidth_Diff)
14610       .Header_vline01.Left = lngNewForm_Width
14620       .Header_vline02.Left = lngNewForm_Width
14630       .Detail_vline01.Left = lngNewForm_Width
14640       .Detail_vline02.Left = lngNewForm_Width
14650       .Footer_vline01.Left = lngNewForm_Width
14660       .Footer_vline02.Left = lngNewForm_Width
14670       .Header_hline01.Width = lngNewForm_Width
14680       .Header_hline02.Width = lngNewForm_Width
14690       .Detail_hline01.Width = lngNewForm_Width
14700       .Detail_hline02.Width = lngNewForm_Width
14710       .Footer_hline01.Width = lngNewForm_Width
14720       .Footer_hline02.Width = lngNewForm_Width
14730       .Width = lngNewForm_Width

14740       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
14750       DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, lngFrm_Height  'lngFrm_Top
14760       If lngMonitorNum > 1& Then
14770         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
14780       End If

14790     End Select
14800   End With

EXITP:
14810   Exit Sub

ERRH:
14820   Select Case ERR.Number
        Case Else
14830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14840   End Select
14850   Resume EXITP

End Sub
