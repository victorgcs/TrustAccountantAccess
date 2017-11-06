VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmJournal_Columns_TaxLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmJournal_Columns_TaxLot"

'VGC 03/22/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   chkIncludeCurrency_AfterUpdate()
' #######################################

'##########################################
'CURRENCY NOT FINISHED!
'##########################################

' ** Shortcut Alt keys responsive from this form:
' **   Average:          A {opgChoice_optAverage}
' **   Cancel:           C {cmdCancel}
' **   FIFO:             F {opgChoice_optFIFO}
' **   High:             H {opgChoice_optHigh}
' **   Low:              L {opgChoice_optLow}
' **   OK:               O {cmdOK}
' **   Print:            P {cmdPrintReport}
' **   Shareface:        R {shareface on frmJournal_Columns_TaxLot_Sub}
' **   Specific:         S {opgChoice_optSpecific}
' **   Location:         T {Loc_Name_display on frmJournal_Columns_TaxLot_Sub}
' **   Include Currency: U {chkIncludeCurrency}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmJournal_Columns_TaxLot_Sub}

' ** Note: Cost calculation is based on the Cost showing in ActiveAssets, not MasterAsset.

' **************************************
' ** VGC 01/13/2013:
' ** NOTE REQUIRED ZZ_'S IN THIS MOD:
' **   zz_qry_System_46_01
' **   zz_qry_System_46_02
' **************************************

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private varRemaining As Variant, varShareface As Variant, varPCash As Variant, varICash As Variant, varCost As Variant
Private dblShareTot As Double, dblICashTot As Double, dblPCashTot As Double, dblCostTot As Double
Private dblShareNew As Double, dblICashNew As Double, dblPCashNew As Double, dblCostNew As Double
Private lngRevcodeID As Long, lngTaxcode As Long, strDesc As String

' ** Array: arr_varLot().
Private lngLots As Long, arr_varLot() As Variant
Private Const L_ELEMS As Integer = 23  ' ** Array's first-element UBound().
Private Const L_LOTNO As Integer = 0
Private Const L_JID   As Integer = 1
Private Const L_ACTNO As Integer = 2
Private Const L_ASTNO As Integer = 3
Private Const L_SHARE As Integer = 4
Private Const L_SCENT As Integer = 5
Private Const L_SNEW  As Integer = 6
Private Const L_SADD  As Integer = 7
Private Const L_ICASH As Integer = 8
Private Const L_ICENT As Integer = 9
Private Const L_INEW  As Integer = 10
Private Const L_IADD  As Integer = 11
Private Const L_PCASH As Integer = 12
Private Const L_PCENT As Integer = 13
Private Const L_PNEW  As Integer = 14
Private Const L_PADD  As Integer = 15
Private Const L_COST  As Integer = 16
Private Const L_CCENT As Integer = 17
Private Const L_CNEW  As Integer = 18
Private Const L_CADD  As Integer = 19
Private Const L_REM   As Integer = 20
'Private Const L_CURID As Integer = 21
'Private Const L_USD   As Integer = 22
'Private Const L_UNEW  As Integer = 23

' ** Array: arr_varRec().
Private lngRecs As Long, arr_varRec() As Variant
Private Const R_ELEMS As Integer = 6  ' ** Array's first-element UBound().
Private Const R_TID   As Integer = 0
Private Const R_ASTNO As Integer = 1
Private Const R_SHRS  As Integer = 2
Private Const R_ADAT  As Integer = 3
Private Const R_ZERO  As Integer = 4
Private Const R_COST  As Integer = 5
Private Const R_PPU   As Integer = 6

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngSub_Width As Long, lngDetailLine_Width As Long
Private lngCancel_Left As Long, lngOK_Left As Long, lngFocusHolder_Left As Long, lngRecCnt_Left As Long
Private lngCurrID_Width As Long, lngCostUSD_Width As Long, lngIncludeCurrency_Left As Long, lngChkBoxLbl_Offset As Long

Private intOpeningChoice As Integer, intLastChoice As Integer, blnFromSaleBtn As Boolean
Private blnGoingToReport As Boolean, blnGoneToReport As Boolean, lngPrintReportLbl_Left As Long, blnWasGTR As Boolean
Private blnMultiLotSale As Boolean, blnHasForEx As Boolean
Private lngLotCnt As Long, strAccountNo As String, lngAssetNo As Long, lngCurrID As Long
Private lngJrnlColID As Long, dblShareface As Double, datAssetDateSale As Date, blnD4D As Boolean
Private dblPCash As Double, dblCost As Double
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long ', lngRecsCur As Long
Private THAT_PROC As String, That_Erl As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, rst As DAO.Recordset, qdf As DAO.QueryDef
        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = Trim(.OpenArgs)
150           If CharCnt(strTmp01, "~") = 4 Then  ' ** Module Function: modStringFuncs.
                ' ** strCallingForm ~ JrnlCol_ID ~ ShareFace ~ AssetDate ~ Dollar-for-Dollar

160             blnIsOpen = True

170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             lngJrnlColID = CLng(Left(strTmp01, (intPos01 - 1)))
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             dblShareface = CDbl(Left(strTmp01, (intPos01 - 1)))
250             strTmp01 = Mid(strTmp01, (intPos01 + 1))
260             intPos01 = InStr(strTmp01, "~")
270             blnD4D = CBool(Mid(strTmp01, (intPos01 + 1)))
280             strTmp01 = Left(strTmp01, (intPos01 - 1))

290             If strTmp01 <> "Null" Then
300               datAssetDateSale = CDate(strTmp01)  ' ** This is the assetdate the user entered for the Sale.
310             End If

320             CLR_DISABLED_FG = CLR_DKGRY
330             CLR_DISABLED_BG = CLR_LTTEAL

340             blnWasGTR = gblnGoToReport
350             gstrFormQuerySpec = THIS_NAME
360             blnMultiLotSale = False
370             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
380             lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
390             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

400             lngForm_Width = .Width
410             lngSub_Width = .frmJournal_Columns_TaxLot_Sub.Width
420             lngDetailLine_Width = .Detail_hline01.Width
430             lngCancel_Left = .cmdCancel.Left
440             lngOK_Left = .cmdOK.Left
450             lngFocusHolder_Left = .FocusHolder.Left
460             lngRecCnt_Left = .RecordCount_lbl_box.Left
470             lngCurrID_Width = (.frmJournal_Columns_TaxLot_Sub.Form.curr_id.Width + (4& * lngTpp))
480             lngCostUSD_Width = (.frmJournal_Columns_TaxLot_Sub.Form.cost_usd.Width + (4& * lngTpp))
490             lngIncludeCurrency_Left = .chkIncludeCurrency.Left
500             lngChkBoxLbl_Offset = .chkIncludeCurrency_lbl.Left - .chkIncludeCurrency.Left

                ' ** Variables are fed empty, then populated ByRef.
510             GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

520             FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

530             lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
540             If lngMonitorCnt > 1& Then
550               EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
560               lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
570               If lngMonitorNum > 1& Then
580                 LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
590               End If
600             End If

610             Set dbs = CurrentDb
620             With dbs
                  ' ** Table tmpEdit04, all fields.
630               Set qdf = .QueryDefs("qryLotInformation_27b")
640               Set rst = qdf.OpenRecordset
650               With rst
660                 If .BOF = True And .EOF = True Then
670                   lngLotCnt = 0&
680                 Else
690                   .MoveLast
700                   lngLotCnt = .RecordCount
710                   strAccountNo = ![accountno]
720                   lngAssetNo = ![assetno]
730                   lngCurrID = ![curr_id]
740                 End If
750                 .Close
760               End With
770               .Close
780             End With

790             If blnD4D = True Then
800               .opgChoice_optAverage.Enabled = False
810             End If
820             .ShareFaceSale = dblShareface

830             If strTmp01 <> "Null" Then
840               .AssetDateSale = datAssetDateSale
850               .AssetDateNull = False
860             Else
870               .AssetDateNull = True
880             End If

890             If lngCurrID <> 150& Then
900               .frmJournal_Columns_TaxLot_Sub.Form.Cost.Visible = False
910               .frmJournal_Columns_TaxLot_Sub.Form.cost_str.Visible = True
920               .frmJournal_Columns_TaxLot_Sub.Form.cost_sum_box.Visible = True
930               .frmJournal_Columns_TaxLot_Sub.Form.priceperunit_avg_box.Visible = True
940             Else
950               .frmJournal_Columns_TaxLot_Sub.Form.Cost.Visible = True
960               .frmJournal_Columns_TaxLot_Sub.Form.cost_str.Visible = False
970               .frmJournal_Columns_TaxLot_Sub.Form.cost_sum_box.Visible = False
980               .frmJournal_Columns_TaxLot_Sub.Form.priceperunit_avg_box.Visible = False
990             End If

1000            .RecordCount_lbl.Caption = "Lot Count = " & CStr(lngLotCnt)

1010            blnIsOpen = False

1020          Else
1030            Cancel = -1
1040          End If
1050        Else
1060          Cancel = -1
1070        End If
1080      Else
1090        Cancel = -1
1100      End If
1110    End With

1120    If Cancel = -1 Then
1130      DoCmd.Hourglass False
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

        Dim rst As DAO.Recordset
        Dim lngX As Long

1310    With Me

1320      .opgChoice = .opgChoice_optFIFO.OptionValue  ' ** Default to FIFO.

1330      Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1340      Set rst = .frmJournal_Columns_TaxLot_Sub.Form.RecordsetClone
1350      rst.MoveFirst
          ' ** VGC 10/24/2014: If any are averaged, all should be, so taking just 1st record should be OK.
1360      If rst![IsAverage] = True Then
1370        .opgChoice = .opgChoice_optAverage.OptionValue
1380        .opgChoice_optFIFO.Enabled = False
1390        .opgChoice_optAverage.Enabled = True
1400        .opgChoice_optSpecific.Enabled = False
1410        .opgChoice_optHigh.Enabled = False
1420        .opgChoice_optLow.Enabled = False
1430        With .frmJournal_Columns_TaxLot_Sub.Form
1440          .averagepriceperunit.Visible = True
1450          .averagepriceperunit_lbl.Visible = True
1460          .averagepriceperunit_lbl_line.Visible = True
1470          .averagepriceperunit_avg.Visible = True
1480          .priceperunit.Visible = False
1490          .priceperunit_lbl.Visible = False
1500          .priceperunit_lbl_line.Visible = False
1510          .priceperunit_avg.Visible = False
1520        End With
            ' ** CLR_DKGRY  = 92/92/92  Average ForeColor
            ' ** CLR_VDKGRY = 46/46/46  Normal ForeColor
            '.opgChoice_optFIFO_lbl.ForeColor = CLR_DKGRY
            '.opgChoice_optSpecific_lbl.ForeColor = CLR_DKGRY
            '.opgChoice_optHigh_lbl.ForeColor = CLR_DKGRY
            '.opgChoice_optLow_lbl.ForeColor = CLR_DKGRY
1530      End If
1540      rst.Close
1550      Set rst = Nothing

1560      .shortname.ForeColor = CLR_DISABLED_FG
1570      .shortname.BackColor = CLR_DISABLED_BG
1580      .totdesc.ForeColor = CLR_DISABLED_FG
1590      .totdesc.BackColor = CLR_DISABLED_BG
1600      .cusip.ForeColor = CLR_DISABLED_FG
1610      .cusip.BackColor = CLR_DISABLED_BG
1620      .assettype_description.ForeColor = CLR_DISABLED_FG
1630      .assettype_description.BackColor = CLR_DISABLED_BG
1640      .ShareFaceSale.ForeColor = CLR_DISABLED_FG
1650      .ShareFaceSale.BackColor = CLR_DISABLED_BG

1660      opgChoice_AfterUpdate  ' ** Procedure: Below.
1670      DoEvents

1680      blnHasForEx = False
1690      RecArray_Load  ' ** Procedure: Below.

1700      If .frmJournal_Columns_TaxLot_Sub.Form.shareface = 0 Then
1710        For lngX = 0& To (lngRecs - 1&)
1720          If arr_varRec(R_ZERO, lngX) = False Then
1730            .frmJournal_Columns_TaxLot_Sub.Form.MoveRec 0, arr_varRec(R_TID, lngX)  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
1740            Exit For
1750          End If
1760        Next
1770      End If

1780      Select Case blnHasForEx
          Case True
            ' ** If this asset is in a foreign currency, don't let them turn it off.
1790        .chkIncludeCurrency = True
1800        .chkIncludeCurrency.Locked = True
1810      Case False
            ' ** If it isn't, they can check it or not, as they wish.
1820        .chkIncludeCurrency.Locked = False
1830      End Select

1840      chkIncludeCurrency_AfterUpdate  ' ** Procedure: below.

1850      lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
1860      If lngMonitorCnt > 1& Then
1870        EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
1880        lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
1890        If lngMonitorNum > 1& Then
1900          LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
1910        End If
1920      End If

1930      intOpeningChoice = .opgChoice
          'blnFromSaleBtn = Forms(strCallingForm).frmJournal_Sub4_Sold.Form.FromSaleBtn

1940      If gblnGoToReport = True Then
1950        blnGoneToReport = True
1960        .cmdCancel.SetFocus
1970        .cmdOK.Enabled = False
1980      Else
1990        DoCmd.Hourglass False
2000        Select Case blnFromSaleBtn
            Case True
2010          .cmdCancel.SetFocus
2020          .cmdOK.Enabled = False
2030        Case False
2040          .cmdOK.SetFocus
2050        End Select
2060      End If

2070    End With

EXITP:
2080    Set rst = Nothing
2090    Exit Sub

ERRH:
2100    DoCmd.Hourglass False
2110    Select Case ERR.Number
        Case Else
2120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2130    End Select
2140    Resume EXITP

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
        ' **   Shareface:        R {shareface on frmJournal_Columns_TaxLot_Sub}
        ' **   Location:         T {Loc_Name_display on frmJournal_Columns_TaxLot_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmJournal_Columns_TaxLot_Sub}

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyEscape
2270        intRetVal = 0
2280        cmdCancel_Click  ' ** Procedure: Below.
2290      End Select
2300    End If

        ' ** Alt keys.
2310    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2320      Select Case intRetVal
          Case vbKeyR
2330        With Me
2340          intRetVal = 0
2350          If .frmJournal_Columns_TaxLot_Sub.Form.shareface.Enabled = True Then
2360            .frmJournal_Columns_TaxLot_Sub.SetFocus
2370            .frmJournal_Columns_TaxLot_Sub.Form.shareface.SetFocus
2380          Else
2390            Beep
2400          End If
2410        End With
2420      Case vbKeyT
2430        With Me
2440          intRetVal = 0
2450          .frmJournal_Columns_TaxLot_Sub.SetFocus
2460          .frmJournal_Columns_TaxLot_Sub.Form.Loc_Name_display.SetFocus
2470        End With
2480      End Select
2490    End If

        ' ** Ctrl keys.
2500    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2510      Select Case intRetVal
          Case vbKeyS
2520        With Me
2530          intRetVal = 0
2540          .frmJournal_Columns_TaxLot_Sub.SetFocus
2550          .frmJournal_Columns_TaxLot_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
2560        End With
2570      End Select
2580    End If

        ' ** Ctrl-Shift keys.
2590    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2600      Select Case intRetVal
          Case vbKeyF
2610        With Me
2620          intRetVal = 0
2630          .FocusHolder.SetFocus
2640        End With
2650      End Select
2660    End If

EXITP:
2670    KeyCode = intRetVal
2680    Exit Sub

ERRH:
2690    intRetVal = 0
2700    THAT_PROC = THIS_PROC
2710    That_Erl = Erl
2720    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2730    Resume EXITP

End Sub

Private Sub Form_Timer()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2810    With Me
2820      .TimerInterval = 0&
2830      If gblnSetFocus = True Then
2840        gblnSetFocus = False
2850        strFormName = vbNullString
2860  On Error Resume Next
2870        Do Until strFormName = THIS_NAME
2880          DoCmd.SelectObject acForm, THIS_NAME, False
2890          strFormName = Screen.ActiveForm.Name
2900        Loop
2910  On Error GoTo ERRH
2920        Select Case blnFromSaleBtn
            Case True
2930          .cmdCancel.SetFocus
2940          .cmdOK.Enabled = False
2950        Case False
2960          .cmdOK.SetFocus
2970        End Select
2980      End If  ' ** gblnSetFocus
2990      If gblnGoToReport = True Or .GoToReport_arw_pr_img.Visible = True Then
3000        Select Case blnGoingToReport
            Case True
3010          gblnGoToReport = False
3020          blnGoingToReport = False
3030          DoCmd.Hourglass False
3040          .GoToReport_arw_pr_img.Visible = False
3050          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
3060          DoEvents
3070        Case False
3080          DoCmd.Hourglass True  ' ** Make sure it's still running.
3090          DoEvents
3100          blnGoingToReport = True
3110          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3120          .GoToReport_arw_pr_img.Visible = True
3130          DoCmd.Hourglass False
3140          Beep  '##MoveBeepT
3150          gblnGoToReport = True
3160          DoEvents
3170          Forms("frmJournal_Columns").GTRStuff 1, False  ' ** Form Function: frmJournal_Columns.
3180          Forms(THIS_NAME).TimerInterval = (GTR_WAIT * 2&)
3190          gblnGoToReport = True
3200          blnGoingToReport = True
3210        End Select
3220      End If
3230    End With

EXITP:
3240    Exit Sub

ERRH:
3250    DoCmd.Hourglass False
3260    Select Case ERR.Number
        Case Else
3270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3280    End Select
3290    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3310    If DataErr > 0 Then
3320      Select Case DataErr
          Case 13, 2105, 2108, 2110, 2116, 2169, 2501, 3020, 3314, 7753, 7787, 8519
            ' ** 13  : Type mismatch.
            ' ** 2108: You must save the field before you execute the GoToControl action, the GoToControl method, or the SetFocus method.
            ' ** 2110: Microsoft Access can't move the focus to the control |.
            ' ** 2116: The value in the field or record violates the validation rule for the record or field.
            ' ** 2169: You can't save this record at this time.
            ' ** 2501: The '|' action was Canceled.
            ' ** 3020: Update or CancelUpdate without AddNew or Edit.
            ' ** 3314: The field '|' cannot contain a Null value because the Required property for this field is set to True.
            ' ** 7753: The value you entered doesn't meet the validation rule defined for the field or control.
            ' ** 7787: The record your are modifying has been changed by another user.
            ' ** 8519: You are about to delete | record(s).
            ' ** Do nothing.
3330        Response = acDataErrContinue
3340      Case 2046  ' ** The command or action isn't available now (first or last record).
3350        Response = acDataErrContinue
3360      Case Else
3370        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3380        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3390      End Select
3400    End If

EXITP:
3410    Exit Sub

ERRH:
3420    Select Case ERR.Number
        Case Else
3430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3440    End Select
3450    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3510    Select Case blnWasGTR
        Case True
3520      blnWasGTR = False
3530    Case False
          ' ** Don't save settings imposed by GoToReport!
3540      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
3550    End Select

3560    gstrFormQuerySpec = vbNullString

3570    If blnFromSaleBtn = True Then
          'Forms(strCallingForm).frmJournal_Sub4_Sold.Form.FromSaleBtn = False
3580    End If

3590    If IsLoaded("rptListOfTaxLots", acReport) = True Then  ' ** Module Functions: modFileUtilities.
3600      DoCmd.Close acReport, "rptListOfTaxLots"
3610    End If

3620    DoCmd.Hourglass False

EXITP:
3630    Exit Sub

ERRH:
3640    Select Case ERR.Number
        Case Else
3650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3660    End Select
3670    Resume EXITP

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

        Dim frm As Access.Form, frmSub As Access.Form
        Dim lngTmp01 As Long
        Dim lngX As Long, lngY As Long
        Dim blnContinue As Boolean, blnFound As Boolean

4110    blnContinue = True

4120    lngRevcodeID = 0&: lngTaxcode = 0&: strDesc = vbNullString

4130    Set frm = Forms(strCallingForm).frmJournal_Columns_Sub.Form

4140    With Me

4150      Set frmSub = .frmJournal_Columns_TaxLot_Sub.Form

4160      If frmSub.shareface = 0 Then
4170        blnFound = False
4180        If .opgChoice = .opgChoice_optSpecific.OptionValue Then
4190          blnContinue = False
4200          Beep
4210          MsgBox "The chosen Tax Lot has already been sold," & vbCrLf & _
                "and is unavailable for this sale.", vbInformation + vbOKOnly, "Select Lot"
4220        Else
              ' *************************************************
              ' ** Array: arr_varRec()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     tmpedit04_id    R_TID
              ' **     2       1     assetno         R_ASTNO
              ' **     3       2     shareface       R_SHRS
              ' **     4       3     assetdate       R_ADAT
              ' **     5       4     IsZero          R_ZERO
              ' **     6       5     cost            R_COST
              ' **     7       6     priceperunit    R_PPU
              ' **
              ' *************************************************
4230          lngTmp01 = frmSub.tmpedit04_id
              ' ** This needs to be in the same order as the subform.
4240          For lngX = 0& To (lngRecs - 1&)
4250            If arr_varRec(R_TID, lngX) = lngTmp01 Then
4260              If lngX = (lngRecs - 1&) Then
                    ' ** We're at the end of the list.
4270                For lngY = 0& To (lngRecs - 1&)
4280                  If arr_varRec(R_TID, lngY) = lngTmp01 Then
                        ' ** They appear to all be sold!
4290                    blnContinue = False
4300                    Beep
4310                    MsgBox "All the Tax Lots have been previously sold," & vbCrLf & _
                          "and are unavailable for this sale.", vbInformation + vbOKOnly, "Nothing To Do"
4320                    .cmdCancel.SetFocus
4330                    DoEvents
4340                    .cmdOK.Enabled = False
4350                  Else
4360                    Select Case arr_varRec(R_ZERO, lngY)
                        Case True
                          ' ** Keep looking.
4370                    Case False
                          ' ** We'll stop here.
4380                      blnFound = True
4390                      frmSub.MoveRec 0, arr_varRec(R_TID, lngY)  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
4400                      DoEvents
4410                      Exit For
4420                    End Select
4430                  End If
4440                Next
                    ' ** If a suitable alternative was found, we can continue, below.
                    ' ** If none were found, that's all she wrote!
4450                Exit For
4460              Else
4470                Select Case arr_varRec(R_ZERO, (lngX + 1&))
                    Case True
                      ' ** Keep looking.
4480                Case False
                      ' ** We'll stop here.
4490                  blnFound = True
4500                  frmSub.MoveRec 0, arr_varRec(R_TID, lngX)  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
4510                  DoEvents
4520                  Exit For
4530                End Select
4540              End If
4550            End If
4560          Next
4570        End If
4580        If blnFound = False And blnContinue = True Then
              ' ** I think it should have hit one of the messages above.
              ' ** If no, oh well.
4590          blnContinue = False
4600          Beep
4610        End If
4620      End If

4630      If blnContinue = True Then

4640        lngRevcodeID = frm.revcode_ID
4650        lngTaxcode = frm.taxcode
4660        If IsNull(frm.description) = False Then
4670          strDesc = frm.description
4680        End If

4690        If .ShareFaceSale > frmSub.shareface Then
              ' ** A greater number than is in this lot is requested for sale.
4700          If .opgChoice <> .opgChoice_optFIFO.OptionValue And .opgChoice <> .opgChoice_optAverage.OptionValue Then
                ' ** Includes Specific, High, and Low.
4710            If .opgChoice = .opgChoice_optHigh.OptionValue Or .opgChoice = .opgChoice_optLow.OptionValue Then
                  ' ** High or Low.
4720              blnContinue = DoMultiLots  ' ** Procedure: Below.
4730              If blnContinue = True Then
4740                JColLocUpdate True, 0&  ' ** Procedure: Below.
4750                Select Case blnMultiLotSale
                    Case True
4760                  Forms(strCallingForm).ToTaxLot = 3&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...
4770                Case False
4780                  Forms(strCallingForm).ToTaxLot = 2&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...
4790                End Select
4800                frm.Requery
4810                DoEvents
4820                DoCmd.Close acForm, THIS_NAME
4830                gblnSetFocus = True
4840                Forms(strCallingForm).TimerInterval = 250&
4850                DoCmd.SelectObject acForm, "frmJournal_Columns", False
4860              End If
4870            Else
                  ' ** Specific.
4880              MsgBox "You will need to select another lot with enough" & vbCrLf & _
                    "shares/face for this request, or select First In First Out.", vbInformation + vbOKOnly, "Select Lot"
4890            End If
4900          Else
                ' ** FIFO or Average.
4910            blnContinue = DoMultiLots  ' ** Procedure: Below.
4920            If blnContinue = True Then
                  ' ** DoMultiLots() puts in IsAverage, if that's the case.
4930              JColLocUpdate True, 0&  ' ** Procedure: Below.
4940              Select Case blnMultiLotSale
                  Case True
4950                Forms(strCallingForm).ToTaxLot = 3&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...
4960              Case False
4970                Forms(strCallingForm).ToTaxLot = 2&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...
4980              End Select
4990              frm.Requery
5000              DoEvents
5010              DoCmd.Close acForm, THIS_NAME
5020              gblnSetFocus = True
5030              Forms(strCallingForm).TimerInterval = 250&
5040              DoCmd.SelectObject acForm, "frmJournal_Columns", False
5050            End If
5060          End If
5070        Else
              ' ** Asset covers sale.
5080          If frm.shareface <> 0 Then
5090            frm.Cost = Format(((.ShareFaceSale * (frmSub.Cost / frmSub.shareface)) * -1), "Currency")
5100            gstrSaleCost = CStr(dblCost)
5110            gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
5120            frm.PurchaseDate = frmSub.assetdate
5130            If IsNull(frm.assetdate) Then
5140              frm.assetdate = Now()
5150            End If
                ' ** VGC 12/28/2009: If they've checked IsAverage, put it in the Journal too; just 'cause I'd like to see it!
5160            If .opgChoice = .opgChoice_optAverage.OptionValue Then
5170              frm.IsAverage = True  'ISAVERAGE GETS PUT IN HERE!
5180            End If
5190          End If
5200          JColLocUpdate False, frmSub.Location_ID  ' ** Procedure: Below.
5210          Forms(strCallingForm).ToTaxLot = 2&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...
5220          frm.Requery
5230          DoEvents
5240          DoCmd.Close acForm, THIS_NAME
5250          gblnSetFocus = True
5260          Forms(strCallingForm).TimerInterval = 250&
5270          DoCmd.SelectObject acForm, "frmJournal_Columns", False
5280        End If

5290      End If  ' ** blnContinue.

5300    End With

EXITP:
5310    Set frmSub = Nothing
5320    Set frm = Nothing
5330    Exit Sub

ERRH:
5340    DoCmd.Hourglass False
5350    Select Case ERR.Number
        Case 2488  ' ** You can't use the ApplyFilter action on this window.
          ' ** Skip the ApplyFilter error
5360    Case Else
5370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5380    End Select
5390    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5420    intShiftDown = (Shift And acShiftMask) > 0
5430    intAltDown = (Shift And acAltMask) > 0
5440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5460      Select Case intRetVal
          Case vbKeyTab
5470        With Me
5480          intRetVal = 0
5490          .cmdCancel.SetFocus
5500        End With
5510      End Select
5520    End If

        ' ** Shift keys.
5530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5540      Select Case intRetVal
          Case vbKeyTab
5550        With Me
5560          intRetVal = 0
5570          If .cmdPrintReport.Enabled = True Then
5580            .cmdPrintReport.SetFocus
5590          Else
5600            .frmJournal_Columns_TaxLot_Sub.SetFocus
5610            .frmJournal_Columns_TaxLot_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
5620            .frmJournal_Columns_TaxLot_Sub.Form.FocusHolder.SetFocus
5630          End If
5640        End With
5650      End Select
5660    End If

EXITP:
5670    KeyCode = intRetVal
5680    Exit Sub

ERRH:
5690    intRetVal = 0
5700    THAT_PROC = THIS_PROC
5710    That_Erl = Erl
5720    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5730    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim frmSub As Access.Form
        Dim blnGoneToReport2 As Boolean, blnRecsTotUpdate As Boolean

5810    Forms(strCallingForm).ToTaxLot = -4&  ' ** 0 = Nothing; 1 = To Tax Lot; 2 = OK Single; 3 = OK Multi; -4 = Cancel Return; PLUS...

5820    Set frmSub = Forms("frmJournal_Columns").frmJournal_Columns_Sub.Form

5830    If blnGoneToReport = True Then
          ' ** This has become an unholy mess!
5840      DoCmd.Hourglass True
5850      DoEvents
5860      blnGoneToReport2 = frmSub.GoneToReport2_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
5870      blnRecsTotUpdate = frmSub.RecsTotUpdate_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
5880      JC_Rec_DelRec blnGoneToReport, blnGoneToReport2, blnRecsTotUpdate, frmSub  ' ** Module Procedure: modJrnlCol_Recs.
5890      frmSub.GoneToReport2_GetSet False, blnGoneToReport2  ' ** Form Function: frmJournal_Columns_Sub.
5900      frmSub.RecsTotUpdate_GetSet False, blnRecsTotUpdate  ' ** Form Function: frmJournal_Columns_Sub.
5910      DoEvents
5920    End If

5930    DoCmd.Close acForm, THIS_NAME
5940    gblnSetFocus = True
5950    DoCmd.SelectObject acForm, strCallingForm, False
5960    If blnGoneToReport = False Then
5970      Forms(strCallingForm).TimerInterval = 250&
5980    End If

EXITP:
5990    Set frmSub = Nothing
6000    Exit Sub

ERRH:
6010    Select Case ERR.Number
        Case Else
6020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6030    End Select
6040    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6120    intShiftDown = (Shift And acShiftMask) > 0
6130    intAltDown = (Shift And acAltMask) > 0
6140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6160      Select Case intRetVal
          Case vbKeyTab
6170        With Me
6180          intRetVal = 0
6190          .shortname.SetFocus
6200        End With
6210      End Select
6220    End If

        ' ** Shift keys.
6230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6240      Select Case intRetVal
          Case vbKeyTab
6250        With Me
6260          intRetVal = 0
6270          If .cmdOK.Enabled = True Then
6280            .cmdOK.SetFocus
6290          Else
6300            .frmJournal_Columns_TaxLot_Sub.SetFocus
6310          End If
6320        End With
6330      End Select
6340    End If

EXITP:
6350    KeyCode = intRetVal
6360    Exit Sub

ERRH:
6370    intRetVal = 0
6380    THAT_PROC = THIS_PROC
6390    That_Erl = Erl
6400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6410    Resume EXITP

End Sub

Private Sub shortname_GotFocus()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_GotFocus"

6510    With Me
6520      .shortname.SelLength = 0
6530      .shortname.SelStart = 99
6540    End With

EXITP:
6550    Exit Sub

ERRH:
6560    Select Case ERR.Number
        Case Else
6570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6580    End Select
6590    Resume EXITP

End Sub

Private Sub shortname_KeyDown(KeyCode As Integer, Shift As Integer)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6620    intShiftDown = (Shift And acShiftMask) > 0
6630    intAltDown = (Shift And acAltMask) > 0
6640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6670        With Me
6680          intRetVal = 0
6690          .opgChoice.SetFocus
6700        End With
6710      End Select
6720    End If

        ' ** Shift keys.
6730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6750        With Me
6760          intRetVal = 0
6770          .cmdCancel.SetFocus
6780        End With
6790      End Select
6800    End If

EXITP:
6810    KeyCode = intRetVal
6820    Exit Sub

ERRH:
6830    intRetVal = 0
6840    THAT_PROC = THIS_PROC
6850    That_Erl = Erl
6860    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6870    Resume EXITP

End Sub

Private Sub opgChoice_Enter()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_Enter"

6910    With Me
6920      intLastChoice = .opgChoice
6930    End With

EXITP:
6940    Exit Sub

ERRH:
6950    Select Case ERR.Number
        Case Else
6960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6970    End Select
6980    Resume EXITP

End Sub

Public Sub opgChoice_AfterUpdate()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, frmSub As Access.Form
        Dim dblMax As Double, dblMin As Double
        Dim datAssetDate As Date
        Dim strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim lngX As Long

7010    With Me

          ' ** Even though we're now using an array, let this
          ' ** sorting continue so they'll see what's happening.

7020      strAccountNo = .accountno
7030      lngAssetNo = .assetno

7040      .opgChoice_optFIFO_lbl.FontBold = False
7050      .opgChoice_optFIFO_box.Visible = False
7060      .opgChoice_optAverage_lbl.FontBold = False
7070      .opgChoice_optAverage_box.Visible = False
7080      .opgChoice_optSpecific_lbl.FontBold = False
7090      .opgChoice_optSpecific_box.Visible = False
7100      .opgChoice_optHigh_lbl.FontBold = False
7110      .opgChoice_optHigh_box.Visible = False
7120      .opgChoice_optLow_lbl.FontBold = False
7130      .opgChoice_optLow_box.Visible = False

7140      Set frmSub = .frmJournal_Columns_TaxLot_Sub.Form

7150      Set dbs = CurrentDb

7160      Select Case .opgChoice
          Case .opgChoice_optFIFO.OptionValue

7170        .opgChoice_optFIFO_lbl.FontBold = True
7180        .opgChoice_optFIFO_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
7190        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
7200        frmSub.SortNow "assetdate_display_lbl_DblClick"  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.

            ' ** ActiveAssets, grouped by assetno, accountno, assetdate, by specified [actno], [astno].
7210        Set qdf = dbs.QueryDefs("qryJournal_Columns_25_01")
7220        With qdf.Parameters
7230          ![actno] = strAccountNo
7240          ![astno] = lngAssetNo
7250        End With
7260        Set rst1 = qdf.OpenRecordset
7270        With rst1
7280          .MoveFirst
7290          dblCost = ((![Cost] / ![shareface]) * Me.ShareFaceSale) * -1
7300          dblPCash = ((![Cost] / ![shareface]) * Me.ShareFaceSale)
7310          datAssetDate = ![assetdate]
7320          .Close
7330        End With
7340        Set rst1 = Nothing
7350        Set qdf = Nothing

            ' ** Find the appropriate record.
7360        Set rst1 = frmSub.RecordsetClone
7370        With rst1
7380          .MoveFirst
7390          For lngX = 1& To lngLotCnt
7400            If ![assetdate] = datAssetDate Then
7410              frmSub.Bookmark = .Bookmark
7420              Exit For
7430            End If
7440            If lngX < lngLotCnt Then .MoveNext
7450          Next
7460          .Close
7470        End With
7480        Set rst1 = Nothing

7490        frmSub.FocusHolder.SetFocus
7500        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
7510        .opgChoice = .opgChoice_optFIFO.OptionValue

7520      Case .opgChoice_optAverage.OptionValue

7530        .opgChoice_optAverage_lbl.FontBold = True
7540        .opgChoice_optAverage_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
7550        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
7560        frmSub.SortNow "assetdate_display_lbl_DblClick"  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.

7570        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
7580        If frmSub.IsAverage = False Then

7590          Select Case blnFromSaleBtn
              Case True
7600            strMsg = "If you choose to proceed, the price will be fixed at the" & vbCrLf & _
                  "average price of this asset until all the lots for this asset are sold." & vbCrLf & vbCrLf & _
                  "This holds even if the Tax Lot Holdings window is canceled." & vbCrLf & vbCrLf & _
                  "Do you want to proceed?"
7610          Case False
7620            strMsg = "If you choose to proceed, the price will be fixed at the" & vbCrLf & _
                  "average price of this asset until all the lots for this asset are sold." & vbCrLf & vbCrLf & _
                  "Do you want to proceed?"
7630          End Select
7640          msgResponse = MsgBox(strMsg, vbInformation + vbOKCancel, "Compute Average Price")

7650          Select Case msgResponse
              Case vbOK

                ' ** Update ActiveAssets, for IsAverage = True, by specified [actno], [astno].
7660            Set qdf = dbs.QueryDefs("qryJournal_Columns_25_02")
7670            With qdf.Parameters
7680              ![actno] = strAccountNo
7690              ![astno] = lngAssetNo
7700            End With
7710            qdf.Execute
7720            Set qdf = Nothing

                ' ** This will set get the average prices for all lots for NOW; uses table called "tmpAveragePrice".

7730            If TableExists("tmpAveragePrice") = False Then  ' ** Module Function: modFileUtilities.
                  ' ** Data-Definition: Create table tmpAveragePrice.
7740              Set qdf = dbs.QueryDefs("zz_qry_System_46_01")
7750              qdf.Execute
7760              Set qdf = Nothing
                  ' ** Data-Definition: Create index [assetno], [accountno] PrimaryKey on table tmpAveragePrice.
7770              Set qdf = dbs.QueryDefs("zz_qry_System_46_02")
7780              qdf.Execute
7790              Set qdf = Nothing
7800            End If

                ' ** Empty tmpAveragePrice.
7810            Set qdf = dbs.QueryDefs("qryJournal_Columns_25_05")
7820            qdf.Execute
7830            Set qdf = Nothing

                ' ** Append qryJournal_Columns_25_04 (qryJournal_Columns_25_03 (ActiveAssets, grouped and summed,
                ' ** by accountno, assetno, by specified [actno], [astno]), with averagePrice) to tmpAveragePrice.
7840            Set qdf = dbs.QueryDefs("qryJournal_Columns_25_06")
7850            With qdf.Parameters
7860              ![actno] = strAccountNo
7870              ![astno] = lngAssetNo
7880            End With
7890            qdf.Execute
7900            Set qdf = Nothing

                ' ** This will apply those average prices to the ActiveAssets
                ' ** averagepriceperunit field.  This field will be used when the
                ' ** specified lots of an asset,acccount number pair have been averaged.

                ' ** NEW: 10/23/2014!
                ' ** Update qryJournal_Columns_25_07_05 (qryJournal_Columns_25_07_01 (ActiveAssets, linked to
                ' ** tmpAveragePrice, with uniqueid), with DLookups() to qryJournal_Columns_25_07_04
                ' ** (ActiveAssets, linked to qryJournal_Columns_25_07_03 (qryJournal_Columns_25_07_02
                ' ** (ActiveAssets, linked to tmpAveragePrice, grouped and summed), with costpershare),
                ' ** with averagepriceperunit_new, cost_new)), Update ActiveAssets, from tmpAveragePrice.
7910            Set qdf = dbs.QueryDefs("qryJournal_Columns_25_07")
7920            qdf.Execute
7930            Set qdf = Nothing

                ' ** ActiveAssets, grouped by assetno, accountno, assetdate, by specified [actno], [astno].
7940            Set qdf = dbs.QueryDefs("qryJournal_Columns_25_08")
7950            With qdf.Parameters
7960              ![actno] = strAccountNo
7970              ![astno] = lngAssetNo
7980            End With
7990            Set rst1 = qdf.OpenRecordset
8000            rst1.MoveLast
8010            rst1.MoveFirst

8020            With frmSub

8030              .FocusHolder.SetFocus
8040              DoEvents

8050              .averagepriceperunit.Visible = True
8060              .averagepriceperunit_lbl.Visible = True
8070              .averagepriceperunit_lbl_line.Visible = True
8080              .averagepriceperunit_avg.Visible = True
8090              .priceperunit.Visible = False
8100              .priceperunit_lbl.Visible = False
8110              .priceperunit_lbl_line.Visible = False
8120              .priceperunit_avg.Visible = False
8130              DoEvents

8140              .MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.

8150            End With

                ' ** Find the appropriate record.
                'APPROPRIATE RECORD? THEY ALL HAVE TO CHANGE!
                'For lngX = 1& To lngLotCnt
                '  rst1.FindFirst "[accountno] = '" & .accountno & "' And [assetno] = " & CStr(.assetno) & " And " & _
                '    "[assetdate] = #" & Format(frmSub.assetdate, "mm/dd/yyyy hh:nn:ss") & "#"
                '  If rst1.NoMatch = False Then
                '    frmSub.priceperunit = rst1![priceperunit]
                '  End If
                '  DoEvents
                '  If lngX < lngLotCnt Then
                '    frmSub.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
                '    frmSub.FocusHolder.SetFocus
                '    DoEvents
                '  End If
                'Next
                'rst1.Close
                'Set rst1 = Nothing
                'Set qdf = Nothing

8160            Set rst2 = frmSub.RecordsetClone
8170            rst2.MoveFirst
8180            For lngX = 1& To lngLotCnt
8190              rst1.FindFirst "[accountno] = '" & .accountno & "' And [assetno] = " & CStr(.assetno) & " And " & _
                    "[assetdate] = #" & Format(rst2![assetdate], "mm/dd/yyyy hh:nn:ss") & "#"
8200              If rst1.NoMatch = False Then
8210                rst2.Edit
8220                rst2![averagepriceperunit] = rst1![averagepriceperunit]
8230                rst2![Cost] = rst1![Cost]
8240                rst2.Update
8250              End If
8260              If lngX < lngLotCnt Then rst2.MoveNext
8270            Next
8280            rst2.Close
8290            Set rst2 = Nothing
8300            rst1.Close
8310            Set rst1 = Nothing
8320            Set qdf = Nothing

8330            frmSub.Refresh
8340            frmSub.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
8350            frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.

8360            .opgChoice = .opgChoice_optAverage.OptionValue
8370            .opgChoice_optFIFO.Enabled = False
8380            .opgChoice_optAverage.Enabled = True
8390            .opgChoice_optSpecific.Enabled = False
8400            .opgChoice_optHigh.Enabled = False
8410            .opgChoice_optLow.Enabled = False

8420            frmSub.Refresh

8430            .opgChoice = .opgChoice_optAverage.OptionValue

8440          Case Else
8450            .opgChoice = intLastChoice
8460            .opgChoice_optAverage_lbl.FontBold = False
8470            Select Case .opgChoice
                Case .opgChoice_optFIFO.OptionValue
8480              .opgChoice_optFIFO_lbl.FontBold = True
8490            Case .opgChoice_optSpecific.OptionValue
8500              .opgChoice_optSpecific_lbl.FontBold = True
8510            Case .opgChoice_optLow.OptionValue
8520              .opgChoice_optLow_lbl.FontBold = True
8530            Case .opgChoice_optHigh.OptionValue
8540              .opgChoice_optHigh_lbl.FontBold = True
8550            End Select
8560          End Select
8570        End If

8580      Case .opgChoice_optSpecific.OptionValue

8590        .opgChoice_optSpecific_lbl.FontBold = True
8600        .opgChoice_optSpecific_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
8610        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
8620        frmSub.SortNow "assetdate_display_lbl_DblClick"  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.

8630        dblCost = ((frmSub.priceperunit * .ShareFaceSale) * -1)
8640        dblPCash = (frmSub.priceperunit * .ShareFaceSale)

8650        frmSub.FocusHolder.SetFocus
8660        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
8670        .opgChoice = .opgChoice_optSpecific.OptionValue

8680      Case .opgChoice_optHigh.OptionValue

8690        .opgChoice_optHigh_lbl.FontBold = True
8700        .opgChoice_optHigh_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
8710        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
8720        frmSub.SortNow "priceperunit_lbl_DblClick"  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.

8730        Set rst1 = frmSub.RecordsetClone
8740        With rst1
8750          dblMax = 0
8760          .MoveFirst
8770          For lngX = 1& To lngLotCnt
8780            If dblMax < ![priceperunit] Then
8790              dblMax = ![priceperunit]
8800            End If
8810            If lngX < lngLotCnt Then .MoveNext
8820          Next
8830          .Close
8840        End With
8850        Set rst1 = Nothing

8860        dblCost = ((dblMax * .ShareFaceSale) * -1)
8870        dblPCash = (dblMax * .ShareFaceSale)

            ' ** Find the appropriate record.
8880        Set rst1 = frmSub.RecordsetClone
8890        With rst1
8900          .MoveFirst
8910          For lngX = 1& To lngLotCnt
8920            If Format(![priceperunit], "####.#####") = Format(dblMax, "####.#####") Then
8930              frmSub.Bookmark = .Bookmark
8940              Exit For
8950            End If
8960            If lngX < lngLotCnt Then .MoveNext
8970          Next
8980          .Close
8990        End With
9000        Set rst1 = Nothing

9010        frmSub.FocusHolder.SetFocus
9020        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
9030        .opgChoice = .opgChoice_optHigh.OptionValue

9040      Case .opgChoice_optLow.OptionValue

9050        .opgChoice_optLow_lbl.FontBold = True
9060        .opgChoice_optLow_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
9070        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
9080        frmSub.SortNow "priceperunit_lbl_DblClick"  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
9090        DoEvents
            ' ** Then do it again to get descending.
9100        frmSub.SortNow "priceperunit_lbl_DblClick"  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.

9110        Set rst1 = frmSub.RecordsetClone
9120        With rst1
9130          dblMin = 999999999
9140          .MoveFirst
9150          For lngX = 1& To lngLotCnt
9160            If dblMin > ![priceperunit] Then
9170              dblMin = ![priceperunit]
9180            End If
9190            If lngX < lngLotCnt Then .MoveNext
9200          Next
9210          .Close
9220        End With
9230        Set rst1 = Nothing

9240        dblCost = ((dblMin * .ShareFaceSale) * -1)
9250        dblPCash = (dblMin * .ShareFaceSale)

            ' ** Find the appropriate record.
9260        Set rst1 = frmSub.RecordsetClone
9270        With rst1
9280          .MoveFirst
9290          For lngX = 1& To lngLotCnt
9300            If Format(![priceperunit], "####.#####") = Format(dblMin, "####.#####") Then
9310              frmSub.Bookmark = .Bookmark
9320              Exit For
9330            End If
9340            If lngX < lngLotCnt Then .MoveNext
9350          Next
9360          .Close
9370        End With
9380        Set rst1 = Nothing

9390        frmSub.FocusHolder.SetFocus
9400        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
9410        .opgChoice = .opgChoice_optLow.OptionValue

9420      End Select

9430      DoEvents
9440      RecArray_Load  ' ** Procedure: Below.

9450      intLastChoice = .opgChoice

9460    End With

9470    dbs.Close

EXITP:
9480    Set frmSub = Nothing
9490    Set rst1 = Nothing
9500    Set rst2 = Nothing
9510    Set qdf = Nothing
9520    Set dbs = Nothing
9530    Exit Sub

ERRH:
9540    Select Case ERR.Number
        Case Else
9550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9560    End Select
9570    Resume EXITP

End Sub

Private Sub opgChoice_optFIFO_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optFIFO_KeyDown"

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
9690          .frmJournal_Columns_TaxLot_Sub.SetFocus
9700          .frmJournal_Columns_TaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
9710          .frmJournal_Columns_TaxLot_Sub.Form.FocusHolder.SetFocus
9720        End With
9730      End Select
9740    End If

        ' ** Shift keys.
9750    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9770        With Me
9780          intRetVal = 0
9790          .shortname.SetFocus
9800        End With
9810      End Select
9820    End If

EXITP:
9830    KeyCode = intRetVal
9840    Exit Sub

ERRH:
9850    intRetVal = 0
9860    THAT_PROC = THIS_PROC
9870    That_Erl = Erl
9880    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9890    Resume EXITP

End Sub

Private Sub opgChoice_optAverage_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optAverage_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9920    intShiftDown = (Shift And acShiftMask) > 0
9930    intAltDown = (Shift And acAltMask) > 0
9940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9970        With Me
9980          intRetVal = 0
9990          .frmJournal_Columns_TaxLot_Sub.SetFocus
10000         .frmJournal_Columns_TaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
10010         .frmJournal_Columns_TaxLot_Sub.Form.FocusHolder.SetFocus
10020       End With
10030     End Select
10040   End If

        ' ** Shift keys.
10050   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10070       With Me
10080         intRetVal = 0
10090         .shortname.SetFocus
10100       End With
10110     End Select
10120   End If

EXITP:
10130   KeyCode = intRetVal
10140   Exit Sub

ERRH:
10150   intRetVal = 0
10160   THAT_PROC = THIS_PROC
10170   That_Erl = Erl
10180   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10190   Resume EXITP

End Sub

Private Sub opgChoice_optSpecific_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optSpecific_KeyDown"

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
10290         .frmJournal_Columns_TaxLot_Sub.SetFocus
10300         .frmJournal_Columns_TaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
10310         .frmJournal_Columns_TaxLot_Sub.Form.FocusHolder.SetFocus
10320       End With
10330     End Select
10340   End If

        ' ** Shift keys.
10350   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10370       With Me
10380         intRetVal = 0
10390         .shortname.SetFocus
10400       End With
10410     End Select
10420   End If

EXITP:
10430   KeyCode = intRetVal
10440   Exit Sub

ERRH:
10450   intRetVal = 0
10460   THAT_PROC = THIS_PROC
10470   That_Erl = Erl
10480   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10490   Resume EXITP

End Sub

Private Sub opgChoice_optHigh_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optHigh_KeyDown"

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
10590         .frmJournal_Columns_TaxLot_Sub.SetFocus
10600         .frmJournal_Columns_TaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
10610         .frmJournal_Columns_TaxLot_Sub.Form.FocusHolder.SetFocus
10620       End With
10630     End Select
10640   End If

        ' ** Shift keys.
10650   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10670       With Me
10680         intRetVal = 0
10690         .shortname.SetFocus
10700       End With
10710     End Select
10720   End If

EXITP:
10730   KeyCode = intRetVal
10740   Exit Sub

ERRH:
10750   intRetVal = 0
10760   THAT_PROC = THIS_PROC
10770   That_Erl = Erl
10780   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10790   Resume EXITP

End Sub

Private Sub opgChoice_optLow_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optLow_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10820   intShiftDown = (Shift And acShiftMask) > 0
10830   intAltDown = (Shift And acAltMask) > 0
10840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         intRetVal = 0
10890         .frmJournal_Columns_TaxLot_Sub.SetFocus
10900         .frmJournal_Columns_TaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
10910         .frmJournal_Columns_TaxLot_Sub.Form.FocusHolder.SetFocus
10920       End With
10930     End Select
10940   End If

        ' ** Shift keys.
10950   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10970       With Me
10980         intRetVal = 0
10990         .shortname.SetFocus
11000       End With
11010     End Select
11020   End If

EXITP:
11030   KeyCode = intRetVal
11040   Exit Sub

ERRH:
11050   intRetVal = 0
11060   THAT_PROC = THIS_PROC
11070   That_Erl = Erl
11080   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11090   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

11110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11120   Exit Sub

ERRH:
11130   Select Case ERR.Number
        Case Else
11140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11150   End Select
11160   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

11210   With Me

11220     Select Case .chkIncludeCurrency
          Case True
11230       .chkIncludeCurrency_lbl.FontBold = True
11240       .frmJournal_Columns_TaxLot_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
11250       IncludeCurrency True  ' ** Procedure: Below.
11260     Case False
11270       .chkIncludeCurrency_lbl.FontBold = False
11280       .frmJournal_Columns_TaxLot_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
11290       IncludeCurrency False  ' ** Procedure: Below.
11300     End Select

11310     If lngMonitorCnt > 1& And blnIsOpen = False Then
11320       EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
11330       lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
11340       If lngMonitorNum > 1& Then
11350         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
11360       End If
11370     End If

11380   End With

EXITP:
11390   Exit Sub

ERRH:
11400   Select Case ERR.Number
        Case Else
11410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11420   End Select
11430   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

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
11590         If .cmdPrintReport.Enabled = True Then
11600           .cmdPrintReport.SetFocus
11610         ElseIf .cmdOK.Enabled = True Then
11620           .cmdOK.SetFocus
11630         Else
11640           .cmdCancel.SetFocus
11650         End If
11660       End With
11670     End Select
11680   End If

        ' ** Shift keys.
11690   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11700     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11710       With Me
11720         intRetVal = 0
11730         .frmJournal_Columns_TaxLot_Sub.SetFocus
11740       End With
11750     End Select
11760   End If

EXITP:
11770   KeyCode = intRetVal
11780   Exit Sub

ERRH:
11790   intRetVal = 0
11800   THAT_PROC = THIS_PROC
11810   That_Erl = Erl
11820   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11830   Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

11910   With Me
11920     blnPrintReport_Focus = True
11930     .cmdPrintReport_raised_semifocus_dots_img.Visible = True
11940     .cmdPrintReport_raised_img.Visible = False
11950     .cmdPrintReport_raised_focus_img.Visible = False
11960     .cmdPrintReport_raised_focus_dots_img.Visible = False
11970     .cmdPrintReport_sunken_focus_dots_img.Visible = False
11980     .cmdPrintReport_raised_img_dis.Visible = False
11990   End With

EXITP:
12000   Exit Sub

ERRH:
12010   Select Case ERR.Number
        Case Else
12020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12030   End Select
12040   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

12110   With Me
12120     blnPrintReport_MouseDown = True
12130     .cmdPrintReport_sunken_focus_dots_img.Visible = True
12140     .cmdPrintReport_raised_img.Visible = False
12150     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
12160     .cmdPrintReport_raised_focus_img.Visible = False
12170     .cmdPrintReport_raised_focus_dots_img.Visible = False
12180     .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_Click()

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

12310   With Me
          ' ** Uses qryLotInformation_07.
12320     strDocName = "rptListOfTaxLots"
12330     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
12340       .Modal = False
12350       DoCmd.OpenReport strDocName, acViewPreview, , , , THIS_NAME
12360       DoCmd.Maximize
12370       DoCmd.RunCommand acCmdFitToWindow
12380     Else
            '##GTR_Ref: rptListOfTaxLots
12390       DoCmd.OpenReport strDocName, acViewNormal, , , , THIS_NAME
12400     End If
12410   End With

EXITP:
12420   Exit Sub

ERRH:
12430   Select Case ERR.Number
        Case Else
12440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12450   End Select
12460   Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
          Case vbKeyTab
12570       With Me
12580         intRetVal = 0
12590         If .cmdOK.Enabled = True Then
12600           .cmdOK.SetFocus
12610         Else
12620           .cmdCancel.SetFocus
12630         End If
12640       End With
12650     End Select
12660   End If

        ' ** Shift keys.
12670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12680     Select Case intRetVal
          Case vbKeyTab
12690       With Me
12700         intRetVal = 0
12710         .frmJournal_Columns_TaxLot_Sub.SetFocus
12720         .frmJournal_Columns_TaxLot_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
12730         .frmJournal_Columns_TaxLot_Sub.Form.FocusHolder.SetFocus
12740       End With
12750     End Select
12760   End If

EXITP:
12770   KeyCode = intRetVal
12780   Exit Sub

ERRH:
12790   intRetVal = 0
12800   THAT_PROC = THIS_PROC
12810   That_Erl = Erl
12820   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12830   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

12910   With Me
12920     If blnPrintReport_MouseDown = False Then
12930       Select Case blnPrintReport_Focus
            Case True
12940         .cmdPrintReport_raised_focus_dots_img.Visible = True
12950         .cmdPrintReport_raised_focus_img.Visible = False
12960       Case False
12970         .cmdPrintReport_raised_focus_img.Visible = True
12980         .cmdPrintReport_raised_focus_dots_img.Visible = False
12990       End Select
13000       .cmdPrintReport_raised_img.Visible = False
13010       .cmdPrintReport_raised_semifocus_dots_img.Visible = False
13020       .cmdPrintReport_sunken_focus_dots_img.Visible = False
13030       .cmdPrintReport_raised_img_dis.Visible = False
13040     End If
13050   End With

EXITP:
13060   Exit Sub

ERRH:
13070   Select Case ERR.Number
        Case Else
13080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13090   End Select
13100   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

13210   With Me
13220     .cmdPrintReport_raised_focus_dots_img.Visible = True
13230     .cmdPrintReport_raised_img.Visible = False
13240     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
13250     .cmdPrintReport_raised_focus_img.Visible = False
13260     .cmdPrintReport_sunken_focus_dots_img.Visible = False
13270     .cmdPrintReport_raised_img_dis.Visible = False
13280     blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

13410   With Me
13420     .cmdPrintReport_raised_img.Visible = True
13430     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
13440     .cmdPrintReport_raised_focus_img.Visible = False
13450     .cmdPrintReport_raised_focus_dots_img.Visible = False
13460     .cmdPrintReport_sunken_focus_dots_img.Visible = False
13470     .cmdPrintReport_raised_img_dis.Visible = False
13480     blnPrintReport_Focus = False
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

Private Function DoMultiLots() As Boolean

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "DoMultiLots"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim frm As Access.Form, frmSub As Access.Form
        Dim varAccountNo As Variant, varSaletype As Variant, varTranDate As Variant
        Dim varAssetDate As Variant, varAssetNo As Variant
        Dim intChoice As Integer
        Dim lngNews As Long, arr_lngNew() As Long
        Dim dblPerShare_First As Double, dblPerShare_Loop As Double
        Dim varTmp00 As Variant, lngTmp01 As Long, arr_varTmp01() As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

13610   blnRetVal = True

13620   Set frm = Forms(strCallingForm).frmJournal_Columns_Sub.Form

13630   With Me

13640     Set frmSub = .frmJournal_Columns_TaxLot_Sub.Form

13650     dblPerShare_First = 0#: dblPerShare_Loop = 0#

13660     lngNews = 0&
13670     ReDim arr_lngNew(0)

13680     lngLots = 0&
13690     ReDim arr_varLots(L_ELEMS, 0)
13700     intChoice = .opgChoice  ' ** Save this!

          ' ** Instead of resorting the form, just sort the array.
13710     lngTmp01 = 0&
13720     ReDim arr_varTmp01(R_ELEMS, 0)
13730     For lngX = 0& To (lngRecs - 1&)
13740       If arr_varRec(R_ZERO, lngX) = False Then
13750         lngTmp01 = lngTmp01 + 1&
13760         lngE = lngTmp01 - 1&
13770         ReDim Preserve arr_varTmp01(R_ELEMS, lngE)
13780         For lngZ = 0& To R_ELEMS
13790           arr_varTmp01(lngZ, lngE) = arr_varRec(lngZ, lngX)
13800         Next
13810       End If
13820     Next

          ' *************************************************
          ' ** Array: arr_varRec()
          ' **
          ' **   Field  Element  Name            Constant
          ' **   =====  =======  ==============  ==========
          ' **     1       0     tmpedit04_id    R_TID
          ' **     2       1     assetno         R_ASTNO
          ' **     3       2     shareface       R_SHRS
          ' **     4       3     assetdate       R_ADAT
          ' **     5       4     IsZero          R_ZERO
          ' **     6       5     cost            R_COST
          ' **     7       6     priceperunit    R_PPU
          ' **
          ' *************************************************

13830     Select Case .opgChoice
          Case .opgChoice_optFIFO.OptionValue
            ' ** Binary Sort arr_varTmp01() array by R_ADAT, oldest to newest.
13840       For lngX = UBound(arr_varTmp01, 2) To 1 Step -1
13850         For lngY = 0 To (lngX - 1)
13860           If arr_varTmp01(R_ADAT, lngY) > arr_varTmp01(R_ADAT, (lngY + 1)) Then
13870             For lngZ = 0& To R_ELEMS
13880               varTmp00 = arr_varTmp01(lngZ, lngY)
13890               arr_varTmp01(lngZ, lngY) = arr_varTmp01(lngZ, (lngY + 1&))
13900               arr_varTmp01(lngZ, (lngY + 1&)) = varTmp00
13910             Next
13920           End If
13930         Next
13940       Next
13950     Case .opgChoice_optAverage.OptionValue
            ' ** Changes already made in opgChoice_AfterUpdate().
            ' ** Binary Sort arr_varTmp01() array by R_ADAT, oldest to newest.
13960       For lngX = UBound(arr_varTmp01, 2) To 1 Step -1
13970         For lngY = 0 To (lngX - 1)
13980           If arr_varTmp01(R_ADAT, lngY) > arr_varTmp01(R_ADAT, (lngY + 1)) Then
13990             For lngZ = 0& To R_ELEMS
14000               varTmp00 = arr_varTmp01(lngZ, lngY)
14010               arr_varTmp01(lngZ, lngY) = arr_varTmp01(lngZ, (lngY + 1&))
14020               arr_varTmp01(lngZ, (lngY + 1&)) = varTmp00
14030             Next
14040           End If
14050         Next
14060       Next
14070     Case .opgChoice_optSpecific.OptionValue
            ' ** Handled above, so shouldn't ever get here.
14080       blnRetVal = False
14090     Case .opgChoice_optHigh.OptionValue
            'CONFIRM THIS SORTS CORRECTLY!
            ' ** Binary Sort arr_varTmp01() array by R_PPU, highest to lowest.
14100       For lngX = UBound(arr_varTmp01, 2) To 1 Step -1
14110         For lngY = 0 To (lngX - 1)
14120           If arr_varTmp01(R_PPU, lngY) < arr_varTmp01(R_PPU, (lngY + 1)) Then
14130             For lngZ = 0& To R_ELEMS
14140               varTmp00 = arr_varTmp01(lngZ, lngY)
14150               arr_varTmp01(lngZ, lngY) = arr_varTmp01(lngZ, (lngY + 1&))
14160               arr_varTmp01(lngZ, (lngY + 1&)) = varTmp00
14170             Next
14180           End If
14190         Next
14200       Next
14210     Case .opgChoice_optLow.OptionValue
            ' ** Binary Sort arr_varTmp01() array by R_PPU, lowest to highest.
14220       For lngX = UBound(arr_varTmp01, 2) To 1 Step -1
14230         For lngY = 0 To (lngX - 1)
14240           If arr_varTmp01(R_PPU, lngY) > arr_varTmp01(R_PPU, (lngY + 1)) Then
14250             For lngZ = 0& To R_ELEMS
14260               varTmp00 = arr_varTmp01(lngZ, lngY)
14270               arr_varTmp01(lngZ, lngY) = arr_varTmp01(lngZ, (lngY + 1&))
14280               arr_varTmp01(lngZ, (lngY + 1&)) = varTmp00
14290             Next
14300           End If
14310         Next
14320       Next
14330     End Select

14340     lngRecs = 0&
14350     ReDim arr_varRec(R_ELEMS, 0&)

14360     For lngX = 0& To (lngTmp01 - 1&)
14370       lngRecs = lngRecs + 1
14380       lngE = lngRecs - 1&
14390       ReDim Preserve arr_varRec(R_ELEMS, lngE)
14400       For lngZ = 0& To R_ELEMS
14410         arr_varRec(lngZ, lngE) = arr_varTmp01(lngZ, lngX)
14420       Next
14430     Next

          ' ** We'll use lngZ for navigating the arr_varRec() array.
14440     lngZ = 0&

14450     frm.Cost = Format(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), "Currency")
          ' ** First item takes that item's cost, but recalculates it. Do we get a remainder problem here?

14460     gstrSaleCost = Format(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), "Currency")
14470     gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
14480     frm.PurchaseDate = arr_varRec(R_ADAT, lngZ)
14490     frm.shareface = arr_varRec(R_SHRS, lngZ)  ' ** 1st Tax Lot shareface put into existing, 1st Journal entry.
14500     varTmp00 = CDbl(Val(gstrSalePCash))
14510     varPCash = CCur(varTmp00)
14520     frm.PCash = (varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
14530     varTmp00 = CDbl(Val(gstrSaleICash))
14540     varICash = CCur(varTmp00)
14550     frm.ICash = (varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
14560     varRemaining = (.ShareFaceSale - arr_varRec(R_SHRS, lngZ))  ' ** Local .ShareFaceSale is total to sell.

          ' ** VGC 12/28/2009: If they've checked IsAverage, put it in the Journal too; just 'cause I'd like to see it!
14570     If .opgChoice = .opgChoice_optAverage.OptionValue Then
14580       frm.IsAverage = True
14590     End If

14600     varAccountNo = frm.accountno
14610     varSaletype = frm.journaltype
14620     varTranDate = frm.transdate

14630     If IsNull(frm.assetdate) Then
14640       frm.assetdate = Now()
14650     ElseIf Not IsDate(frm.assetdate) Then
14660       frm.assetdate = Now()
14670     End If

14680     varAssetDate = frm.assetdate
14690     varAssetNo = frm.assetno

          ' ** OK, the idea is to first let it find all the pieces needed to
          ' ** cover the sale. Then, when that's done, pull all the fractions
          ' ** out, add them up, and distribute that among the chosen pieces.
          ' ** A penny here, a penny there, etc.

14700     lngLots = lngLots + 1&  ' ** First record created in sale.
14710     lngE = lngLots - 1&
14720     ReDim Preserve arr_varLot(L_ELEMS, lngE)
          ' *******************************************
          ' ** Array: arr_varLot()
          ' **
          ' **   Element  Name             Constant
          ' **   =======  ===============  ==========
          ' **      0     Lot Number       L_LOTNO
          ' **      1     Journal ID       L_JID
          ' **      2     accountno        L_ACTNO
          ' **      3     assetno          L_ASTNO
          ' **      4     Shareface        L_SHARE
          ' **      5     100ths           L_SCENT
          ' **      6     Shareface New    L_SNEW
          ' **      7     Distribution     L_SADD
          ' **      8     ICash            L_ICASH
          ' **      9     100ths           L_ICENT
          ' **     10     ICash New        L_INEW
          ' **     11     Distribution     L_IADD
          ' **     12     PCash            L_PCASH
          ' **     13     100ths           L_PCENT
          ' **     14     PCash New        L_PNEW
          ' **     15     Distribution     L_PADD
          ' **     16     Cost             L_COST
          ' **     17     100ths           L_CCENT
          ' **     18     Cost New         L_CNEW
          ' **     19     Distribution     L_CADD
          ' **     20     Share Remain     L_REM
          ' **
          ' *******************************************
          ' ** We're still on the first Tax Lot.
14730     arr_varLot(L_JID, lngE) = frm.JrnlCol_ID
14740     arr_varLot(L_LOTNO, lngE) = lngLots
14750     arr_varLot(L_ACTNO, lngE) = CStr(varAccountNo)
14760     arr_varLot(L_ASTNO, lngE) = CLng(varAssetNo)
14770     arr_varLot(L_SHARE, lngE) = CDbl(arr_varRec(R_SHRS, lngZ))
14780     arr_varLot(L_SCENT, lngE) = CCenti(arr_varRec(R_SHRS, lngZ), 4)
14790     arr_varLot(L_SNEW, lngE) = CDbl(0)
14800     arr_varLot(L_SADD, lngE) = CDbl(0)
14810     arr_varLot(L_ICASH, lngE) = CDbl(varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
14820     arr_varLot(L_ICENT, lngE) = CCenti((varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale)), 2)  ' ** Function: Below.
14830     arr_varLot(L_INEW, lngE) = CDbl(0)
14840     arr_varLot(L_IADD, lngE) = CDbl(0)
14850     arr_varLot(L_PCASH, lngE) = CDbl(varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
14860     arr_varLot(L_PCENT, lngE) = CCenti((varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale)), 2)  ' ** Function: Below.
14870     arr_varLot(L_PNEW, lngE) = CDbl(0)
14880     arr_varLot(L_PADD, lngE) = CDbl(0)
14890     arr_varLot(L_COST, lngE) = CDbl((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1)
14900     arr_varLot(L_CCENT, lngE) = CCenti(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), 2)
14910     arr_varLot(L_CNEW, lngE) = CDbl(0)
14920     arr_varLot(L_CADD, lngE) = CDbl(0)
14930     arr_varLot(L_REM, lngE) = CDbl(varRemaining)

14940     dblPerShare_First = frm.pershare

          ' ** Add the 1st record to the array.
14950     frm.NewRecAdd frm.JrnlCol_ID  ' ** Form Procedure: frmJournal_Columns_Sub.

14960   End With

14970   Set dbs = CurrentDb
14980   Set rst = dbs.OpenRecordset("tblJournal_Column", dbOpenDynaset, dbConsistent)

        ' ** This is the array loop.
14990   Do While CDbl(varRemaining) >= 0.0001  ' ** varRemaining is ShareFace; ignore floating point garbage.

15000     With Me

            ' ** The array is already in the proper order.
15010       lngZ = lngZ + 1&

15020       If blnRetVal = True Then

15030         With rst
15040           .AddNew
15050           ![Journal_ID] = CLng(0)
15060           ![posted] = False
15070           ![transdate] = varTranDate
15080           ![journaltype] = varSaletype
15090           ![journalSubtype] = Null
15100           ![journaltype_sortorder] = frm.journaltype_sortorder
15110           ![accountno] = varAccountNo
15120           ![shortname] = frm.shortname
15130           ![assetno] = varAssetNo
15140           ![assetno_description] = frm.assetno_description
15150           ![assetdate] = varAssetDate
15160           ![assetdate_display] = CDate(Format(varAssetDate, "mm/dd/yyyy"))
15170           ![PurchaseDate] = arr_varRec(R_ADAT, lngZ)
15180           ![revcode_ID] = lngRevcodeID
15190           ![taxcode] = lngTaxcode
15200           If strDesc <> vbNullString Then
15210             ![description] = strDesc
15220           End If
15230           ![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
15240         End With

15250         If varRemaining > arr_varRec(R_SHRS, lngZ) Then

15260           rst![PCash] = (varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15270           rst![ICash] = (varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15280           rst![Cost] = Format(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), "Currency")
                ' ** Each one takes that item's cost, but recalculates it.
                ' ** If it's selling all of them, wouldn't it be simpler to just take that item's cost, period?
                ' ** Like this?
                ' ** rst![cost] = Format((.cost * -1), "Currency")

15290           gstrSaleCost = Format(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), "Currency")
15300           gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
15310           rst![shareface] = arr_varRec(R_SHRS, lngZ)
15320           varRemaining = (varRemaining - arr_varRec(R_SHRS, lngZ))

15330           lngLots = lngLots + 1&
15340           lngE = lngLots - 1&
15350           ReDim Preserve arr_varLot(L_ELEMS, lngE)
15360           arr_varLot(L_JID, lngE) = CLng(0)
15370           arr_varLot(L_LOTNO, lngE) = lngLots
15380           arr_varLot(L_ACTNO, lngE) = CStr(varAccountNo)
15390           arr_varLot(L_ASTNO, lngE) = CLng(varAssetNo)
15400           arr_varLot(L_SHARE, lngE) = CDbl(arr_varRec(R_SHRS, lngZ))
15410           arr_varLot(L_SCENT, lngE) = CCenti(arr_varRec(R_SHRS, lngZ), 4)
15420           arr_varLot(L_SNEW, lngE) = CDbl(0)
15430           arr_varLot(L_SADD, lngE) = CDbl(0)
15440           arr_varLot(L_ICASH, lngE) = CDbl(varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15450           arr_varLot(L_ICENT, lngE) = CCenti((varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale)), 2)  ' ** Function: Below.
15460           arr_varLot(L_INEW, lngE) = CDbl(0)
15470           arr_varLot(L_IADD, lngE) = CDbl(0)
15480           arr_varLot(L_PCASH, lngE) = CDbl(varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15490           arr_varLot(L_PCENT, lngE) = CCenti((varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale)), 2)  ' ** Function: Below.
15500           arr_varLot(L_PNEW, lngE) = CDbl(0)
15510           arr_varLot(L_PADD, lngE) = CDbl(0)
15520           arr_varLot(L_COST, lngE) = CDbl((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1)
15530           arr_varLot(L_CCENT, lngE) = CCenti(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), 2)
15540           arr_varLot(L_CNEW, lngE) = CDbl(0)
15550           arr_varLot(L_CADD, lngE) = CDbl(0)
15560           arr_varLot(L_REM, lngE) = CDbl(varRemaining)

15570         Else
                ' ** Last one.

15580           rst![PCash] = (varPCash * (varRemaining / .ShareFaceSale))
15590           rst![ICash] = (varICash * (varRemaining / .ShareFaceSale))
15600           rst![Cost] = Format(((varRemaining * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), "Currency")
15610           gstrSaleCost = Format(((varRemaining * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), "Currency")
15620           gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.

15630           lngLots = lngLots + 1&
15640           lngE = lngLots - 1&
15650           ReDim Preserve arr_varLot(L_ELEMS, lngE)
15660           arr_varLot(L_JID, lngE) = CLng(0)
15670           arr_varLot(L_LOTNO, lngE) = lngLots
15680           arr_varLot(L_ACTNO, lngE) = CStr(varAccountNo)
15690           arr_varLot(L_ASTNO, lngE) = CLng(varAssetNo)
15700           arr_varLot(L_SHARE, lngE) = arr_varRec(R_SHRS, lngZ)
15710           arr_varLot(L_SCENT, lngE) = CCenti(arr_varRec(R_SHRS, lngZ), 4)
15720           arr_varLot(L_SADD, lngE) = CDbl(0)
15730           arr_varLot(L_ICASH, lngE) = CDbl(varICash * (varRemaining / .ShareFaceSale))
15740           arr_varLot(L_ICENT, lngE) = CCenti((varICash * (varRemaining / .ShareFaceSale)), 2)  ' ** Function: Below.
15750           arr_varLot(L_INEW, lngE) = CDbl(0)
15760           arr_varLot(L_IADD, lngE) = CDbl(0)
15770           arr_varLot(L_PCASH, lngE) = CDbl(varPCash * (varRemaining / .ShareFaceSale))
15780           arr_varLot(L_PCENT, lngE) = CCenti((varPCash * (varRemaining / .ShareFaceSale)), 2)  ' ** Function: Below.
15790           arr_varLot(L_PNEW, lngE) = CDbl(0)
15800           arr_varLot(L_PADD, lngE) = CDbl(0)
15810           arr_varLot(L_COST, lngE) = CDbl((varRemaining * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1)
15820           arr_varLot(L_CCENT, lngE) = CCenti(((varRemaining * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), 2)
15830           arr_varLot(L_CNEW, lngE) = CDbl(0)
15840           arr_varLot(L_CADD, lngE) = CDbl(0)
15850           arr_varLot(L_SNEW, lngE) = CDbl(varRemaining)
15860           arr_varLot(L_REM, lngE) = CDbl(0)
15870           rst![shareface] = varRemaining
15880           varRemaining = 0

15890         End If

              ' ** VGC 12/28/2009: If they've checked IsAverage, put it in the Journal too; just 'cause I'd like to see it!
15900         If .opgChoice = .opgChoice_optAverage.OptionValue Then
15910           rst![IsAverage] = True
15920         End If

15930         With rst
15940           If IsNull(frm.description) = False Then
15950             If frm.description <> vbNullString Then
15960               ![description] = frm.description
15970             End If
15980           End If
15990           ![FocusHolder] = Null
16000           ![Calendar1] = Null
16010           ![Calendar2] = Null
16020           ![Recur_Name] = Null
16030           ![Recur_Type] = Null
16040           ![RecurringItem_ID] = Null
16050           ![pershare] = ((Abs(![ICash]) + Abs(![PCash]) / ![shareface]))
16060           If dblPerShare_Loop = 0# Then
16070             dblPerShare_Loop = ![pershare]
16080           End If
16090           ![Reinvested] = False
16100           ![PrintCheck] = False
16110           ![CheckNum] = Null
16120           ![Location_ID] = frm.Location_ID
16130           ![Loc_Name] = frm.Loc_Name
16140           ![Loc_Name_display] = frm.Loc_Name_display
16150           ![revcode_ID] = frm.revcode_ID
16160           ![revcode_DESC] = frm.revcode_DESC
16170           ![revcode_DESC_display] = frm.revcode_DESC_display
16180           ![revcode_TYPE] = frm.revcode_TYPE
16190           ![taxcode] = frm.taxcode
16200           ![taxcode_description] = frm.taxcode_description
16210           ![taxcode_description_display] = frm.taxcode_description_display
16220           ![taxcode_type] = frm.taxcode_type
16230           ![rate] = 0#
16240           ![due] = Null
16250           ![assettype] = Null
16260           ![JrnlCol_DateModified] = Now()
16270           .Update
16280           .Bookmark = .LastModified
16290         End With
              'TBL: tblJournal_Column  FLDS: 48
              'Journal_ID                   True
              'JrnlCol_ID {AutoNumber}      False
              'FocusHolder                  False
              'posted                       True
              'journaltype                  False
              'journalSubtype               False
              'journaltype_sortorder        False
              'transdate                    False
              'Calendar1                    False
              'accountno                    False
              'shortname                    False
              'assetdate                    False
              'assetdate_display            False
              'Calendar2                    False
              'purchaseDate                 False
              'assetno                      False
              'assetno_description          False
              'Recur_Name                   False
              'Recur_Type                   False
              'RecurringItem_ID             False
              'shareface                    False
              'pershare                     False
              'icash                        False
              'pcash                        False
              'cost                         False
              'Reinvested                   True
              'PrintCheck                   True
              'CheckNum                     False
              'JrnlMemo_Memo                False
              'JrnlMemo_HasMemo             True
              'description                  False
              'Location_ID                  True
              'Loc_Name                     False
              'Loc_Name_display             False
              'revcode_ID                   True
              'revcode_DESC                 False
              'revcode_DESC_display         False
              'revcode_TYPE                 False
              'taxcode                      True
              'taxcode_description          False
              'taxcode_description_display  False
              'TaxCode_Type                 False
              'journal_USER                 True
              'rate                         False
              'due                          False
              'assettype                    False
              'IsAverage                    True
              'JrnlCol_DateModified         True

16300         lngNews = lngNews + 1&  ' ** Subsequent records created in sale.
16310         ReDim Preserve arr_lngNew(lngNews - 1&)
16320         arr_lngNew(lngNews - 1&) = rst![JrnlCol_ID]
16330         arr_varLot(L_JID, lngE) = rst![JrnlCol_ID]

              ' ** Add the subsequent records to the array.
16340         frm.NewRecAdd rst![JrnlCol_ID]  ' ** Form Procedure: frmJournal_Columns_Sub.

16350       Else
16360         Exit Do
16370       End If  ' ** blnRetVal.

16380     End With  ' ** Me.

16390   Loop  ' ** varRemaining.

16400   rst.Close

16410   If dblPerShare_First <> dblPerShare_Loop Then
16420     frm.pershare = dblPerShare_Loop
16430   End If

16440 On Error GoTo ERRH

16450   If blnRetVal = True Then
16460     blnRetVal = DoMultiLots_Round  ' ** Function: Below.
16470   End If  ' ** blnRetVal.

16480   If blnRetVal = True Then
16490     With dbs
16500       Set rst = .OpenRecordset("tblJournal_Column", dbOpenDynaset, dbConsistent)
16510       With rst
16520         For lngX = 0& To (lngLots - 1&)
16530           If arr_varLot(L_JID, lngX) = frm.JrnlCol_ID Then
16540             If arr_varLot(L_INEW, lngX) <> 0# Then
16550               frm.ICash = CCur(arr_varLot(L_INEW, lngX) + arr_varLot(L_IADD, lngX))
16560             ElseIf Nz(frm.ICash, 0@) <> 0@ Then
16570               frm.ICash = 0@
16580             End If
16590             If arr_varLot(L_PNEW, lngX) <> 0# Then
16600               frm.PCash = CCur(arr_varLot(L_PNEW, lngX) + arr_varLot(L_PADD, lngX))
16610             ElseIf Nz(frm.PCash, 0@) <> 0@ Then
16620               frm.PCash = 0@
16630             End If
16640             If arr_varLot(L_CNEW, lngX) <> 0# Then
16650               frm.Cost = CCur(arr_varLot(L_CNEW, lngX) + arr_varLot(L_CADD, lngX))
16660             ElseIf Nz(![Cost], 0@) <> 0@ Then
16670               frm.Cost = 0@
16680             End If
16690             If arr_varLot(L_SNEW, lngX) <> 0# Then
16700               frm.shareface = Abs(arr_varLot(L_SNEW, lngX))
16710             ElseIf arr_varLot(L_SNEW, lngX) = 0# And Nz(![shareface], 0@) <> 0@ Then
16720               frm.shareface = 0@
16730             End If
16740           Else
16750             .FindFirst "[JrnlCol_ID] = " & CStr(arr_varLot(L_JID, lngX))
16760             If .NoMatch = False Then
16770               .Edit
16780               If arr_varLot(L_INEW, lngX) <> 0# Then
16790                 ![ICash] = CCur(arr_varLot(L_INEW, lngX) + arr_varLot(L_IADD, lngX))
16800               ElseIf Nz(![ICash], 0@) <> 0@ Then
16810                 ![ICash] = 0@
16820               End If
16830               If arr_varLot(L_PNEW, lngX) <> 0# Then
16840                 ![PCash] = CCur(arr_varLot(L_PNEW, lngX) + arr_varLot(L_PADD, lngX))
16850               ElseIf Nz(![PCash], 0@) <> 0@ Then
16860                 ![PCash] = 0@
16870               End If
16880               If arr_varLot(L_CNEW, lngX) <> 0# Then
16890                 ![Cost] = CCur(arr_varLot(L_CNEW, lngX) + arr_varLot(L_CADD, lngX))
16900               ElseIf Nz(![Cost], 0@) <> 0@ Then       ' ** Though this shouldn't be the case,
16910                 ![Cost] = 0@                          ' ** it'll certainly stick out if I've done something wrong!
16920               End If
16930               If arr_varLot(L_SNEW, lngX) <> 0# Then
16940                 ![shareface] = Abs(arr_varLot(L_SNEW, lngX))
16950               ElseIf arr_varLot(L_SNEW, lngX) = 0# And Nz(![shareface], 0@) <> 0@ Then
                      ' ** Though this shouldn't be the case, it'll certainly stick out if I've done something wrong!
16960                 ![shareface] = 0@
16970               End If
16980               .Update
16990             Else
17000               blnRetVal = False
17010               MsgBox "Journal entry not found.", vbInformation + vbOKOnly, "Multiple-Tax-Lot Error"
17020               Exit For
17030             End If
17040           End If
17050         Next
17060         .Close
17070       End With
17080       .Close
17090     End With
17100   End If  ' ** blnRetVal.

17110   frm.Requery

17120   If lngLots > 1& Then
17130     If blnRetVal = True Then
17140       blnMultiLotSale = True
17150       MsgBox "Sale of multiple tax lots." & Space(30), vbInformation + vbOKOnly, "Multiple Tax Lots"
17160     Else
17170       MsgBox "A problem occurred during multiple-tax-lot processing." & vbCrLf & _
              "Check the entries generated for accuracy.", vbExclamation + vbOKOnly, "Multiple-Tax-Lot Error"
17180     End If
17190   End If

EXITP:
17200   Set frmSub = Nothing
17210   Set frm = Nothing
17220   Set rst = Nothing
17230   Set qdf = Nothing
17240   Set dbs = Nothing
17250   DoMultiLots = blnRetVal
17260   Exit Function

ERRH:
17270   blnRetVal = False
17280   Select Case ERR.Number
        Case Else
17290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17300   End Select
17310   Resume EXITP

End Function

Private Function DoMultiLots_Round() As Boolean
' ** Now go through the array, rounding as appropriate.
' ** Remember, ShareFace CAN go to hundredths!

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "DoMultiLots_Round"

        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

17410   blnRetVal = True

        ' ** First, see if there are any tenths of a penny in the sale.
        ' ** The Centi elements hold hundredths of a penny, so 10 = $0.001.
17420   varShareface = CDbl(0): varICash = CDbl(0): varPCash = CDbl(0): varCost = CDbl(0)
17430   For lngX = 0& To (lngLots - 1&)
17440     varShareface = varShareface + arr_varLot(L_SCENT, lngX)
17450     varICash = varICash + arr_varLot(L_ICENT, lngX)
17460     varPCash = varPCash + arr_varLot(L_PCENT, lngX)
17470     varCost = varCost + arr_varLot(L_CCENT, lngX)
17480   Next

        ' *******************************************
        ' ** Array: arr_varLot()
        ' **
        ' **   Element  Name             Constant
        ' **   =======  ===============  ==========
        ' **      0     Lot Number       L_LOTNO
        ' **      1     Journal ID       L_JID
        ' **      2     accountno        L_ACTNO
        ' **      3     assetno          L_ASTNO
        ' **      4     Shareface        L_SHARE
        ' **      5     100ths           L_SCENT
        ' **      6     Shareface New    L_SNEW
        ' **      7     Distribution     L_SADD
        ' **      8     ICash            L_ICASH
        ' **      9     100ths           L_ICENT
        ' **     10     ICash New        L_INEW
        ' **     11     Distribution     L_IADD
        ' **     12     PCash            L_PCASH
        ' **     13     100ths           L_PCENT
        ' **     14     PCash New        L_PNEW
        ' **     15     Distribution     L_PADD
        ' **     16     Cost             L_COST
        ' **     17     100ths           L_CCENT
        ' **     18     Cost New         L_CNEW
        ' **     19     Distribution     L_CADD
        ' **     20     Share Remain     L_REM
        ' **
        ' *******************************************

17490   If varICash <> 0 Or varPCash <> 0 Or varCost <> 0 Then
          ' ** Yes, there are tenths of a penny to deal with.

          ' ** Caclulate the totals of the raw numbers.
17500     dblShareTot = 0#: dblICashTot = 0#: dblPCashTot = 0#: dblCostTot = 0#
17510     For lngX = 0& To (lngLots - 1&)
17520       If lngX < (lngLots - 1&) Then
17530         dblShareTot = dblShareTot + arr_varLot(L_SHARE, lngX)
17540       Else
              ' ** The last record should have, in L_SNEW, the remaining portion being
              ' ** taken from L_SHARE, which is that lot's original ActiveAssets number.
              ' ** It's the only element in the array with a non-zero L_SNEW.
              ' ** (At least it should be there, not anywhere else, and not missing.)
17550         dblShareTot = dblShareTot + arr_varLot(L_SNEW, lngX)
              ' ** arr_varLot(L_SNEW, lngE) = CDbl(varRemaining)
              ' ** arr_varLot(L_REM, lngE) = CDbl(0)
              ' ** rst![shareface] = varRemaining
              ' ** varRemaining = 0
17560       End If
17570       dblICashTot = dblICashTot + arr_varLot(L_ICASH, lngX)
17580       dblPCashTot = dblPCashTot + arr_varLot(L_PCASH, lngX)
17590       dblCostTot = dblCostTot + arr_varLot(L_COST, lngX)
17600     Next

          ' ** Now round the totals.
          ' ** These will be the numbers I have to end up with.
17610     dblShareTot = Round(dblShareTot, 4)
17620     dblShareTot = (Val(CStr(dblShareTot * 10000)) / 10000)
17630     dblICashTot = Round(dblICashTot, 2)
17640     dblICashTot = (Val(CStr(dblICashTot * 100)) / 100)
17650     dblPCashTot = Round(dblPCashTot, 2)
17660     dblPCashTot = (Val(CStr(dblPCashTot * 100)) / 100)
17670     dblCostTot = Round(dblCostTot, 2)
17680     dblCostTot = (Val(CStr(dblCostTot * 100)) / 100)

          ' ** First, just lop off the tenths-on from the values, regardless of where they might round.
17690     dblShareNew = 0#: dblICashNew = 0#: dblPCashNew = 0#: dblCostNew = 0#
17700     For lngX = 0& To (lngLots - 1&)

17710       If arr_varLot(L_ICENT, lngX) <> 0 Then
17720         strTmp01 = CStr(arr_varLot(L_ICASH, lngX))
17730         intPos01 = InStr(strTmp01, ".")
17740         strTmp01 = Left(strTmp01, (intPos01 + 2))
17750         arr_varLot(L_INEW, lngX) = CDbl(strTmp01)
17760       Else
17770         arr_varLot(L_INEW, lngX) = arr_varLot(L_ICASH, lngX)
17780       End If

17790       If arr_varLot(L_PCENT, lngX) <> 0 Then
17800         strTmp01 = CStr(arr_varLot(L_PCASH, lngX))
17810         intPos01 = InStr(strTmp01, ".")
17820         strTmp01 = Left(strTmp01, (intPos01 + 2))
17830         arr_varLot(L_PNEW, lngX) = CDbl(strTmp01)
17840       Else
17850         arr_varLot(L_PNEW, lngX) = arr_varLot(L_PCASH, lngX)
17860       End If

17870       If arr_varLot(L_CCENT, lngX) <> 0 Then
17880         strTmp01 = CStr(arr_varLot(L_COST, lngX))
17890         intPos01 = InStr(strTmp01, ".")
17900         strTmp01 = Left(strTmp01, (intPos01 + 2))
17910         arr_varLot(L_CNEW, lngX) = CDbl(strTmp01)
17920       Else
17930         arr_varLot(L_CNEW, lngX) = arr_varLot(L_COST, lngX)
17940       End If

17950     Next

          ' ** Now figure out where to put the pennies!
          ' ** dblCostTot = The real final total, to which the pieces must add up.
          ' ** dblCostNew = The total of rounded pieces, which will be less than above.
          ' ** varCost    = The total remainders, as hundredths of a penny, so 10 = $0.001.
          '? CDbl(arr_varLot(L_PNEW, lngX) * 100#) = CLng(arr_varLot(L_PNEW, lngX) * 100#)
          'False
          '? CDbl(Val(CStr(arr_varLot(L_PNEW, lngX) * 100#))) = CLng(Val(CStr(arr_varLot(L_PNEW, lngX) * 100#)))
          'True
          '? dblPCashTot - dblPCashNew
          '1.00000000002183E-02
          '? Rem_ENotation(dblPCashTot - dblPCashNew)
          '0.01000000000021
          '? dblPCashTot
          ' 5009.33
          '? dblPCashNew
          ' 5009.32
          '? CDbl((Val(CStr(dblPCashTot * 100#)) - Val(CStr(dblPCashNew * 100#))) / 100)
          ' 0.01
17960     For lngX = 0& To (lngLots - 1&)
17970       arr_varLot(L_INEW, lngX) = (Val(CStr(arr_varLot(L_INEW, lngX) * 100#)) / 100)
17980       arr_varLot(L_PNEW, lngX) = (Val(CStr(arr_varLot(L_PNEW, lngX) * 100#)) / 100)
17990       arr_varLot(L_CNEW, lngX) = (Val(CStr(arr_varLot(L_CNEW, lngX) * 100#)) / 100)
18000     Next

18010     For lngX = 0& To (lngLots - 1&)
18020       dblICashNew = dblICashNew + arr_varLot(L_INEW, lngX)
18030       dblPCashNew = dblPCashNew + arr_varLot(L_PNEW, lngX)
18040       dblCostNew = dblCostNew + arr_varLot(L_CNEW, lngX)
18050       If lngX < (lngLots - 1&) Then
18060         dblShareNew = dblShareNew + arr_varLot(L_SHARE, lngX)
18070       Else
18080         dblShareNew = dblShareNew + arr_varLot(L_SNEW, lngX)
18090       End If
18100     Next

18110     dblShareNew = (Val(CStr(dblShareNew * 10000)) / 10000)
18120     dblICashNew = (Val(CStr(dblICashNew * 100)) / 100)
18130     dblPCashNew = (Val(CStr(dblPCashNew * 100)) / 100)
18140     dblCostNew = (Val(CStr(dblCostNew * 100)) / 100)

          'dblCostTot
          'dblCostNew
          'varCost
          'varRemaining = dblCostTot - dblCostNew
18150     varRemaining = CDbl((Val(CStr(dblCostTot * 100#)) - Val(CStr(dblCostNew * 100#))) / 100)
          '' ** Lessen precision to remove E-Notation.
          'varRemaining = Rem_ENotation(varRemaining)  ' ** Module Function: modStringFuncs.
18160     If varRemaining <> 0# Then
            ' ** varRemaining should be just a few cents.
            ' ** Do I want to make judgements as to which ones get the penny?
            ' ** That is, should those > $0.005 get them first?
            ' ** I could sort the array in descending order, thus giving precedence?
            ' ** Yah, go for it!
18170       DoMultiLots_Sort L_COST  ' ** Procedure: Below.
18180       Do While varRemaining <> 0&
18190         For lngX = (lngLots - 1&) To 0& Step -1&
                ' ** If Number Is  Sgn() Returns
                ' ** ============  ===================
                ' **     -1        Less than zero
                ' **      0        Equal to zero
                ' **      1        Greater than zero
18200           Select Case Sgn(varRemaining)
                Case -1
18210             arr_varLot(L_CADD, lngX) = arr_varLot(L_CADD, lngX) - 0.01
18220             varRemaining = varRemaining + 0.01
18230           Case Else
18240             arr_varLot(L_CADD, lngX) = arr_varLot(L_CADD, lngX) + 0.01
18250             varRemaining = varRemaining - 0.01
18260           End Select
18270           varRemaining = Round(varRemaining, 2)
                ' ** Lessen precision to remove E-Notation.
18280           varRemaining = Rem_ENotation(varRemaining)  ' ** Module Function: modStringFuncs.
18290           varRemaining = Val(CStr(varRemaining))
18300           If varRemaining = 0# Then
18310             Exit For
18320           End If
18330         Next
18340       Loop
18350     End If

          'dblICashTot
          'dblICashNew
          'varICash
          'varRemaining = dblICashTot - dblICashNew
18360     varRemaining = CDbl((Val(CStr(dblICashTot * 100#)) - Val(CStr(dblICashNew * 100#))) / 100)
          '' ** Lessen precision to remove E-Notation.
          'varRemaining = Rem_ENotation(varRemaining)  ' ** Module Function: modStringFuncs.
18370     If varRemaining <> 0# Then
18380       DoMultiLots_Sort L_ICASH  ' ** Procedure: Below.
18390       Do While varRemaining <> 0&
18400         For lngX = (lngLots - 1&) To 0& Step -1&
18410           Select Case Sgn(varRemaining)
                Case -1
18420             arr_varLot(L_IADD, lngX) = arr_varLot(L_IADD, lngX) - 0.01
18430             varRemaining = varRemaining + 0.01
18440           Case Else
18450             arr_varLot(L_IADD, lngX) = arr_varLot(L_IADD, lngX) + 0.01
18460             varRemaining = varRemaining - 0.01
18470           End Select
18480           varRemaining = Round(varRemaining, 2)
                ' ** Lessen precision to remove E-Notation.
18490           varRemaining = Rem_ENotation(varRemaining)  ' ** Module Function: modStringFuncs.
18500           varRemaining = Val(CStr(varRemaining))
18510           If varRemaining = 0# Then
18520             Exit For
18530           End If
18540         Next
18550       Loop
18560     End If

          'dblPCashTot
          'dblPCashNew
          'varPCash
          'varRemaining = dblPCashTot - dblPCashNew
          'WAIT A MINUTE! SHOULDN'T TOT AND NEW ALREADY BE ROUNDED TO 2 DECIMAL PLACES?!
          'YES! THERE SHOULD BE NO TENTHS OF A PENNY HERE!
18570     varRemaining = CDbl((Val(CStr(dblPCashTot * 100#)) - Val(CStr(dblPCashNew * 100#))) / 100)
          '' ** Lessen precision to remove E-Notation.
          'varRemaining = Rem_ENotation(varRemaining)  ' ** Module Function: modStringFuncs.
18580     If varRemaining <> 0# Then
18590       DoMultiLots_Sort L_PCASH  ' ** Procedure: Below.
18600       Do While varRemaining <> 0#
18610         For lngX = (lngLots - 1&) To 0& Step -1&
18620           Select Case Sgn(varRemaining)
                Case -1
18630             arr_varLot(L_PADD, lngX) = arr_varLot(L_PADD, lngX) - 0.01
18640             varRemaining = varRemaining + 0.01
18650           Case Else
18660             arr_varLot(L_PADD, lngX) = arr_varLot(L_PADD, lngX) + 0.01
18670             varRemaining = varRemaining - 0.01
18680           End Select
18690           varRemaining = Round(varRemaining, 2)
                ' ** Lessen precision to remove E-Notation.
18700           varRemaining = Rem_ENotation(varRemaining)  ' ** Module Function: modStringFuncs.
18710           varRemaining = Val(CStr(varRemaining))
18720           If varRemaining = 0# Then
18730             Exit For
18740           End If
18750         Next
18760       Loop
18770     End If

          ' ** Put the array back into standard, lot number sort.
18780     DoMultiLots_Sort L_LOTNO  ' ** Procedure: Below.

          ' ** dblShareTot = Full total, with last portion.
          ' ** dblShareNew = Same as above.
          ' ** varShareface    = The total remainders, as millionths of a share, so 10 = 0.00001.

          ' ** Do I need to do anything with ShareFace?
          ' ** The individual records need to stay exactly as they were
          ' ** in ActiveAssets, so that they zero-out during Posting.
          ' ** The last one, however, shouldn't leave ActiveAssets
          ' ** with some extreme decimal. So, see about selling
          ' ** some number that will leave ActiveAssets cleanly
          ' ** rounded to 4 decimal places.
          ' **
          ' ** L_SCENT, L_SNEW, and L_SADD are all stil 0's,
          ' ** with the ActiveAssets original in L_SHARE.
          'dblShareTot
          'dblShareNew
          'varShareface
18790     For lngX = 0& To (lngLots - 1&)
18800       If arr_varLot(L_SNEW, lngX) = 0# Then
18810         If lngX < (lngLots - 1&) Then
18820           arr_varLot(L_SNEW, lngX) = arr_varLot(L_SHARE, lngX)
18830         Else
18840           arr_varLot(L_SNEW, lngX) = arr_varLot(L_SHARE, lngX)  ' ** 07/06/2009: This is the line that was missing.
                ' ** I have no idea why it didn't show up before. Harumph!
                ' ** This added another ten-thousandth to the total ShareFace, making
                ' ** it 0.0001 off from the entered amount (836.1193 vs 836.1194).
                ' ** I guess subsequent transactions will just have to deal with that.
                '' ** Special handling for the final record in the multi-lot sale.
                'dblTmp03 = arr_varLot(L_SHARE, lngX)  ' ** ActiveAssets' original value.
                'dblTmp04 = arr_varLot(L_SNEW, lngX)  ' ** The last portion, to be taken from 02, leaving 02 - 03 (04) still on the books.
                'dblTmp05 = dblTmp03 - dblTmp04
                'strTmp01 = CStr(dblTmp05)
                'If InStr(strTmp01, "E") > 0 Then
                '  ' ** Lessen precision to remove E-Notation.
                '  dblTmp05 = Rem_ENotation(dblTmp03 - dblTmp04)  ' ** Module Function: modStringFuncs.
                '  strTmp01 = CStr(dblTmp05)
                'End If
                'intPos01 = InStr(strTmp01, ".")
                'If intPos01 > 0 Then
                '  strTmp02 = Mid(strTmp01, (intPos01 + 1))  ' ** Just the decimal portion of the shares.
                '  If Len(strTmp02) > 4 Then
                '    strTmp02 = Mid(strTmp02, 5)
                '    strTmp02 = ".0000" & strTmp02
                '    dblTmp06 = Val(strTmp02)
                '    ' ** If this amount is added to the last journal's record,
                '    ' ** then the Posting subtraction should leave ActiveAssets
                '    ' ** with just 4 decimal places.
                '    arr_varLot(L_SNEW, lngX) = arr_varLot(L_SNEW, lngX) + dblTmp06
                '  Else
                '    ' ** Let it stand.
                '  End If
                'Else
                '  ' ** Let it stand.
                'End If

18850         End If
18860       End If
18870     Next
          ' ** arr_varLot(L_SNEW, lngE) = CDbl(varRemaining)
          ' ** arr_varLot(L_REM, lngE) = CDbl(0)
          ' ** rst![shareface] = varRemaining
          ' ** varRemaining = 0

          ' ** OK. Now all NEW's have a rounded, unadjusted amount (except L_SNEW).
          ' ** All ADD's have the pennies to be added to NEW, as fully decimalized Doubles.
          ' ** So now those Journal records all have to be updated with these NEW + ADD's.

          'Debug.Print "'dblShareTot: " & CStr(dblShareTot)
          'Debug.Print "'dblShareNew: " & CStr(dblShareNew)
          'Debug.Print "'dblICashTot: " & CStr(dblICashTot)
          'Debug.Print "'dblICashNew: " & CStr(dblICashNew)
          'Debug.Print "'dblPCashTot: " & CStr(dblPCashTot)
          'Debug.Print "'dblPCashNew: " & CStr(dblPCashNew)
          'Debug.Print "'dblCostTot: " & CStr(dblCostTot)
          'Debug.Print "'dblCostNew: " & CStr(dblCostNew)
          'dblShareTot: 836.1193
          'dblShareNew: 836.1193
          'dblICashTot: 0
          'dblICashNew: 0
          'dblPCashTot: 836.12
          'dblPCashNew: 835.9
          'dblCostTot: -15861.83
          'dblCostNew: -15861.82

18880   Else

          ' ** If rounding isn't necessary, just copy the original values.
18890     For lngX = 0& To (lngLots - 1&)
18900       If arr_varLot(L_SNEW, lngX) = 0# Then
18910         arr_varLot(L_SNEW, lngX) = arr_varLot(L_SHARE, lngX)
18920       End If
18930       If arr_varLot(L_INEW, lngX) = 0@ Then
18940         arr_varLot(L_INEW, lngX) = arr_varLot(L_ICASH, lngX)
18950       End If
18960       If arr_varLot(L_PNEW, lngX) = 0@ Then
18970         arr_varLot(L_PNEW, lngX) = arr_varLot(L_PCASH, lngX)
18980       End If
18990       If arr_varLot(L_CNEW, lngX) = 0@ Then
19000         arr_varLot(L_CNEW, lngX) = arr_varLot(L_COST, lngX)
19010       End If
19020     Next

19030   End If

EXITP:
19040   DoMultiLots_Round = blnRetVal
19050   Exit Function

ERRH:
19060   blnRetVal = False
19070   Select Case ERR.Number
        Case Else
19080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19090   End Select
19100   Resume EXITP

End Function

Private Function CCenti(varInput As Variant, intPrec As Integer) As Double
' ** I wanted to use CMill(), but that's really only
' ** tenths of a penny, and I'm multiplying by 100.
' ** Darn!

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "CCenti"

        Dim strTmp01 As String, dblTmp02 As Double
        Dim intPos01 As Integer
        Dim dblRetVal As Double

19210   dblRetVal = 0#

19220   If IsNull(varInput) = False Then
19230     If IsNumeric(varInput) = True Then
19240       If InStr(CStr(varInput), "E") = 0 Then  ' ** Anything producing scientific notation is to small to worry about.
19250         dblTmp02 = varInput
19260         Select Case intPrec
              Case 2
19270           dblTmp02 = dblTmp02 * 100#
19280         Case 4
19290           dblTmp02 = dblTmp02 * 10000#
19300         End Select
19310         strTmp01 = CStr(dblTmp02)
19320         intPos01 = InStr(strTmp01, ".")
19330         If intPos01 > 0 Then
19340           strTmp01 = Mid(strTmp01, (intPos01 + 1))  ' ** Everything from the 3rd/5th decimal place and to the right.
19350           Select Case Len(strTmp01)
                Case 1
                  ' ** Limited to tenths of a penny.
19360             strTmp01 = strTmp01 & "0"
19370             dblRetVal = Val(strTmp01)
19380           Case 2
                  ' ** Hundredths of a penny.
19390             dblRetVal = Val(strTmp01)
19400           Case Else
                  ' ** Thousandths of a penny, etc.
19410             strTmp01 = Left(strTmp01, 2) & "." & Mid(strTmp01, 3)
19420             dblRetVal = Val(strTmp01)
                  ' ** 1. 72304.734288566  Original result
                  ' ** 2. 7230473.4288566  x 100
                  ' ** 3. 4288566
                  ' ** 4. 42.88566
19430           End Select
19440         Else
                ' ** Precision already no more than 2 decimal places.
19450         End If
19460       Else
              ' ** Too small. Throw it out!
19470       End If
19480     End If
19490   End If

EXITP:
19500   CCenti = dblRetVal
19510   Exit Function

ERRH:
19520   dblRetVal = 0#
19530   Select Case ERR.Number
        Case Else
19540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19550   End Select
19560   Resume EXITP

End Function

Private Sub DoMultiLots_Sort(intField As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "DoMultiLots_Sort"

        Dim dblSortA As Double, dblSortB As Double
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long

        ' *******************************************
        ' ** Array: arr_varLot()
        ' **
        ' **   Element  Name             Constant
        ' **   =======  ===============  ==========
        ' **      0     Lot Number       L_LOTNO
        ' **      1     Journal ID       L_JID
        ' **      2     accountno        L_ACTNO
        ' **      3     assetno          L_ASTNO
        ' **      4     Shareface        L_SHARE
        ' **      5     100ths           L_SCENT
        ' **      6     Shareface New    L_SNEW
        ' **      7     Distribution     L_SADD
        ' **      8     ICash            L_ICASH
        ' **      9     100ths           L_ICENT
        ' **     10     ICash New        L_INEW
        ' **     11     Distribution     L_IADD
        ' **     12     PCash            L_PCASH
        ' **     13     100ths           L_PCENT
        ' **     14     PCash New        L_PNEW
        ' **     15     Distribution     L_PADD
        ' **     16     Cost             L_COST
        ' **     17     100ths           L_CCENT
        ' **     18     Cost New         L_CNEW
        ' **     19     Distribution     L_CADD
        ' **     20     Share Remain     L_REM
        ' **
        ' *******************************************

        ' ** Binary Sort arr_varLot() array.
19610   For lngX = UBound(arr_varLot, 2) To 1 Step -1
19620     For lngY = 0 To (lngX - 1)
19630       Select Case intField
            Case L_LOTNO
19640         dblSortA = arr_varLot(L_LOTNO, lngY)
19650         dblSortB = arr_varLot(L_LOTNO, (lngY + 1))
19660       Case L_SHARE
19670         dblSortA = arr_varLot(L_SHARE, lngY)
19680         dblSortB = arr_varLot(L_SHARE, (lngY + 1))
19690       Case L_ICASH
19700         dblSortA = arr_varLot(L_ICASH, lngY)
19710         dblSortB = arr_varLot(L_ICASH, (lngY + 1))
19720       Case L_PCASH
19730         dblSortA = arr_varLot(L_PCASH, lngY)
19740         dblSortB = arr_varLot(L_PCASH, (lngY + 1))
19750       Case L_COST
19760         dblSortA = arr_varLot(L_COST, lngY)
19770         dblSortB = arr_varLot(L_COST, (lngY + 1))
19780       End Select
19790       If dblSortA > dblSortB Then
19800         For lngZ = 0& To L_ELEMS
19810           varTmp00 = arr_varLot(lngZ, lngY)
19820           arr_varLot(lngZ, lngY) = arr_varLot(lngZ, (lngY + 1))
19830           arr_varLot(lngZ, (lngY + 1)) = varTmp00
19840         Next
19850       End If
19860     Next
19870   Next

EXITP:
19880   Exit Sub

ERRH:
19890   Select Case ERR.Number
        Case Else
19900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19910   End Select
19920   Resume EXITP

End Sub

Private Sub JColLocUpdate(blnIsMulti As Boolean, lngThisLocID As Long)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "JColLocUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim lngLocID As Long, strLocName As String

20010   DoCmd.Hourglass True
20020   DoEvents

        ' ** For Multi's:
        ' **   If there's only 1's and one other, use the other.
        ' **   If there's more than one other, leave it Unassigned.
        ' ** For Singles:
        ' **   One Tax Lot, so use this, unless, if it's 1, and
        ' **   others do have a real one, follow above's procedure.

20030   Set dbs = CurrentDb
20040   With dbs
          ' ** qryJournal_Columns_23_03 (tmpEdit04, grouped by accountno, assetno, Location_ID),
          ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max.
20050     Set qdf = .QueryDefs("qryJournal_Columns_23_04")
20060     Set rst = qdf.OpenRecordset
20070     With rst
20080       If .BOF = True And .EOF = True Then
              ' ** Shouldn't happen.
20090       Else
20100         .MoveFirst
20110         Select Case ![cnt]
              Case 1  ' ** All Tax Lots have the same Location_ID.
20120           Select Case blnIsMulti
                Case True
20130             lngLocID = ![Location_ID_max]  ' ** They'll both be the same.
20140           Case False
20150             If ![Location_ID_max] <> lngThisLocID And lngThisLocID > 0& Then
20160               lngLocID = lngThisLocID  ' ** Perhaps they changed it here?
20170             Else
20180               lngLocID = ![Location_ID_max]  ' ** They'll both be the same.
20190             End If
20200           End Select
20210         Case 2  ' ** A total of 2 different Location_ID's among all Tax Lots.
20220           Select Case blnIsMulti
                Case True
20230             If ![Location_ID_min] = 1& Then
20240               lngLocID = ![Location_ID_max]
20250             Else
20260               lngLocID = 1&  ' ** {Unassigned}, {no entry}.
20270             End If
20280           Case False
20290             If lngThisLocID = 1& Then
20300               If ![Location_ID_min] = 1& Then
20310                 lngLocID = ![Location_ID_max]
20320               Else
20330                 lngLocID = 1&  ' ** {Unassigned}, {no entry}.
20340               End If
20350             ElseIf lngThisLocID > 1& Then
20360               lngLocID = lngThisLocID
20370             Else
20380               lngLocID = 1&  ' ** {Unassigned}, {no entry}.
20390             End If
20400           End Select
20410         Case Else  ' ** Several different Location_ID's.
20420           Select Case blnIsMulti
                Case True
20430             lngLocID = 1&  ' ** {Unassigned}, {no entry}.
20440           Case False
20450             If lngThisLocID > 1& Then
20460               lngLocID = lngThisLocID
20470             Else
20480               lngLocID = 1&  ' ** {Unassigned}, {no entry}.
20490             End If
20500           End Select
20510         End Select
20520       End If
20530       .Close
20540     End With
20550     .Close
20560   End With

20570   Set frm = Forms("frmJournal_Columns").frmJournal_Columns_Sub.Form
20580   With frm
20590     If .Location_ID <> lngLocID Then
20600       .Location_ID = lngLocID
20610       If lngLocID > 1& Then
20620         strLocName = DLookup("[Loc_Name]", "qryJournal_Columns_10_Location_01", "[Location_ID] = " & CStr(lngLocID))
20630         .Loc_Name_display = strLocName
20640         .Loc_Name = strLocName
20650       Else
20660         .Loc_Name_display = vbNullString
20670         .Loc_Name = "{no entry}"
20680       End If
20690     End If
20700   End With

EXITP:
20710   Set frm = Nothing
20720   Set rst = Nothing
20730   Set qdf = Nothing
20740   Set dbs = Nothing
20750   Exit Sub

ERRH:
20760   DoCmd.Hourglass False
20770   Select Case ERR.Number
        Case Else
20780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20790   End Select
20800   Resume EXITP

End Sub

Public Sub RecArray_Load()

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "RecArray_Load"

        Dim rst As DAO.Recordset
        Dim lngTmp01 As Long
        Dim lngX As Long, lngE As Long

20910   With Me

20920     lngRecs = 0&
20930     ReDim arr_varRec(R_ELEMS, 0)

20940     Set rst = .frmJournal_Columns_TaxLot_Sub.Form.RecordsetClone
20950     With rst
20960       .MoveLast
20970       lngTmp01 = .RecordCount
20980       .MoveFirst
20990       For lngX = 1& To lngTmp01
21000         lngRecs = lngRecs + 1&
21010         lngE = lngRecs - 1&
21020         ReDim Preserve arr_varRec(R_ELEMS, lngE)
              ' *************************************************
              ' ** Array: arr_varRec()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     tmpedit04_id    R_TID
              ' **     2       1     assetno         R_ASTNO
              ' **     3       2     shareface       R_SHRS
              ' **     4       3     assetdate       R_ADAT
              ' **     5       4     IsZero          R_ZERO
              ' **     6       5     cost            R_COST
              ' **     7       6     priceperunit    R_PPU
              ' **
              ' *************************************************
21030         arr_varRec(R_TID, lngE) = ![tmpedit04_id]
21040         arr_varRec(R_ASTNO, lngE) = ![assetno]
21050         arr_varRec(R_SHRS, lngE) = ![shareface]
21060         arr_varRec(R_ADAT, lngE) = ![assetdate]
21070         arr_varRec(R_ZERO, lngE) = ![IsZero]
21080         arr_varRec(R_COST, lngE) = ![Cost]
21090         Select Case ![IsAverage]
              Case True
21100           arr_varRec(R_PPU, lngE) = ![averagepriceperunit]
21110         Case False
21120           arr_varRec(R_PPU, lngE) = ![priceperunit]
21130         End Select
21140         If lngX < lngTmp01 Then .MoveNext
21150       Next
21160       .Close
21170     End With
21180     Set rst = Nothing
21190   End With

EXITP:
21200   Exit Sub

ERRH:
21210   Select Case ERR.Number
        Case Else
21220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21230   End Select
21240   Resume EXITP

End Sub

Public Function RecArray_Get() As Variant

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "RecArray_Get"

        Dim arr_varRetVal As Variant

21310   If lngRecs = 0& Then
21320     RecArray_Load  ' ** Procedure: Above.
21330   End If

21340   arr_varRetVal = arr_varRec

EXITP:
21350   RecArray_Get = arr_varRetVal
21360   Exit Function

ERRH:
21370   arr_varRetVal = Empty
21380   Select Case ERR.Number
        Case Else
21390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21400   End Select
21410   Resume EXITP

End Function

Private Sub IncludeCurrency(blnShow As Boolean)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long, lngTmp02 As Long

21510   With Me

21520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
21530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
21540     End If

21550     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
21560     lngMonitorNum = 1&: lngTmp02 = 0&
21570     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
21580     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

21590     Select Case blnShow
          Case True
21600       DoCmd.SelectObject acForm, THIS_NAME, False
21610       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
21620       DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, lngFrm_Height  'lngFrm_Top
21630       If lngMonitorNum > 1& Then
21640         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
21650       End If
21660       .Width = lngForm_Width
21670       .cmdCancel.Left = lngCancel_Left
21680       .cmdOK.Left = lngOK_Left
21690       .FocusHolder.Left = lngFocusHolder_Left
21700       .RecordCount_lbl.Left = lngRecCnt_Left
21710       .RecordCount_lbl_box.Left = lngRecCnt_Left
21720       .frmJournal_Columns_TaxLot_Sub.Width = lngSub_Width
21730       .frmJournal_Columns_TaxLot_Sub_box.Width = (.frmJournal_Columns_TaxLot_Sub.Width + (2& * lngTpp))
21740       .Header_vline01.Left = lngForm_Width
21750       .Header_vline02.Left = lngForm_Width
21760       .Detail_vline03.Left = lngForm_Width
21770       .Detail_vline04.Left = lngForm_Width
21780       .Footer_vline01.Left = lngForm_Width
21790       .Footer_vline02.Left = lngForm_Width
21800       .Header_hline01.Width = lngForm_Width
21810       .Header_hline02.Width = lngForm_Width
21820       .Detail_hline01.Width = lngDetailLine_Width
21830       .Detail_hline02.Width = lngDetailLine_Width
21840       .Detail_hline03.Width = lngForm_Width
21850       .Detail_hline04.Width = lngForm_Width
21860       .Footer_hline01.Width = lngForm_Width
21870       .Footer_hline02.Width = lngForm_Width
21880     Case False
21890       lngTmp01 = (lngCurrID_Width + lngCostUSD_Width)
21900       .RecordCount_lbl_box.Left = (lngRecCnt_Left - (lngTmp01 / 2))
21910       .RecordCount_lbl.Left = .RecordCount_lbl_box.Left
21920       .FocusHolder.Left = (lngFocusHolder_Left - lngTmp01)
21930       .cmdOK.Left = (lngOK_Left - lngTmp01)
21940       .cmdCancel.Left = (lngCancel_Left - lngTmp01)
21950       .frmJournal_Columns_TaxLot_Sub.Width = (lngSub_Width - lngTmp01)
21960       .frmJournal_Columns_TaxLot_Sub_box.Width = (.frmJournal_Columns_TaxLot_Sub.Width + (2& * lngTpp))
21970       .Header_vline01.Left = (lngForm_Width - lngTmp01)
21980       .Header_vline02.Left = .Header_vline01.Left
21990       .Detail_vline03.Left = (lngForm_Width - lngTmp01)
22000       .Detail_vline04.Left = .Header_vline01.Left
22010       .Footer_vline01.Left = .Header_vline01.Left
22020       .Footer_vline02.Left = .Header_vline01.Left
22030       .Header_hline01.Width = .Header_vline01.Left
22040       .Header_hline02.Width = .Header_vline01.Left
22050       .Detail_hline01.Width = (lngDetailLine_Width - lngTmp01)
22060       .Detail_hline02.Width = (lngDetailLine_Width - lngTmp01)
22070       .Detail_hline03.Width = .Header_vline01.Left
22080       .Detail_hline04.Width = .Header_vline01.Left
22090       .Footer_hline01.Width = .Header_vline01.Left
22100       .Footer_hline02.Width = .Header_vline01.Left
22110       .Width = (lngForm_Width - lngTmp01)
22120       DoCmd.SelectObject acForm, THIS_NAME, False
22130       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
22140       DoCmd.MoveSize (lngFrm_Left + (lngTmp01 / 2)), lngTmp02, (lngFrm_Width - lngTmp01), lngFrm_Height  'lngFrm_Top
22150       If lngMonitorNum > 1& Then
22160         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
22170       End If
22180     End Select

22190   End With

EXITP:
22200   Exit Sub

ERRH:
22210   Select Case ERR.Number
        Case Else
22220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22230   End Select
22240   Resume EXITP

End Sub
