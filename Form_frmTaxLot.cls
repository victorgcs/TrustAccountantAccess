VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTaxLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmTaxLot"

'VGC 03/22/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Load()
' ##   chkIncludeCurrency_AfterUpdate()
' #######################################

'##########################################
'CURRENCY NOT FINISHED!
'CHF NEEDS FIXING
'CLR_LTGRN ON VALUES & CURR_ID!
'CHECK AVERAGING!
'##########################################

' ** Shortcut Alt keys responsive from this form:
' **   Average:          A {opgChoice_optAverage}
' **   Cancel:           C {cmdCancel}
' **   FIFO:             F {opgChoice_optFIFO}
' **   High:             H {opgChoice_optHigh}
' **   Low:              L {opgChoice_optLow}
' **   OK:               O {cmdOK}
' **   Print:            P {cmdPrintReport}
' **   Shareface:        R {shareface on frmTaxLot_Sub}
' **   Specific:         S {opgChoice_optSpecific}
' **   Location:         T {Location_ID On frmTaxLot_Sub}
' **   Include Currency: U {chkIncludeCurrency}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmTaxLot_Sub}

' ** Cost calculation is based on the Cost showing in ActiveAssets, not MasterAsset.

' ************************************************************************************************
' ** VGC 01/13/2013:
' ** NOTE REQUIRED ZZ_'S IN THIS MOD:
' **   Data-Definition: Create table tmpAveragePrice.
' **   zz_qry_System_46_01
' **   Data-Definition: Create index [assetno], [accountno] PrimaryKey on table tmpAveragePrice.
' **   zz_qry_System_46_02
' ************************************************************************************************

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private varRemaining As Variant, varShareface As Variant, varPCash As Variant, varICash As Variant, varCost As Variant
Private dblShareTot As Double, dblICashTot As Double, dblPCashTot As Double, dblCostTot As Double
Private dblShareNew As Double, dblICashNew As Double, dblPCashNew As Double, dblCostNew As Double
Private lngRevcodeID As Long, lngTaxcode As Long, strDesc As String, strCurrSym As String, strFormat As String

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
Private Const L_CURID As Integer = 21
Private Const L_USD   As Integer = 22
Private Const L_UNEW  As Integer = 23

' ** Array: arr_varRec().
Private lngRecs As Long, arr_varRec() As Variant
Private Const R_ELEMS As Integer = 8  ' ** Array's first-element UBound().
Private Const R_TID   As Integer = 0
Private Const R_ASTNO As Integer = 1
Private Const R_SHRS  As Integer = 2
Private Const R_ADAT  As Integer = 3
Private Const R_ZERO  As Integer = 4
Private Const R_COST  As Integer = 5
Private Const R_USD   As Integer = 6
Private Const R_PPU   As Integer = 7
Private Const R_CURID As Integer = 8

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngForm_Width As Long, lngSub_Width As Long, lngDetailLine_Width As Long
Private lngCancel_Left As Long, lngOK_Left As Long, lngFocusHolder_Left As Long, lngRecCnt_Left As Long
Private lngCurrID_Width As Long, lngCostUSD_Width As Long, lngIncludeCurrency_Left As Long, lngChkBoxLbl_Offset As Long

Private intOpeningChoice As Integer, intLastChoice As Integer, blnFromSaleBtn As Boolean
Private blnGoingToReport As Boolean, blnGoneToReport As Boolean, lngPrintReportLbl_Left As Long, blnWasGTR As Boolean
Private blnMultiLotSale As Boolean, blnHasForEx As Boolean
Private lngLotCnt As Long, strAccountNo As String, lngAssetNo As Long, lngCurrID As Long
Private lngJrnlID As Long, dblShareface As Double, datAssetDateSale As Date, blnD4D As Boolean
Private dblPCash As Double, dblCost As Double, dblCurrRate2 As Double
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
                ' ** strCallingForm ~ ID ~ ShareFace ~ AssetDate ~ Dollar-for-Dollar

160             blnIsOpen = True

170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             lngJrnlID = CLng(Left(strTmp01, (intPos01 - 1)))
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
410             lngSub_Width = .frmTaxLot_Sub.Width
420             lngDetailLine_Width = .Detail_hline01.Width
430             lngCancel_Left = .cmdCancel.Left
440             lngOK_Left = .cmdOK.Left
450             lngFocusHolder_Left = .FocusHolder.Left
460             lngRecCnt_Left = .RecordCount_lbl_box.Left
470             lngCurrID_Width = (.frmTaxLot_Sub.Form.curr_id.Width + (4& * lngTpp))
480             lngCostUSD_Width = (.frmTaxLot_Sub.Form.cost_usd.Width + (4& * lngTpp))
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
630               Set qdf = .QueryDefs("qryLotInformation_27a")
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
820             .ShareFaceSale = dblShareface  ' ** From .saleShareFace on frmJournal_Sub4_Sold, original amount.

830             If strTmp01 <> "Null" Then
840               .AssetDateSale = datAssetDateSale
850               .AssetDateNull = False
860             Else
870               .AssetDateNull = True
880             End If

890             If lngCurrID <> 150& Then
900               .frmTaxLot_Sub.Form.Cost.Visible = False
910               .frmTaxLot_Sub.Form.cost_str.Visible = True
920               .frmTaxLot_Sub.Form.cost_sum_box.Visible = True
930               .frmTaxLot_Sub.Form.priceperunit_avg_box.Visible = True
940             Else
950               .frmTaxLot_Sub.Form.Cost.Visible = True
960               .frmTaxLot_Sub.Form.cost_str.Visible = False
970               .frmTaxLot_Sub.Form.cost_sum_box.Visible = False
980               .frmTaxLot_Sub.Form.priceperunit_avg_box.Visible = False
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

1340      Set rst = .frmTaxLot_Sub.Form.RecordsetClone
1350      rst.MoveFirst
          ' ** VGC 10/24/2014: If any are averaged, all should be, so taking just 1st record should be OK.
1360      If rst![IsAverage] = True Then
1370        .opgChoice = .opgChoice_optAverage.OptionValue
1380        .opgChoice_optFIFO.Enabled = False
1390        .opgChoice_optAverage.Enabled = True
1400        .opgChoice_optSpecific.Enabled = False
1410        .opgChoice_optHigh.Enabled = False
1420        .opgChoice_optLow.Enabled = False
1430        With .frmTaxLot_Sub.Form
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

1660      opgChoice_AfterUpdate  ' ** Procedure: below.
1670      DoEvents

1680      blnHasForEx = False
1690      RecArray_Load  ' ** Procedure: Below.

1700      If .frmTaxLot_Sub.Form.shareface = 0 Then
1710        For lngX = 0& To (lngRecs - 1&)
1720          If arr_varRec(R_ZERO, lngX) = False Then
1730            .frmTaxLot_Sub.Form.MoveRec 0, arr_varRec(R_TID, lngX)  ' ** Form Procedure: frmTaxLot_Sub.
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

1850      intOpeningChoice = .opgChoice
1860      blnFromSaleBtn = Forms(strCallingForm).frmJournal_Sub4_Sold.Form.FromSaleBtn

1870      If gblnGoToReport = True Then
1880        blnGoneToReport = True
1890        .cmdCancel.SetFocus
1900        .cmdOK.Enabled = False
1910      Else
1920        DoCmd.Hourglass False
1930        Select Case blnFromSaleBtn
            Case True
1940          .cmdCancel.SetFocus
1950          .cmdOK.Enabled = False
1960        Case False
1970          .cmdOK.SetFocus
1980        End Select
1990      End If

2000    End With

EXITP:
2010    Set rst = Nothing
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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Shareface:        R {shareface on frmTaxLot_Sub}
        ' **   Location:         T {Loc_Name_display on frmTaxLot_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmTaxLot_Sub}

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyEscape
2170        intRetVal = 0
2180        cmdCancel_Click  ' ** Procedure: Below.
2190      End Select
2200    End If

        ' ** Alt keys.
2210    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2220      Select Case intRetVal
          Case vbKeyR
2230        With Me
2240          intRetVal = 0
2250          If .frmTaxLot_Sub.Form.shareface.Enabled = True Then
2260            .frmTaxLot_Sub.SetFocus
2270            .frmTaxLot_Sub.Form.shareface.SetFocus
2280          Else
2290            Beep
2300          End If
2310        End With
2320      Case vbKeyT
2330        With Me
2340          intRetVal = 0
2350          .frmTaxLot_Sub.SetFocus
2360          .frmTaxLot_Sub.Form.Loc_Name_display.SetFocus
2370        End With
2380      End Select
2390    End If

        ' ** Ctrl keys.
2400    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2410      Select Case intRetVal
          Case vbKeyS
2420        With Me
2430          intRetVal = 0
2440          .frmTaxLot_Sub.SetFocus
2450          .frmTaxLot_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmTaxLot_Sub.
2460        End With
2470      End Select
2480    End If

        ' ** Ctrl-Shift keys.
2490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2500      Select Case intRetVal
          Case vbKeyF
2510        With Me
2520          intRetVal = 0
2530          .FocusHolder.SetFocus
2540        End With
2550      End Select
2560    End If

EXITP:
2570    KeyCode = intRetVal
2580    Exit Sub

ERRH:
2590    intRetVal = 0
2600    THAT_PROC = THIS_PROC
2610    That_Erl = Erl
2620    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
2630    Resume EXITP

End Sub

Private Sub Form_Timer()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2710    With Me
2720      .TimerInterval = False
2730      If gblnSetFocus = True Then
2740        gblnSetFocus = False
2750        strFormName = vbNullString
2760  On Error Resume Next
2770        Do Until strFormName = THIS_NAME
2780          DoCmd.SelectObject acForm, THIS_NAME, False
2790          strFormName = Screen.ActiveForm.Name
2800        Loop
2810  On Error GoTo ERRH
2820        Select Case blnFromSaleBtn
            Case True
2830          .cmdCancel.SetFocus
2840          .cmdOK.Enabled = False
2850        Case False
2860          .cmdOK.SetFocus
2870        End Select
2880      End If  ' ** gblnSetFocus
2890      If gblnGoToReport = True Then
2900        Select Case blnGoingToReport
            Case True
2910          DoCmd.Hourglass False  ' ** It should already be off.
2920          gblnGoToReport = False
2930          blnGoingToReport = False
2940          .GoToReport_arw_pr_img.Visible = False
2950          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
2960          DoEvents
2970        Case False
2980          DoCmd.Hourglass True  ' ** Make sure it's still running.
2990          DoEvents
3000          blnGoingToReport = True
3010          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
3020          .GoToReport_arw_pr_img.Visible = True
3030          Beep  '##MoveBeepT
3040          DoCmd.Hourglass False
3050          DoEvents
3060          Forms("frmJournal").GTREmblem_Set False  ' ** Form Procedure: frmJournal.
3070          .TimerInterval = (GTR_WAIT * 2&)
3080        End Select
3090      End If
3100    End With

EXITP:
3110    Exit Sub

ERRH:
3120    DoCmd.Hourglass False
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3210    If DataErr > 0 Then
3220      Select Case DataErr
          Case 13, 2108, 2110, 2116, 2169, 2501, 3020, 3314, 7753, 7787, 8519
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
3230        Response = acDataErrContinue
3240      Case 2046  ' ** The command or action isn't available now (first or last record).
3250        Response = acDataErrContinue
3260      Case Else
3270        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
3280        zErrorHandler "frmTaxLot", "Undefined Form Error", DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290      End Select
3300    End If

EXITP:
3310    Exit Sub

ERRH:
3320    Select Case ERR.Number
        Case Else
3330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3340    End Select
3350    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

3410    Select Case blnWasGTR
        Case True
3420      blnWasGTR = False
3430    Case False
          ' ** Don't save settings imposed by GoToReport!
3440      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
3450    End Select

3460    gstrFormQuerySpec = vbNullString

3470    If blnFromSaleBtn = True Then
3480      Forms(strCallingForm).frmJournal_Sub4_Sold.Form.FromSaleBtn = False
3490    End If

3500    If IsLoaded("rptListOfTaxLots", acReport) = True Then  ' ** Module Functions: modFileUtilities.
3510      DoCmd.Close acReport, "rptListOfTaxLots"
3520    End If

3530    DoCmd.Hourglass False

EXITP:
3540    Exit Sub

ERRH:
3550    Select Case ERR.Number
        Case Else
3560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3570    End Select
3580    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3610    With Me
3620      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
3630        Select Case blnPrintReport_Focus
            Case True
3640          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3650          .cmdPrintReport_raised_img.Visible = False
3660        Case False
3670          .cmdPrintReport_raised_img.Visible = True
3680          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3690        End Select
3700        .cmdPrintReport_raised_focus_img.Visible = False
3710        .cmdPrintReport_raised_focus_dots_img.Visible = False
3720        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3730        .cmdPrintReport_raised_img_dis.Visible = False
3740      End If
3750    End With

EXITP:
3760    Exit Sub

ERRH:
3770    Select Case ERR.Number
        Case Else
3780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3790    End Select
3800    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub cmdOK_Click()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim frm As Access.Form, frmSub As Access.Form
        Dim lngTmp01 As Long
        Dim lngX As Long, lngY As Long
        Dim blnContinue As Boolean, blnFound As Boolean

4010    blnContinue = True

4020    lngRevcodeID = 0&: lngTaxcode = 0&: strDesc = vbNullString

4030    Set frm = Forms(strCallingForm).frmJournal_Sub4_Sold.Form

4040    With Me

4050      Set frmSub = .frmTaxLot_Sub.Form

4060      If frmSub.shareface = 0 Then
4070        blnFound = False
4080        If .opgChoice = .opgChoice_optSpecific.OptionValue Then
4090          blnContinue = False
4100          Beep
4110          MsgBox "The chosen Tax Lot has already been sold," & vbCrLf & _
                "and is unavailable for this sale.", vbInformation + vbOKOnly, "Select Lot"
4120        Else
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
              ' **     7       6     cost_usd        R_USD
              ' **     8       7     priceperunit    R_PPU
              ' **     9       8     curr_id         R_CURID
              ' **
              ' *************************************************
4130          lngTmp01 = frmSub.tmpedit04_id
              ' ** This needs to be in the same order as the subform.
4140          For lngX = 0& To (lngRecs - 1&)
4150            If arr_varRec(R_TID, lngX) = lngTmp01 Then
4160              If lngX = (lngRecs - 1&) Then
                    ' ** We're at the end of the list.
4170                For lngY = 0& To (lngRecs - 1&)
4180                  If arr_varRec(R_TID, lngY) = lngTmp01 Then
                        ' ** They appear to all be sold!
4190                    blnContinue = False
4200                    Beep
4210                    MsgBox "All the Tax Lots have been previously sold," & vbCrLf & _
                          "and are unavailable for this sale.", vbInformation + vbOKOnly, "Nothing To Do"
4220                    .cmdCancel.SetFocus
4230                    DoEvents
4240                    .cmdOK.Enabled = False
4250                  Else
4260                    Select Case arr_varRec(R_ZERO, lngY)
                        Case True
                          ' ** Keep looking.
4270                    Case False
                          ' ** We'll stop here.
4280                      blnFound = True
4290                      frmSub.MoveRec 0, arr_varRec(R_TID, lngY)  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
4300                      DoEvents
4310                      Exit For
4320                    End Select
4330                  End If
4340                Next
                    ' ** If a suitable alternative was found, we can continue, below.
                    ' ** If none were found, that's all she wrote!
4350                Exit For
4360              Else
4370                Select Case arr_varRec(R_ZERO, (lngX + 1&))
                    Case True
                      ' ** Keep looking.
4380                Case False
                      ' ** We'll stop here.
4390                  blnFound = True
4400                  frmSub.MoveRec 0, arr_varRec(R_TID, lngX)  ' ** Form Procedure: frmJournal_Columns_TaxLot_Sub.
4410                  DoEvents
4420                  Exit For
4430                End Select
4440              End If
4450            End If
4460          Next
4470        End If
4480        If blnFound = False And blnContinue = True Then
              ' ** I think it should have hit one of the messages above.
              ' ** If no, oh well.
4490          blnContinue = False
4500          Beep
4510        End If
4520      End If

4530      If blnContinue = True Then

4540        lngRevcodeID = frm.cmbRevenueCodes
4550        lngTaxcode = frm.cmbTaxCodes
4560        If IsNull(frm.saleDescription) = False Then
4570          strDesc = frm.saleDescription
4580        End If

4590        If .ShareFaceSale > frmSub.shareface Then
              ' ** A greater number than is in this lot is requested for sale.
4600          If .opgChoice <> .opgChoice_optFIFO.OptionValue And .opgChoice <> .opgChoice_optAverage.OptionValue Then
                ' ** Includes Specific, High, and Low.
4610            If .opgChoice = .opgChoice_optHigh.OptionValue Or .opgChoice = .opgChoice_optLow.OptionValue Then
                  ' ** High or Low.
4620              blnContinue = DoMultiLots  ' ** Procedure: Below.
4630              If blnContinue = True Then
4640                frm.Requery
4650                JrnlLocUpdate True, 0&  ' ** Procedure: Below.
4660                gblnSaleChanged = False
4670                frm.saleCost_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
4680                DoEvents
4690                DoCmd.Close acForm, THIS_NAME
4700                gblnSetFocus = True
4710                Forms(strCallingForm).TimerInterval = 250&
4720                DoCmd.SelectObject acForm, strCallingForm, False
4730                Forms(strCallingForm).frmJournal_Sub4_Sold.SetFocus
4740                DoCmd.GoToRecord , , acNewRec
4750                frm.SaleChanged False  ' ** Form Procedure: frmJournal_Sub4_Sold.
4760              End If
4770            Else
                  ' ** Specific.
4780              MsgBox "You will need to select another lot with enough" & vbCrLf & _
                    "shares/face for this request, or select First In First Out.", vbInformation + vbOKOnly, "Select Lot"
4790            End If
4800          Else
                ' ** FIFO or Average.
4810            blnContinue = DoMultiLots  ' ** Procedure: Below.
4820            If blnContinue = True Then
4830              frm.Requery
4840              JrnlLocUpdate True, 0&  ' ** Procedure: Below.
4850              gblnSaleChanged = False
4860              frm.saleCost_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
4870              DoEvents
4880              DoCmd.Close acForm, THIS_NAME
4890              gblnSetFocus = True
4900              Forms(strCallingForm).TimerInterval = 250&
4910              DoCmd.SelectObject acForm, strCallingForm, False
4920              Forms(strCallingForm).frmJournal_Sub4_Sold.SetFocus
4930              DoCmd.GoToRecord , , acNewRec
4940              frm.SaleChanged False  ' ** Form Procedure: frmJournal_Sub4_Sold.
4950            End If
4960          End If
4970        Else
              ' ** Asset covers sale.
4980          If frm.saleShareFace <> 0 Then
4990            frm.saleCost = Format(((.ShareFaceSale * (frmSub.Cost / frmSub.shareface)) * -1), "Currency")
5000            gstrSaleCost = CStr(dblCost)
5010            gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
5020            frm.UnitPurchaseDate = frmSub.assetdate
5030            If IsNull(frm.saleAssetDate) Then
5040              frm.saleAssetDate = Now()
5050            End If
                ' ** VGC 12/28/2009: If they've checked IsAverage, put it in the Journal too; just 'cause I'd like to see it!
5060            If .opgChoice = .opgChoice_optAverage.OptionValue Then
5070              frm.saleIsAverage = True  'ISAVERAGE GETS PUT IN HERE!
5080            End If
5090          End If
5100          JrnlLocUpdate False, frmSub.Location_ID  ' ** Procedure: Below.
5110          frm.saleCost_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
5120          frm.cmdSaleLotInfo.Enabled = False
5130          DoEvents
5140          DoCmd.Close acForm, THIS_NAME
5150          gblnSetFocus = True
5160          Forms(strCallingForm).TimerInterval = 250&
5170          DoCmd.SelectObject acForm, strCallingForm, False
5180          Forms(strCallingForm).frmJournal_Sub4_Sold.SetFocus
5190          Forms(strCallingForm).frmJournal_Sub4_Sold.Form.SoldCostedSet True
5200        End If

5210      End If  ' ** blnContinue.

5220    End With

EXITP:
5230    Set frmSub = Nothing
5240    Set frm = Nothing
5250    Exit Sub

ERRH:
5260    DoCmd.Hourglass False
5270    Select Case ERR.Number
        Case 2488  ' ** You can't use the ApplyFilter action on this window.
          ' ** Skip the ApplyFilter error
5280    Case Else
5290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5300    End Select
5310    Resume EXITP

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
5570          If .chkIncludeCurrency.Enabled = True Then
5580            .chkIncludeCurrency.SetFocus
5590          ElseIf .cmdPrintReport.Enabled = True Then
5600            .cmdPrintReport.SetFocus
5610          Else
5620            .frmTaxLot_Sub.SetFocus
5630            .frmTaxLot_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmTaxLot_Sub.
5640            .frmTaxLot_Sub.Form.FocusHolder.SetFocus
5650          End If
5660        End With
5670      End Select
5680    End If

EXITP:
5690    KeyCode = intRetVal
5700    Exit Sub

ERRH:
5710    intRetVal = 0
5720    THAT_PROC = THIS_PROC
5730    That_Erl = Erl
5740    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5750    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

5810    If blnGoneToReport = True Then
5820      Forms("frmJournal").frmJournal_Sub4_Sold.Form.cmdSaleCancel_Click  ' ** Form Procedure: frmJournal_Sub4_Sold.
5830      DoEvents
5840    End If

5850    DoCmd.Close acForm, THIS_NAME
5860    gblnSetFocus = True
5870    DoCmd.SelectObject acForm, strCallingForm, False
5880    Forms(strCallingForm).frmJournal_Sub4_Sold.SetFocus
5890    Forms(strCallingForm).frmJournal_Sub4_Sold.Form.SoldCostedSet False
5900    If blnGoneToReport = False Then
5910      Forms(strCallingForm).TimerInterval = 250&
5920    End If

EXITP:
5930    Exit Sub

ERRH:
5940    Select Case ERR.Number
        Case Else
5950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5960    End Select
5970    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
6090          .shortname.SetFocus
6100        End With
6110      End Select
6120    End If

        ' ** Shift keys.
6130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6140      Select Case intRetVal
          Case vbKeyTab
6150        With Me
6160          intRetVal = 0
6170          If .cmdOK.Enabled = True Then
6180            .cmdOK.SetFocus
6190          ElseIf .chkIncludeCurrency.Enabled = True Then
6200            .chkIncludeCurrency.SetFocus
6210          ElseIf .cmdPrintReport.Enabled = True Then
6220            .cmdPrintReport.SetFocus
6230          Else
6240            .frmTaxLot_Sub.SetFocus
6250          End If
6260        End With
6270      End Select
6280    End If

EXITP:
6290    KeyCode = intRetVal
6300    Exit Sub

ERRH:
6310    intRetVal = 0
6320    THAT_PROC = THIS_PROC
6330    That_Erl = Erl
6340    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6350    Resume EXITP

End Sub

Private Sub shortname_GotFocus()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_GotFocus"

6410    With Me
6420      .shortname.SelLength = 0
6430      .shortname.SelStart = 99
6440    End With

EXITP:
6450    Exit Sub

ERRH:
6460    Select Case ERR.Number
        Case Else
6470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6480    End Select
6490    Resume EXITP

End Sub

Private Sub shortname_KeyDown(KeyCode As Integer, Shift As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "shortname_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
6570        With Me
6580          intRetVal = 0
6590          .opgChoice.SetFocus
6600        End With
6610      End Select
6620    End If

        ' ** Shift keys.
6630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6650        With Me
6660          intRetVal = 0
6670          .cmdCancel.SetFocus
6680        End With
6690      End Select
6700    End If

EXITP:
6710    KeyCode = intRetVal
6720    Exit Sub

ERRH:
6730    intRetVal = 0
6740    THAT_PROC = THIS_PROC
6750    That_Erl = Erl
6760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6770    Resume EXITP

End Sub

Private Sub opgChoice_Enter()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_Enter"

6810    With Me
6820      intLastChoice = .opgChoice
6830    End With

EXITP:
6840    Exit Sub

ERRH:
6850    Select Case ERR.Number
        Case Else
6860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6870    End Select
6880    Resume EXITP

End Sub

Private Sub opgChoice_AfterUpdate()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, frmSub As Access.Form
        Dim dblMax As Double, dblMin As Double
        Dim datAssetDate As Date
        Dim strMsg As String
        Dim msgResponse As VbMsgBoxResult
        Dim lngX As Long

6910    Set dbs = CurrentDb

6920    With Me

          ' ** Even though we're now using an array, let this
          ' ** sorting continue so they'll see what's happening.

6930      strAccountNo = .accountno
6940      lngAssetNo = .assetno

6950      .opgChoice_optFIFO_lbl.FontBold = False
6960      .opgChoice_optFIFO_box.Visible = False
6970      .opgChoice_optAverage_lbl.FontBold = False
6980      .opgChoice_optAverage_box.Visible = False
6990      .opgChoice_optSpecific_lbl.FontBold = False
7000      .opgChoice_optSpecific_box.Visible = False
7010      .opgChoice_optHigh_lbl.FontBold = False
7020      .opgChoice_optHigh_box.Visible = False
7030      .opgChoice_optLow_lbl.FontBold = False
7040      .opgChoice_optLow_box.Visible = False

7050      Set frmSub = .frmTaxLot_Sub.Form

7060      Select Case .opgChoice
          Case .opgChoice_optFIFO.OptionValue

7070        .opgChoice_optFIFO_lbl.FontBold = True
7080        .opgChoice_optFIFO_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
7090        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmTaxLot_Sub.
7100        frmSub.SortNow "assetdate_display_lbl_DblClick"  ' ** Form Procedure: frmTaxLot_Sub.

            ' ** ActiveAssets, grouped by accountno, assetno, assetdate, by specified [actno], [astno].
7110        Set qdf = dbs.QueryDefs("qryJournal_Sale_09")
7120        With qdf.Parameters
7130          ![actno] = strAccountNo
7140          ![astno] = lngAssetNo
7150        End With
7160        Set rst1 = qdf.OpenRecordset
7170        With rst1
7180          .MoveFirst
7190          dblCost = ((![Cost] / ![shareface]) * Me.ShareFaceSale) * -1
7200          dblPCash = ((![Cost] / ![shareface]) * Me.ShareFaceSale)
7210          datAssetDate = ![assetdate]  ' ** Earliest Tax Lot's assetdate.
7220          .Close
7230        End With
7240        Set rst1 = Nothing
7250        Set qdf = Nothing

            ' ** Find the appropriate record.
7260        Set rst1 = frmSub.RecordsetClone
7270        With rst1
7280          .MoveFirst
7290          For lngX = 1& To lngLotCnt
7300            If ![assetdate] = datAssetDate Then
7310              frmSub.Bookmark = .Bookmark
7320              Exit For
7330            End If
7340            If lngX < lngLotCnt Then .MoveNext
7350          Next
7360          .Close
7370        End With
7380        Set rst1 = Nothing

7390        frmSub.FocusHolder.SetFocus
7400        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmTaxLot_Sub.
7410        .opgChoice = .opgChoice_optFIFO.OptionValue

7420      Case .opgChoice_optAverage.OptionValue

7430        .opgChoice_optAverage_lbl.FontBold = True
7440        .opgChoice_optAverage_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
7450        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmTaxLot_Sub.
7460        frmSub.SortNow "assetdate_display_lbl_DblClick"  ' ** Form Procedure: frmTaxLot_Sub.

7470        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmTaxLot_Sub.
7480        If frmSub.IsAverage = False Then

7490          Select Case blnFromSaleBtn
              Case True
7500            strMsg = "If you choose to proceed, the price will be fixed at the" & vbCrLf & _
                  "average price of this asset until all the lots for this asset are sold." & vbCrLf & vbCrLf & _
                  "This holds even if the Tax Lot Holdings window is canceled." & vbCrLf & vbCrLf & _
                  "Do you want to proceed?"
7510          Case False
7520            strMsg = "If you choose to proceed, the price will be fixed at the" & vbCrLf & _
                  "average price of this asset until all the lots for this asset are sold." & vbCrLf & vbCrLf & _
                  "Do you want to proceed?"
7530          End Select
7540          msgResponse = MsgBox(strMsg, vbInformation + vbOKCancel, "Compute Average Price")

7550          Select Case msgResponse
              Case vbOK

                ' ** Update ActiveAssets, for IsAverage = True, by specified [actno], [astno].
7560            Set qdf = dbs.QueryDefs("qryJournal_Sale_10")
7570            With qdf.Parameters
7580              ![actno] = strAccountNo
7590              ![astno] = lngAssetNo
7600            End With
7610            qdf.Execute
7620            Set qdf = Nothing

                ' ** This will set get the average prices for all lots for NOW; uses table called 'tmpAveragePrice'.

7630            If TableExists("tmpAveragePrice") = False Then  ' ** Module Function: modFileUtilities.
                  ' ** Data-Definition: Create table tmpAveragePrice.
7640              Set qdf = dbs.QueryDefs("zz_qry_System_46_01")
7650              qdf.Execute
7660              Set qdf = Nothing
                  ' ** Data-Definition: Create index [assetno], [accountno] PrimaryKey on table tmpAveragePrice.
7670              Set qdf = dbs.QueryDefs("zz_qry_System_46_02")
7680              qdf.Execute
7690              Set qdf = Nothing
7700            End If

                ' ** Empty tmpAveragePrice.
7710            Set qdf = dbs.QueryDefs("qryLotInformation_03x")
7720            qdf.Execute
7730            Set qdf = Nothing

                ' ** Append qryLotInformation_02 (ActiveAssets, grouped and summed, by accountno, assetno,
                ' ** by specified [actno], [astno]) to tmpAveragePrice.
7740            Set qdf = dbs.QueryDefs("qryLotInformation_03")
7750            With qdf.Parameters
7760              ![actno] = strAccountNo
7770              ![astno] = lngAssetNo
7780            End With
7790            qdf.Execute
7800            Set qdf = Nothing

                ' ** This will apply those average prices to the ActiveAssets
                ' ** averagepriceperunit field.  This field will be used when the
                ' ** specified lots of an asset,acccount number pair have been averaged.

                ' ** NEW: 10/23/2014!
                ' ** Update qryLotInformation_04_05 (qryLotInformation_04_01 (ActiveAssets, linked to
                ' ** tmpAveragePrice, with uniqueid), with DLookups() to qryLotInformation_04_04
                ' ** (ActiveAssets, linked to qryLotInformation_04_03 (qryLotInformation_04_02
                ' ** (ActiveAssets, linked to tmpAveragePrice, grouped and summed), with costpershare),
                ' ** with averagepriceperunit_new, cost_new)), Update ActiveAssets, from tmpAveragePrice.
7810            Set qdf = dbs.QueryDefs("qryLotInformation_04")
7820            qdf.Execute
7830            Set qdf = Nothing

                ' ** ActiveAssets, grouped by accountno, assetno, assetdate, by specified [actno], [astno].
7840            Set qdf = dbs.QueryDefs("qryJournal_Sale_11")
7850            With qdf.Parameters
7860              ![actno] = strAccountNo
7870              ![astno] = lngAssetNo
7880            End With
7890            Set rst1 = qdf.OpenRecordset
7900            rst1.MoveLast
7910            rst1.MoveFirst

7920            With frmSub

7930              .FocusHolder.SetFocus
7940              DoEvents

7950              .averagepriceperunit.Visible = True
7960              .averagepriceperunit_lbl.Visible = True
7970              .averagepriceperunit_lbl_line.Visible = True
7980              .averagepriceperunit_avg.Visible = True
7990              .priceperunit.Visible = False
8000              .priceperunit_lbl.Visible = False
8010              .priceperunit_lbl_line.Visible = False
8020              .priceperunit_avg.Visible = False
8030              DoEvents

8040              .MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmTaxLot_Sub.

8050            End With

                ' ** Find the appropriate record.
                'APPROPRIATE RECORD? THEY ALL HAVE TO CHANGE!
                'For lngX = 1& To lngLotCnt
                '  rst1.FindFirst "[accountno] = '" & .accountno & "' And [assetno] = " & CStr(.assetno) & " And " & _
                '    "[assetdate] = #" & Format(frmSub.assetdate, "mm/dd/yyyy hh:nn:ss") & "#"
                '  If rst1.NoMatch = False Then
                '    'frmSub.priceperunit = rst1![priceperunit]
                '    frmSub.averagepriceperunit = rst1![averagepriceperunit]
                '    frmSub.Cost = rst1![Cost]
                '  End If
                '  DoEvents
                '  If lngX < lngLotCnt Then
                '    frmSub.MoveRec acCmdRecordsGoToNext  ' ** Form Procedure: frmTaxLot_Sub.
                '    frmSub.FocusHolder.SetFocus
                '    DoEvents
                '  End If
                'Next
                'rst1.Close
                'Set rst1 = Nothing
                'Set qdf = Nothing

8060            Set rst2 = frmSub.RecordsetClone
8070            rst2.MoveFirst
8080            For lngX = 1& To lngLotCnt
8090              rst1.FindFirst "[accountno] = '" & .accountno & "' And [assetno] = " & CStr(.assetno) & " And " & _
                    "[assetdate] = #" & Format(rst2![assetdate], "mm/dd/yyyy hh:nn:ss") & "#"
8100              If rst1.NoMatch = False Then
8110                rst2.Edit
8120                rst2![averagepriceperunit] = rst1![averagepriceperunit]
8130                rst2![Cost] = rst1![Cost]
8140                rst2.Update
8150              End If
8160              If lngX < lngLotCnt Then rst2.MoveNext
8170            Next
8180            rst2.Close
8190            Set rst2 = Nothing
8200            rst1.Close
8210            Set rst1 = Nothing
8220            Set qdf = Nothing

8230            frmSub.Refresh
8240            frmSub.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmTaxLot_Sub.
8250            frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmTaxLot_Sub.

8260            .opgChoice = .opgChoice_optAverage.OptionValue
8270            .opgChoice_optFIFO.Enabled = False
8280            .opgChoice_optAverage.Enabled = True
8290            .opgChoice_optSpecific.Enabled = False
8300            .opgChoice_optHigh.Enabled = False
8310            .opgChoice_optLow.Enabled = False

8320            frmSub.Refresh

8330            .opgChoice = .opgChoice_optAverage.OptionValue

8340          Case Else
8350            .opgChoice = intLastChoice
8360            .opgChoice_optAverage_lbl.FontBold = False
8370            Select Case .opgChoice
                Case .opgChoice_optFIFO.OptionValue
8380              .opgChoice_optFIFO_lbl.FontBold = True
8390            Case .opgChoice_optSpecific.OptionValue
8400              .opgChoice_optSpecific_lbl.FontBold = True
8410            Case .opgChoice_optLow.OptionValue
8420              .opgChoice_optLow_lbl.FontBold = True
8430            Case .opgChoice_optHigh.OptionValue
8440              .opgChoice_optHigh_lbl.FontBold = True
8450            End Select
8460          End Select
8470        End If

8480      Case .opgChoice_optSpecific.OptionValue

8490        .opgChoice_optSpecific_lbl.FontBold = True
8500        .opgChoice_optSpecific_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
8510        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmTaxLot_Sub.
8520        frmSub.SortNow "assetdate_display_lbl_DblClick"  ' ** Form Procedure: frmTaxLot_Sub.

8530        dblCost = ((frmSub.priceperunit * .ShareFaceSale) * -1)
8540        dblPCash = (frmSub.priceperunit * .ShareFaceSale)

8550        frmSub.FocusHolder.SetFocus
8560        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmTaxLot_Sub.
8570        .opgChoice = .opgChoice_optSpecific.OptionValue

8580      Case .opgChoice_optHigh.OptionValue

8590        .opgChoice_optHigh_lbl.FontBold = True
8600        .opgChoice_optHigh_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
8610        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmTaxLot_Sub.
8620        frmSub.SortNow "priceperunit_lbl_DblClick"  ' ** Form Procedure: frmTaxLot_Sub.

8630        Set rst1 = frmSub.RecordsetClone
8640        With rst1
8650          dblMax = 0
8660          .MoveFirst
8670          For lngX = 1& To lngLotCnt
8680            If dblMax < ![priceperunit] Then
8690              dblMax = ![priceperunit]
8700            End If
8710            If lngX < lngLotCnt Then .MoveNext
8720          Next
8730          .Close
8740        End With
8750        Set rst1 = Nothing

8760        dblCost = ((dblMax * .ShareFaceSale) * -1)
8770        dblPCash = (dblMax * .ShareFaceSale)

            ' ** Find the appropriate record.
8780        Set rst1 = frmSub.RecordsetClone
8790        With rst1
8800          .MoveFirst
8810          For lngX = 1& To lngLotCnt
8820            If Format(![priceperunit], "####.#####") = Format(dblMax, "####.#####") Then
8830              frmSub.Bookmark = .Bookmark
8840              Exit For
8850            End If
8860            If lngX < lngLotCnt Then .MoveNext
8870          Next
8880          .Close
8890        End With
8900        Set rst1 = Nothing

8910        frmSub.FocusHolder.SetFocus
8920        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmTaxLot_Sub.
8930        .opgChoice = .opgChoice_optHigh.OptionValue

8940      Case .opgChoice_optLow.OptionValue

8950        .opgChoice_optLow_lbl.FontBold = True
8960        .opgChoice_optLow_box.Visible = True

            ' ** Blank out the current sort so it doesn't do descending.
8970        frmSub.SortNow_Set vbNullString  ' ** Form Procedure: frmTaxLot_Sub.
8980        frmSub.SortNow "priceperunit_lbl_DblClick"  ' ** Form Procedure: frmTaxLot_Sub.
8990        DoEvents
            ' ** Then do it again to get descending.
9000        frmSub.SortNow "priceperunit_lbl_DblClick"  ' ** Form Procedure: frmTaxLot_Sub.

9010        Set rst1 = frmSub.RecordsetClone
9020        With rst1
9030          dblMin = 999999999
9040          .MoveFirst
9050          For lngX = 1& To lngLotCnt
9060            If dblMin > ![priceperunit] Then
9070              dblMin = ![priceperunit]
9080            End If
9090            If lngX < lngLotCnt Then .MoveNext
9100          Next
9110          .Close
9120        End With
9130        Set rst1 = Nothing

9140        dblCost = ((dblMin * .ShareFaceSale) * -1)
9150        dblPCash = (dblMin * .ShareFaceSale)

            ' ** Find the appropriate record.
9160        Set rst1 = frmSub.RecordsetClone
9170        With rst1
9180          .MoveFirst
9190          For lngX = 1& To lngLotCnt
9200            If Format(![priceperunit], "####.#####") = Format(dblMin, "####.#####") Then
9210              frmSub.Bookmark = .Bookmark
9220              Exit For
9230            End If
9240            If lngX < lngLotCnt Then .MoveNext
9250          Next
9260          .Close
9270        End With
9280        Set rst1 = Nothing

9290        frmSub.FocusHolder.SetFocus
9300        frmSub.MoveRec acCmdSelectRecord  ' ** Form Procedure: frmTaxLot_Sub.
9310        .opgChoice = .opgChoice_optLow.OptionValue

9320      End Select

9330      DoEvents
9340      RecArray_Load  ' ** Procedure: Below.

9350      intLastChoice = .opgChoice

9360    End With

9370    dbs.Close

EXITP:
9380    Set frmSub = Nothing
9390    Set rst1 = Nothing
9400    Set rst2 = Nothing
9410    Set qdf = Nothing
9420    Set dbs = Nothing
9430    Exit Sub

ERRH:
9440    Select Case ERR.Number
        Case Else
9450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9460    End Select
9470    Resume EXITP

End Sub

Private Sub opgChoice_optFIFO_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optFIFO_KeyDown"

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
9590          .frmTaxLot_Sub.SetFocus
9600          .frmTaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmTaxLot_Sub.
9610          .frmTaxLot_Sub.Form.FocusHolder.SetFocus
9620        End With
9630      End Select
9640    End If

        ' ** Shift keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9670        With Me
9680          intRetVal = 0
9690          .shortname.SetFocus
9700        End With
9710      End Select
9720    End If

EXITP:
9730    KeyCode = intRetVal
9740    Exit Sub

ERRH:
9750    intRetVal = 0
9760    THAT_PROC = THIS_PROC
9770    That_Erl = Erl
9780    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9790    Resume EXITP

End Sub

Private Sub opgChoice_optAverage_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optAverage_KeyDown"

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
9890          .frmTaxLot_Sub.SetFocus
9900          .frmTaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmTaxLot_Sub.
9910          .frmTaxLot_Sub.Form.FocusHolder.SetFocus
9920        End With
9930      End Select
9940    End If

        ' ** Shift keys.
9950    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9970        With Me
9980          intRetVal = 0
9990          .shortname.SetFocus
10000       End With
10010     End Select
10020   End If

EXITP:
10030   KeyCode = intRetVal
10040   Exit Sub

ERRH:
10050   intRetVal = 0
10060   THAT_PROC = THIS_PROC
10070   That_Erl = Erl
10080   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10090   Resume EXITP

End Sub

Private Sub opgChoice_optSpecific_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optSpecific_KeyDown"

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
10190         .frmTaxLot_Sub.SetFocus
10200         .frmTaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmTaxLot_Sub.
10210         .frmTaxLot_Sub.Form.FocusHolder.SetFocus
10220       End With
10230     End Select
10240   End If

        ' ** Shift keys.
10250   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10270       With Me
10280         intRetVal = 0
10290         .shortname.SetFocus
10300       End With
10310     End Select
10320   End If

EXITP:
10330   KeyCode = intRetVal
10340   Exit Sub

ERRH:
10350   intRetVal = 0
10360   THAT_PROC = THIS_PROC
10370   That_Erl = Erl
10380   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10390   Resume EXITP

End Sub

Private Sub opgChoice_optHigh_KeyDown(KeyCode As Integer, Shift As Integer)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optHigh_KeyDown"

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
10490         .frmTaxLot_Sub.SetFocus
10500         .frmTaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmTaxLot_Sub.
10510         .frmTaxLot_Sub.Form.FocusHolder.SetFocus
10520       End With
10530     End Select
10540   End If

        ' ** Shift keys.
10550   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10570       With Me
10580         intRetVal = 0
10590         .shortname.SetFocus
10600       End With
10610     End Select
10620   End If

EXITP:
10630   KeyCode = intRetVal
10640   Exit Sub

ERRH:
10650   intRetVal = 0
10660   THAT_PROC = THIS_PROC
10670   That_Erl = Erl
10680   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10690   Resume EXITP

End Sub

Private Sub opgChoice_optLow_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgChoice_optLow_KeyDown"

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
10790         .frmTaxLot_Sub.SetFocus
10800         .frmTaxLot_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmTaxLot_Sub.
10810         .frmTaxLot_Sub.Form.FocusHolder.SetFocus
10820       End With
10830     End Select
10840   End If

        ' ** Shift keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         intRetVal = 0
10890         .shortname.SetFocus
10900       End With
10910     End Select
10920   End If

EXITP:
10930   KeyCode = intRetVal
10940   Exit Sub

ERRH:
10950   intRetVal = 0
10960   THAT_PROC = THIS_PROC
10970   That_Erl = Erl
10980   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10990   Resume EXITP

End Sub

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

11010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
11020   Exit Sub

ERRH:
11030   Select Case ERR.Number
        Case Else
11040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11050   End Select
11060   Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

11110   With Me
11120     blnPrintReport_Focus = True
11130     .cmdPrintReport_raised_semifocus_dots_img.Visible = True
11140     .cmdPrintReport_raised_img.Visible = False
11150     .cmdPrintReport_raised_focus_img.Visible = False
11160     .cmdPrintReport_raised_focus_dots_img.Visible = False
11170     .cmdPrintReport_sunken_focus_dots_img.Visible = False
11180     .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

11310   With Me
11320     blnPrintReport_MouseDown = True
11330     .cmdPrintReport_sunken_focus_dots_img.Visible = True
11340     .cmdPrintReport_raised_img.Visible = False
11350     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
11360     .cmdPrintReport_raised_focus_img.Visible = False
11370     .cmdPrintReport_raised_focus_dots_img.Visible = False
11380     .cmdPrintReport_raised_img_dis.Visible = False
11390   End With

EXITP:
11400   Exit Sub

ERRH:
11410   Select Case ERR.Number
        Case Else
11420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11430   End Select
11440   Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

11510   With Me
          ' ** Uses qryLotInformation_07.
11520     strDocName = "rptListOfTaxLots"
11530     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
11540       .Modal = False
11550       DoCmd.OpenReport strDocName, acViewPreview, , , , THIS_NAME
11560       DoCmd.Maximize
11570       DoCmd.RunCommand acCmdFitToWindow
11580       .Visible = False
11590       Forms("frmJournal").Visible = False
11600     Else
            '##GTR_Ref: rptListOfTaxLots
11610       DoCmd.OpenReport strDocName, acViewNormal, , , , THIS_NAME
11620     End If
11630   End With

EXITP:
11640   Exit Sub

ERRH:
11650   Forms("frmJournal").Visible = True
11660   Me.Visible = True
11670   Select Case ERR.Number
        Case Else
11680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11690   End Select
11700   Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11820   intShiftDown = (Shift And acShiftMask) > 0
11830   intAltDown = (Shift And acAltMask) > 0
11840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11860     Select Case intRetVal
          Case vbKeyTab
11870       With Me
11880         intRetVal = 0
11890         If .chkIncludeCurrency.Enabled = True Then
11900           .chkIncludeCurrency.SetFocus
11910         ElseIf .cmdOK.Enabled = True Then
11920           .cmdOK.SetFocus
11930         Else
11940           .cmdCancel.SetFocus
11950         End If
11960       End With
11970     End Select
11980   End If

        ' ** Shift keys.
11990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12000     Select Case intRetVal
          Case vbKeyTab
12010       With Me
12020         intRetVal = 0
12030         .frmTaxLot_Sub.SetFocus
12040         .frmTaxLot_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmTaxLot_Sub.
12050         .frmTaxLot_Sub.Form.FocusHolder.SetFocus
12060       End With
12070     End Select
12080   End If

EXITP:
12090   KeyCode = intRetVal
12100   Exit Sub

ERRH:
12110   intRetVal = 0
12120   THAT_PROC = THIS_PROC
12130   That_Erl = Erl
12140   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12150   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

12210   With Me
12220     If blnPrintReport_MouseDown = False Then
12230       Select Case blnPrintReport_Focus
            Case True
12240         .cmdPrintReport_raised_focus_dots_img.Visible = True
12250         .cmdPrintReport_raised_focus_img.Visible = False
12260       Case False
12270         .cmdPrintReport_raised_focus_img.Visible = True
12280         .cmdPrintReport_raised_focus_dots_img.Visible = False
12290       End Select
12300       .cmdPrintReport_raised_img.Visible = False
12310       .cmdPrintReport_raised_semifocus_dots_img.Visible = False
12320       .cmdPrintReport_sunken_focus_dots_img.Visible = False
12330       .cmdPrintReport_raised_img_dis.Visible = False
12340     End If
12350   End With

EXITP:
12360   Exit Sub

ERRH:
12370   Select Case ERR.Number
        Case Else
12380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12390   End Select
12400   Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

12510   With Me
12520     .cmdPrintReport_raised_focus_dots_img.Visible = True
12530     .cmdPrintReport_raised_img.Visible = False
12540     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
12550     .cmdPrintReport_raised_focus_img.Visible = False
12560     .cmdPrintReport_sunken_focus_dots_img.Visible = False
12570     .cmdPrintReport_raised_img_dis.Visible = False
12580     blnPrintReport_MouseDown = False
12590   End With

EXITP:
12600   Exit Sub

ERRH:
12610   Select Case ERR.Number
        Case Else
12620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12630   End Select
12640   Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

12710   With Me
12720     .cmdPrintReport_raised_img.Visible = True
12730     .cmdPrintReport_raised_semifocus_dots_img.Visible = False
12740     .cmdPrintReport_raised_focus_img.Visible = False
12750     .cmdPrintReport_raised_focus_dots_img.Visible = False
12760     .cmdPrintReport_sunken_focus_dots_img.Visible = False
12770     .cmdPrintReport_raised_img_dis.Visible = False
12780     blnPrintReport_Focus = False
12790   End With

EXITP:
12800   Exit Sub

ERRH:
12810   Select Case ERR.Number
        Case Else
12820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12830   End Select
12840   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

12910   With Me

12920     Select Case .chkIncludeCurrency
          Case True
12930       .chkIncludeCurrency_lbl.FontBold = True
12940       .frmTaxLot_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmTaxLot_Sub.
12950       IncludeCurrency True  ' ** Procedure: Below.
12960     Case False
12970       .chkIncludeCurrency_lbl.FontBold = False
12980       .frmTaxLot_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmTaxLot_Sub.
12990       IncludeCurrency False  ' ** Procedure: Below.
13000     End Select

13010     If lngMonitorCnt > 1& And blnIsOpen = False Then
13020       EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
13030       lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
13040       If lngMonitorNum > 1& Then
13050         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
13060       End If
13070     End If

13080   End With

EXITP:
13090   Exit Sub

ERRH:
13100   Select Case ERR.Number
        Case Else
13110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13120   End Select
13130   Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

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
13290         If .cmdPrintReport.Enabled = True Then
13300           .cmdPrintReport.SetFocus
13310         ElseIf .cmdOK.Enabled = True Then
13320           .cmdOK.SetFocus
13330         Else
13340           .cmdCancel.SetFocus
13350         End If
13360       End With
13370     End Select
13380   End If

        ' ** Shift keys.
13390   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13400     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13410       With Me
13420         intRetVal = 0
13430         .frmTaxLot_Sub.SetFocus
13440       End With
13450     End Select
13460   End If

EXITP:
13470   KeyCode = intRetVal
13480   Exit Sub

ERRH:
13490   intRetVal = 0
13500   THAT_PROC = THIS_PROC
13510   That_Erl = Erl
13520   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13530   Resume EXITP

End Sub

Private Sub FocusHolder_KeyDown(KeyCode As Integer, Shift As Integer)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "FocusHolder_KeyDown"

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
13690         If .cmdOK.Enabled = True Then
13700           .cmdOK.SetFocus
13710         Else
13720           .cmdCancel.SetFocus
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
13810         .chkIncludeCurrency.SetFocus
13820       End With
13830     End Select
13840   End If

EXITP:
13850   KeyCode = intRetVal
13860   Exit Sub

ERRH:
13870   intRetVal = 0
13880   THAT_PROC = THIS_PROC
13890   That_Erl = Erl
13900   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13910   Resume EXITP

End Sub

Private Function DoMultiLots() As Boolean

14000 On Error GoTo ERRH

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

14010   blnRetVal = True

14020   Set frm = Forms("frmJournal").frmJournal_Sub4_Sold.Form

14030   With Me

14040     Set frmSub = .frmTaxLot_Sub.Form

14050     dblPerShare_First = 0#: dblPerShare_Loop = 0#

14060     lngNews = 0&
14070     ReDim arr_lngNew(0)

14080     lngLots = 0&
14090     ReDim arr_varLot(L_ELEMS, 0)
14100     intChoice = .opgChoice  ' ** Save this!

          'ARR_VARREC() SHOULD BE THE CURRENT SUBFORM LIST!
          ' ** Instead of resorting the form, just sort the array.
14110     lngTmp01 = 0&
14120     ReDim arr_varTmp01(R_ELEMS, 0)
14130     For lngX = 0& To (lngRecs - 1&)
14140       If arr_varRec(R_ZERO, lngX) = False Then
14150         lngTmp01 = lngTmp01 + 1&
14160         lngE = lngTmp01 - 1&
14170         ReDim Preserve arr_varTmp01(R_ELEMS, lngE)
14180         For lngZ = 0& To R_ELEMS
14190           arr_varTmp01(lngZ, lngE) = arr_varRec(lngZ, lngX)
14200         Next
14210       End If
14220     Next

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
          ' **     7       6     cost_usd        R_USD
          ' **     8       7     priceperunit    R_PPU
          ' **     9       8     curr_id         R_CURID
          ' **
          ' *************************************************

14230     Select Case .opgChoice
          Case .opgChoice_optFIFO.OptionValue
            ' ** Binary Sort arr_varTmp01() array by R_ADAT, oldest to newest.
14240       For lngX = UBound(arr_varTmp01, 2) To 1 Step -1
14250         For lngY = 0 To (lngX - 1)
14260           If arr_varTmp01(R_ADAT, lngY) > arr_varTmp01(R_ADAT, (lngY + 1)) Then
14270             For lngZ = 0& To R_ELEMS
14280               varTmp00 = arr_varTmp01(lngZ, lngY)
14290               arr_varTmp01(lngZ, lngY) = arr_varTmp01(lngZ, (lngY + 1&))
14300               arr_varTmp01(lngZ, (lngY + 1&)) = varTmp00
14310             Next
14320           End If
14330         Next
14340       Next
14350     Case .opgChoice_optAverage.OptionValue
            ' ** Changes already made in opgChoice_AfterUpdate().
            ' ** Binary Sort arr_varTmp01() array by R_ADAT, oldest to newest.
14360       For lngX = UBound(arr_varTmp01, 2) To 1 Step -1
14370         For lngY = 0 To (lngX - 1)
14380           If arr_varTmp01(R_ADAT, lngY) > arr_varTmp01(R_ADAT, (lngY + 1)) Then
14390             For lngZ = 0& To R_ELEMS
14400               varTmp00 = arr_varTmp01(lngZ, lngY)
14410               arr_varTmp01(lngZ, lngY) = arr_varTmp01(lngZ, (lngY + 1&))
14420               arr_varTmp01(lngZ, (lngY + 1&)) = varTmp00
14430             Next
14440           End If
14450         Next
14460       Next
14470     Case .opgChoice_optSpecific.OptionValue
            ' ** Handled above, so shouldn't ever get here.
14480       blnRetVal = False
14490     Case .opgChoice_optHigh.OptionValue
            'CONFIRM THIS SORTS CORRECTLY!
            ' ** Binary Sort arr_varTmp01() array by R_PPU, highest to lowest.
14500       For lngX = UBound(arr_varTmp01, 2) To 1 Step -1
14510         For lngY = 0 To (lngX - 1)
14520           If arr_varTmp01(R_PPU, lngY) < arr_varTmp01(R_PPU, (lngY + 1)) Then
14530             For lngZ = 0& To R_ELEMS
14540               varTmp00 = arr_varTmp01(lngZ, lngY)
14550               arr_varTmp01(lngZ, lngY) = arr_varTmp01(lngZ, (lngY + 1&))
14560               arr_varTmp01(lngZ, (lngY + 1&)) = varTmp00
14570             Next
14580           End If
14590         Next
14600       Next
14610     Case .opgChoice_optLow.OptionValue
            ' ** Binary Sort arr_varTmp01() array by R_PPU, lowest to highest.
14620       For lngX = UBound(arr_varTmp01, 2) To 1 Step -1
14630         For lngY = 0 To (lngX - 1)
14640           If arr_varTmp01(R_PPU, lngY) > arr_varTmp01(R_PPU, (lngY + 1)) Then
14650             For lngZ = 0& To R_ELEMS
14660               varTmp00 = arr_varTmp01(lngZ, lngY)
14670               arr_varTmp01(lngZ, lngY) = arr_varTmp01(lngZ, (lngY + 1&))
14680               arr_varTmp01(lngZ, (lngY + 1&)) = varTmp00
14690             Next
14700           End If
14710         Next
14720       Next
14730     End Select

14740     lngRecs = 0&
14750     ReDim arr_varRec(R_ELEMS, 0&)

14760     For lngX = 0& To (lngTmp01 - 1&)
14770       lngRecs = lngRecs + 1
14780       lngE = lngRecs - 1&
14790       ReDim Preserve arr_varRec(R_ELEMS, lngE)
14800       For lngZ = 0& To R_ELEMS
14810         arr_varRec(lngZ, lngE) = arr_varTmp01(lngZ, lngX)
14820       Next
14830     Next

          ' ** We'll use lngZ for navigating the arr_varRec() array.
14840     lngZ = 0&

14850     frm.saleCost = Format(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), strFormat)
          ' ** First item takes that item's cost, but recalculates it. Do we get a remainder problem here?

14860     gstrSaleCost = Format(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), strFormat)
14870     gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
14880     If blnHasForEx = True Then
            ' ** Since Rem_Dollar() won't remove foreign currency symbols, let's try this.
14890       gstrSaleCost = StringReplace(gstrSaleCost, strCurrSym, vbNullString)  ' ** Module Function: modStringFuncs.
14900     End If
14910     frm.UnitPurchaseDate = arr_varRec(R_ADAT, lngZ)
14920     frm.saleShareFace = arr_varRec(R_SHRS, lngZ)  ' ** 1st Tax Lot shareface put into existing, 1st Journal entry.
14930     varTmp00 = CDbl(Val(gstrSalePCash))
14940     varPCash = CCur(varTmp00)
14950     frm.salePCash = (varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
14960     varTmp00 = CDbl(Val(gstrSaleICash))
14970     varICash = CCur(varTmp00)
14980     frm.saleICash = (varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
14990     varRemaining = (.ShareFaceSale - arr_varRec(R_SHRS, lngZ))  ' ** Local .ShareFaceSale is total to sell.

          ' ** VGC 12/28/2009: If they've checked IsAverage, put it in the Journal too; just 'cause I'd like to see it!
15000     If .opgChoice = .opgChoice_optAverage.OptionValue Then
15010       frm.saleIsAverage = True
15020     End If

15030     varAccountNo = frm.saleAccountNo
15040     varSaletype = frm.saleType
15050     varTranDate = frm.saleTransDate

15060     If IsNull(frm.saleAssetDate) Then
15070       frm.saleAssetDate = Now
15080     ElseIf Not IsDate(frm.saleAssetDate) Then
15090       frm.saleAssetDate = Now
15100     End If

15110     varAssetDate = frm.saleAssetDate
15120     varAssetNo = frm.saleAssetno

          ' ** OK, the idea is to first let it find all the pieces needed to
          ' ** cover the sale. Then, when that's done, pull all the fractions
          ' ** out, add them up, and distribute that among the chosen pieces.
          ' ** A penny here, a penny there, etc.

15130     lngLots = lngLots + 1&  ' ** First record created in sale.
15140     lngE = lngLots - 1&
15150     ReDim Preserve arr_varLot(L_ELEMS, lngE)
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
          ' **     21     curr_id          L_CURID
          ' **     22     cost_usd         L_USD
          ' **     23     cost_usd_new     L_UNEW
          ' **
          ' *******************************************
          ' ** We're still on the first Tax Lot.
15160     arr_varLot(L_JID, lngE) = frm.saleID
15170     arr_varLot(L_LOTNO, lngE) = lngLots
15180     arr_varLot(L_ACTNO, lngE) = CStr(varAccountNo)
15190     arr_varLot(L_ASTNO, lngE) = CLng(varAssetNo)
15200     arr_varLot(L_SHARE, lngE) = CDbl(arr_varRec(R_SHRS, lngZ))
15210     arr_varLot(L_SCENT, lngE) = CCenti(arr_varRec(R_SHRS, lngZ), 4)
15220     arr_varLot(L_SNEW, lngE) = CDbl(0)
15230     arr_varLot(L_SADD, lngE) = CDbl(0)
15240     arr_varLot(L_ICASH, lngE) = CDbl(varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15250     arr_varLot(L_ICENT, lngE) = CCenti((varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale)), 2)  ' ** Function: Below.
15260     arr_varLot(L_INEW, lngE) = CDbl(0)
15270     arr_varLot(L_IADD, lngE) = CDbl(0)
15280     arr_varLot(L_PCASH, lngE) = CDbl(varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15290     arr_varLot(L_PCENT, lngE) = CCenti((varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale)), 2)  ' ** Function: Below.
15300     arr_varLot(L_PNEW, lngE) = CDbl(0)
15310     arr_varLot(L_PADD, lngE) = CDbl(0)
15320     arr_varLot(L_COST, lngE) = CDbl((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1)
15330     arr_varLot(L_CCENT, lngE) = CCenti(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), 2)
15340     arr_varLot(L_CNEW, lngE) = CDbl(0)
15350     arr_varLot(L_CADD, lngE) = CDbl(0)
15360     arr_varLot(L_REM, lngE) = CDbl(varRemaining)
15370     arr_varLot(L_CURID, lngE) = lngCurrID
15380     If lngCurrID = 150& Then  ' ** USD.
15390       arr_varLot(L_USD, lngE) = arr_varLot(L_COST, lngE)
15400     Else
15410       arr_varLot(L_USD, lngE) = Round((arr_varLot(L_COST, lngE) * dblCurrRate2), 2)
15420     End If
15430     arr_varLot(L_UNEW, lngE) = CDbl(0)

15440     dblPerShare_First = frm.pershare

15450   End With

15460   Set dbs = CurrentDb
15470   Set rst = dbs.OpenRecordset("journal", dbOpenDynaset, dbConsistent)

        ' ** This is the array loop.
15480   Do While CDbl(varRemaining) >= 0.0001  ' ** varRemaining is ShareFace; ignore floating point garbage.

15490     With Me

            ' ** The array is already in the proper order.
15500       lngZ = lngZ + 1&

15510       If blnRetVal = True Then

15520         With rst
15530           .AddNew
15540           ![transdate] = varTranDate
15550           ![journaltype] = varSaletype
15560           ![accountno] = varAccountNo
15570           ![assetno] = varAssetNo
15580           ![assetdate] = varAssetDate
15590           ![PurchaseDate] = arr_varRec(R_ADAT, lngZ)
15600           ![revcode_ID] = lngRevcodeID
15610           ![taxcode] = lngTaxcode
15620           If strDesc <> vbNullString Then
15630             ![description] = strDesc
15640           End If
15650           ![journal_USER] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
15660           ![curr_id] = lngCurrID
15670         End With

15680         If varRemaining > arr_varRec(R_SHRS, lngZ) Then

15690           rst![PCash] = (varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15700           rst![ICash] = (varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15710           rst![Cost] = Format(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), strFormat)
                ' ** Each one takes that item's cost, but recalculates it.
                ' ** If it's selling all of them, wouldn't it be simpler to just take that item's cost, period?
                ' ** Like this?
                ' ** rst![cost] = Format((.cost * -1), "Currency")

15720           gstrSaleCost = Format(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), strFormat)
15730           gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
15740           If blnHasForEx = True Then
                  ' ** Since Rem_Dollar() won't remove foreign currency symbols, let's try this.
15750             gstrSaleCost = StringReplace(gstrSaleCost, strCurrSym, vbNullString)  ' ** Module Function: modStringFuncs.
15760           End If
15770           rst![shareface] = arr_varRec(R_SHRS, lngZ)
15780           varRemaining = (varRemaining - arr_varRec(R_SHRS, lngZ))

15790           lngLots = lngLots + 1&
15800           lngE = lngLots - 1&
15810           ReDim Preserve arr_varLot(L_ELEMS, lngE)
15820           arr_varLot(L_JID, lngE) = CLng(0)
15830           arr_varLot(L_LOTNO, lngE) = lngLots
15840           arr_varLot(L_ACTNO, lngE) = CStr(varAccountNo)
15850           arr_varLot(L_ASTNO, lngE) = CLng(varAssetNo)
15860           arr_varLot(L_SHARE, lngE) = CDbl(arr_varRec(R_SHRS, lngZ))
15870           arr_varLot(L_SCENT, lngE) = CCenti(arr_varRec(R_SHRS, lngZ), 4)
15880           arr_varLot(L_SNEW, lngE) = CDbl(0)
15890           arr_varLot(L_SADD, lngE) = CDbl(0)
15900           arr_varLot(L_ICASH, lngE) = CDbl(varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15910           arr_varLot(L_ICENT, lngE) = CCenti((varICash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale)), 2)  ' ** Function: Below.
15920           arr_varLot(L_INEW, lngE) = CDbl(0)
15930           arr_varLot(L_IADD, lngE) = CDbl(0)
15940           arr_varLot(L_PCASH, lngE) = CDbl(varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale))
15950           arr_varLot(L_PCENT, lngE) = CCenti((varPCash * (arr_varRec(R_SHRS, lngZ) / .ShareFaceSale)), 2)  ' ** Function: Below.
15960           arr_varLot(L_PNEW, lngE) = CDbl(0)
15970           arr_varLot(L_PADD, lngE) = CDbl(0)
15980           arr_varLot(L_COST, lngE) = CDbl((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1)
15990           arr_varLot(L_CCENT, lngE) = CCenti(((arr_varRec(R_SHRS, lngZ) * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), 2)
16000           arr_varLot(L_CNEW, lngE) = CDbl(0)
16010           arr_varLot(L_CADD, lngE) = CDbl(0)
16020           arr_varLot(L_REM, lngE) = CDbl(varRemaining)
16030           arr_varLot(L_CURID, lngE) = lngCurrID
16040           If lngCurrID = 150& Then  ' ** USD.
16050             arr_varLot(L_USD, lngE) = arr_varLot(L_COST, lngE)
16060           Else
16070             arr_varLot(L_USD, lngE) = Round((arr_varLot(L_COST, lngE) * dblCurrRate2), 2)
16080           End If
16090           arr_varLot(L_UNEW, lngE) = CDbl(0)

16100         Else
                ' ** Last one.

16110           rst![PCash] = (varPCash * (varRemaining / .ShareFaceSale))
16120           rst![ICash] = (varICash * (varRemaining / .ShareFaceSale))
16130           rst![Cost] = Format(((varRemaining * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), strFormat)
16140           gstrSaleCost = Format(((varRemaining * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), strFormat)
16150           gstrSaleCost = Rem_Dollar(gstrSaleCost)  ' ** Module Function: modStringFuncs.
16160           If blnHasForEx = True Then
                  ' ** Since Rem_Dollar() won't remove foreign currency symbols, let's try this.
16170             gstrSaleCost = StringReplace(gstrSaleCost, strCurrSym, vbNullString)  ' ** Module Function: modStringFuncs.
16180           End If

16190           lngLots = lngLots + 1&
16200           lngE = lngLots - 1&
16210           ReDim Preserve arr_varLot(L_ELEMS, lngE)
16220           arr_varLot(L_JID, lngE) = CLng(0)
16230           arr_varLot(L_LOTNO, lngE) = lngLots
16240           arr_varLot(L_ACTNO, lngE) = CStr(varAccountNo)
16250           arr_varLot(L_ASTNO, lngE) = CLng(varAssetNo)
16260           arr_varLot(L_SHARE, lngE) = arr_varRec(R_SHRS, lngZ)
16270           arr_varLot(L_SCENT, lngE) = CCenti(arr_varRec(R_SHRS, lngZ), 4)
16280           arr_varLot(L_SADD, lngE) = CDbl(0)
16290           arr_varLot(L_ICASH, lngE) = CDbl(varICash * (varRemaining / .ShareFaceSale))
16300           arr_varLot(L_ICENT, lngE) = CCenti((varICash * (varRemaining / .ShareFaceSale)), 2)  ' ** Function: Below.
16310           arr_varLot(L_INEW, lngE) = CDbl(0)
16320           arr_varLot(L_IADD, lngE) = CDbl(0)
16330           arr_varLot(L_PCASH, lngE) = CDbl(varPCash * (varRemaining / .ShareFaceSale))
16340           arr_varLot(L_PCENT, lngE) = CCenti((varPCash * (varRemaining / .ShareFaceSale)), 2)  ' ** Function: Below.
16350           arr_varLot(L_PNEW, lngE) = CDbl(0)
16360           arr_varLot(L_PADD, lngE) = CDbl(0)
16370           arr_varLot(L_COST, lngE) = CDbl((varRemaining * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1)
16380           arr_varLot(L_CCENT, lngE) = CCenti(((varRemaining * (arr_varRec(R_COST, lngZ) / arr_varRec(R_SHRS, lngZ))) * -1), 2)
16390           arr_varLot(L_CNEW, lngE) = CDbl(0)
16400           arr_varLot(L_CADD, lngE) = CDbl(0)
16410           arr_varLot(L_SNEW, lngE) = CDbl(varRemaining)
16420           arr_varLot(L_REM, lngE) = CDbl(0)
16430           arr_varLot(L_CURID, lngE) = lngCurrID
16440           If lngCurrID = 150& Then  ' ** USD.
16450             arr_varLot(L_USD, lngE) = arr_varLot(L_COST, lngE)
16460           Else
16470             arr_varLot(L_USD, lngE) = Round((arr_varLot(L_COST, lngE) * dblCurrRate2), 2)
16480           End If
16490           arr_varLot(L_UNEW, lngE) = CDbl(0)
16500           rst![shareface] = varRemaining
16510           varRemaining = 0

16520         End If

              ' ** VGC 12/28/2009: If they've checked IsAverage, put it in the Journal too; just 'cause I'd like to see it!
16530         If .opgChoice = .opgChoice_optAverage.OptionValue Then
16540           rst![IsAverage] = True
16550         End If

16560         If IsLoaded("frmJournal") = True Then  ' ** Module Function: modFileUtilities.
16570           If Forms("frmJournal").frmJournal_Sub4_Sold.Visible = True Then
16580             If IsNull(frm.saleDescription) = False Then
16590               If frm.saleDescription <> vbNullString Then
16600                 rst![description] = frm.saleDescription
16610               End If
16620             End If
16630           End If
16640         End If
16650         rst.Update
16660         rst.Bookmark = rst.LastModified

16670         lngNews = lngNews + 1&
16680         ReDim Preserve arr_lngNew(lngNews - 1&)
16690         arr_lngNew(lngNews - 1&) = rst![ID]
16700         arr_varLot(L_JID, lngE) = rst![ID]

16710       Else
16720         Exit Do
16730       End If  ' ** blnRetVal.

16740     End With  ' ** Me.

16750   Loop  ' ** varRemaining.

16760   rst.Close

16770   If dblPerShare_First <> dblPerShare_Loop Then
16780     frm.pershare = dblPerShare_Loop
16790   End If

16800 On Error GoTo ERRH

16810   If blnRetVal = True Then
16820     blnRetVal = DoMultiLots_Round  ' ** Function: Below.
16830   End If  ' ** blnRetVal.

16840   If blnRetVal = True Then
16850     With dbs
16860       Set rst = .OpenRecordset("journal", dbOpenDynaset, dbConsistent)
16870       With rst
16880         For lngX = 0& To (lngLots - 1&)
16890           If arr_varLot(L_JID, lngX) = frm.saleID Then
16900             If arr_varLot(L_INEW, lngX) <> 0# Then
16910               frm.saleICash = CCur(arr_varLot(L_INEW, lngX) + arr_varLot(L_IADD, lngX))
16920             ElseIf Nz(frm.saleICash, 0@) <> 0@ Then
16930               frm.saleICash = 0@
16940             End If
16950             If arr_varLot(L_PNEW, lngX) <> 0# Then
16960               frm.salePCash = CCur(arr_varLot(L_PNEW, lngX) + arr_varLot(L_PADD, lngX))
16970             ElseIf Nz(frm.salePCash, 0@) <> 0@ Then
16980               frm.salePCash = 0@
16990             End If
17000             If arr_varLot(L_CNEW, lngX) <> 0# Then
17010               frm.saleCost = CCur(arr_varLot(L_CNEW, lngX) + arr_varLot(L_CADD, lngX))
17020             ElseIf Nz(![Cost], 0@) <> 0@ Then
17030               frm.saleCost = 0@
17040             End If
17050             If arr_varLot(L_SNEW, lngX) <> 0# Then
17060               frm.saleShareFace = Abs(arr_varLot(L_SNEW, lngX))
17070             ElseIf arr_varLot(L_SNEW, lngX) = 0# And Nz(![shareface], 0@) <> 0@ Then
17080               frm.saleShareFace = 0@
17090             End If
17100           Else
17110             .FindFirst "[ID] = " & CStr(arr_varLot(L_JID, lngX))
17120             If .NoMatch = False Then
17130               .Edit
17140               If arr_varLot(L_INEW, lngX) <> 0# Then
17150                 ![ICash] = CCur(arr_varLot(L_INEW, lngX) + arr_varLot(L_IADD, lngX))
17160               ElseIf Nz(![ICash], 0@) <> 0@ Then
17170                 ![ICash] = 0@
17180               End If
17190               If arr_varLot(L_PNEW, lngX) <> 0# Then
17200                 ![PCash] = CCur(arr_varLot(L_PNEW, lngX) + arr_varLot(L_PADD, lngX))
17210               ElseIf Nz(![PCash], 0@) <> 0@ Then
17220                 ![PCash] = 0@
17230               End If
17240               If arr_varLot(L_CNEW, lngX) <> 0# Then
17250                 ![Cost] = CCur(arr_varLot(L_CNEW, lngX) + arr_varLot(L_CADD, lngX))
17260               ElseIf Nz(![Cost], 0@) <> 0@ Then       ' ** Though this shouldn't be the case,
17270                 ![Cost] = 0@                          ' ** it'll certainly stick out if I've done something wrong!
17280               End If
17290               If arr_varLot(L_SNEW, lngX) <> 0# Then
17300                 ![shareface] = Abs(arr_varLot(L_SNEW, lngX))
17310               ElseIf arr_varLot(L_SNEW, lngX) = 0# And Nz(![shareface], 0@) <> 0@ Then
                      ' ** Though this shouldn't be the case, it'll certainly stick out if I've done something wrong!
17320                 ![shareface] = 0@
17330               End If
17340               .Update
17350             Else
17360               blnRetVal = False
17370               MsgBox "Journal entry not found.", vbInformation + vbOKOnly, "Multiple-Tax-Lot Error"
17380               Exit For
17390             End If
17400           End If
17410         Next
17420         .Close
17430       End With
17440       .Close
17450     End With
17460   End If  ' ** blnRetVal.

17470   DoCmd.SelectObject acForm, "frmJournal", False
17480   Forms("frmJournal").frmJournal_Sub4_Sold.SetFocus
17490   frm.Requery

17500   If lngLots > 1& Then
17510     If blnRetVal = True Then
17520       MsgBox "Sale of multiple tax lots." & Space(30), vbInformation + vbOKOnly, "Multiple Tax Lots"
17530     Else
17540       MsgBox "A problem occurred during multiple-tax-lot processing." & vbCrLf & _
              "Check the entries generated for accuracy.", vbExclamation + vbOKOnly, "Multiple-Tax-Lot Error"
17550     End If
17560   End If

17570 On Error Resume Next
17580   If frmSub.OrderByOn = True Then
17590     frmSub.OrderByOn = False
17600   End If
17610 On Error GoTo ERRH

EXITP:
17620   Set frm = Nothing
17630   Set frmSub = Nothing
17640   Set rst = Nothing
17650   Set qdf = Nothing
17660   Set dbs = Nothing
17670   DoMultiLots = blnRetVal
17680   Exit Function

ERRH_MULTI:
17690   blnRetVal = False
17700   If frmSub.OrderByOn = True Then
17710     frmSub.OrderByOn = False
17720   End If
17730   MsgBox "Unable to continue processing." & vbCrLf & "Please review the transactions already created.", _
          vbInformation + vbOKOnly, "Unable To Complete Operation"
17740   If lngNews > 0& Then
17750     For lngX = 0& To (lngNews - 1&)
            ' ** Delete Journal, by specidifed [jid].
17760       Set qdf = dbs.QueryDefs("qryLotInformation_05")
17770       With qdf.Parameters
17780         ![jid] = arr_lngNew(lngX)
17790       End With
17800       qdf.Execute
17810     Next
17820   End If
17830   rst.Close
17840   dbs.Close
17850   Me.cmdOK.Enabled = False
17860   DoCmd.SelectObject acForm, strCallingForm, False
17870   Forms(strCallingForm).frmJournal_Sub4_Sold.SetFocus
17880   frm.cmdSaleClose.Enabled = False
17890   frm.cmdSaleCancel.SetFocus
17900   Resume EXITP

ERRH:
17910   blnRetVal = False
17920   Select Case ERR.Number
        Case Else
17930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17940   End Select
17950   Resume EXITP

End Function

Private Function DoMultiLots_Round() As Boolean
' ** Now go through the array, rounding as appropriate.
' ** Remember, ShareFace CAN go to hundredths of a penny!

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "DoMultiLots_Round"

        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

18010   blnRetVal = True

        ' ** First, see if there are any tenths of a penny in the sale.
        ' ** The Centi elements hold hundredths of a penny, so 10 = $0.001.
18020   varShareface = CDbl(0): varICash = CDbl(0): varPCash = CDbl(0): varCost = CDbl(0)
18030   For lngX = 0& To (lngLots - 1&)
18040     varShareface = varShareface + arr_varLot(L_SCENT, lngX)
18050     varICash = varICash + arr_varLot(L_ICENT, lngX)
18060     varPCash = varPCash + arr_varLot(L_PCENT, lngX)
18070     varCost = varCost + arr_varLot(L_CCENT, lngX)
18080   Next

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
        ' **     21     curr_id          L_CURID
        ' **     22     cost_usd         L_USD
        ' **     23     cost_usd_new     L_UNEW
        ' **
        ' *******************************************

18090   If varICash <> 0 Or varPCash <> 0 Or varCost <> 0 Then
          ' ** Yes, there are tenths of a penny to deal with.

          ' ** Caclulate the totals of the raw numbers.
18100     dblShareTot = 0#: dblICashTot = 0#: dblPCashTot = 0#: dblCostTot = 0#
18110     For lngX = 0& To (lngLots - 1&)
18120       If lngX < (lngLots - 1&) Then
18130         dblShareTot = dblShareTot + arr_varLot(L_SHARE, lngX)
18140       Else
              ' ** The last record should have, in L_SNEW, the remaining portion being
              ' ** taken from L_SHARE, which is that lot's original ActiveAssets number.
              ' ** It's the only element in the array with a non-zero L_SNEW.
              ' ** (At least it should be there, not anywhere else, and not missing.)
18150         dblShareTot = dblShareTot + arr_varLot(L_SNEW, lngX)
              ' ** arr_varLot(L_SNEW, lngE) = CDbl(varRemaining)
              ' ** arr_varLot(L_REM, lngE) = CDbl(0)
              ' ** rst![shareface] = varRemaining
              ' ** varRemaining = 0
18160       End If
18170       dblICashTot = dblICashTot + arr_varLot(L_ICASH, lngX)
18180       dblPCashTot = dblPCashTot + arr_varLot(L_PCASH, lngX)
18190       dblCostTot = dblCostTot + arr_varLot(L_COST, lngX)
18200     Next

          ' ** Now round the totals.
          ' ** These will be the numbers I have to end up with.
18210     dblShareTot = Round(dblShareTot, 4)
18220     dblShareTot = (Val(CStr(dblShareTot * 10000)) / 10000)
18230     dblICashTot = Round(dblICashTot, 2)
18240     dblICashTot = (Val(CStr(dblICashTot * 100)) / 100)
18250     dblPCashTot = Round(dblPCashTot, 2)
18260     dblPCashTot = (Val(CStr(dblPCashTot * 100)) / 100)
18270     dblCostTot = Round(dblCostTot, 2)
18280     dblCostTot = (Val(CStr(dblCostTot * 100)) / 100)

          ' ** First, just lop off the tenths-on from the values, regardless of where they might round.
18290     dblShareNew = 0#: dblICashNew = 0#: dblPCashNew = 0#: dblCostNew = 0#
18300     For lngX = 0& To (lngLots - 1&)

18310       If arr_varLot(L_ICENT, lngX) <> 0 Then
18320         strTmp01 = CStr(arr_varLot(L_ICASH, lngX))
18330         intPos01 = InStr(strTmp01, ".")
18340         strTmp01 = Left(strTmp01, (intPos01 + 2))
18350         arr_varLot(L_INEW, lngX) = CDbl(strTmp01)
18360       Else
18370         arr_varLot(L_INEW, lngX) = arr_varLot(L_ICASH, lngX)
18380       End If

18390       If arr_varLot(L_PCENT, lngX) <> 0 Then
18400         strTmp01 = CStr(arr_varLot(L_PCASH, lngX))
18410         intPos01 = InStr(strTmp01, ".")
18420         strTmp01 = Left(strTmp01, (intPos01 + 2))
18430         arr_varLot(L_PNEW, lngX) = CDbl(strTmp01)
18440       Else
18450         arr_varLot(L_PNEW, lngX) = arr_varLot(L_PCASH, lngX)
18460       End If

18470       If arr_varLot(L_CCENT, lngX) <> 0 Then
18480         strTmp01 = CStr(arr_varLot(L_COST, lngX))
18490         intPos01 = InStr(strTmp01, ".")
18500         strTmp01 = Left(strTmp01, (intPos01 + 2))
18510         arr_varLot(L_CNEW, lngX) = CDbl(strTmp01)
18520         If arr_varLot(L_CURID, lngX) = 150& Then  ' ** USD.
18530           arr_varLot(L_UNEW, lngX) = arr_varLot(L_CNEW, lngX)
18540         Else
18550           arr_varLot(L_UNEW, lngX) = Round((arr_varLot(L_CNEW, lngX) * dblCurrRate2), 2)
18560         End If
18570       Else
18580         arr_varLot(L_CNEW, lngX) = arr_varLot(L_COST, lngX)
18590         If arr_varLot(L_CURID, lngX) = 150& Then  ' ** USD.
18600           arr_varLot(L_UNEW, lngX) = arr_varLot(L_CNEW, lngX)
18610         Else
18620           arr_varLot(L_UNEW, lngX) = Round((arr_varLot(L_CNEW, lngX) * dblCurrRate2), 2)
18630         End If
18640       End If

18650     Next

          ' ** Now figure out where to put the pennies!
          ' ** dblCostTot = The real final total, to which the pieces must add up.
          ' ** dblCostNew = The total of rounded pieces, which will be less than above.
          ' ** varCost    = The total remainders, as hundredths of a penny, so 10 = $0.001.

18660     For lngX = 0& To (lngLots - 1&)
18670       arr_varLot(L_INEW, lngX) = (Val(CStr(arr_varLot(L_INEW, lngX) * 100#)) / 100)
18680       arr_varLot(L_PNEW, lngX) = (Val(CStr(arr_varLot(L_PNEW, lngX) * 100#)) / 100)
18690       arr_varLot(L_CNEW, lngX) = (Val(CStr(arr_varLot(L_CNEW, lngX) * 100#)) / 100)
18700     Next

18710     For lngX = 0& To (lngLots - 1&)
18720       dblICashNew = dblICashNew + arr_varLot(L_INEW, lngX)
18730       dblPCashNew = dblPCashNew + arr_varLot(L_PNEW, lngX)
18740       dblCostNew = dblCostNew + arr_varLot(L_CNEW, lngX)
18750       If lngX < (lngLots - 1&) Then
18760         dblShareNew = dblShareNew + arr_varLot(L_SHARE, lngX)
18770       Else
18780         dblShareNew = dblShareNew + arr_varLot(L_SNEW, lngX)
18790       End If
18800     Next

18810     dblShareNew = (Val(CStr(dblShareNew * 10000)) / 10000)
18820     dblICashNew = (Val(CStr(dblICashNew * 100)) / 100)
18830     dblPCashNew = (Val(CStr(dblPCashNew * 100)) / 100)
18840     dblCostNew = (Val(CStr(dblCostNew * 100)) / 100)

18850     varRemaining = CDbl((Val(CStr(dblCostTot * 100#)) - Val(CStr(dblCostNew * 100#))) / 100)
18860     If varRemaining <> 0# Then
            ' ** varRemaining should be just a few cents.
            ' ** Do I want to make judgements as to which ones get the penny?
            ' ** That is, should those > $0.005 get them first?
            ' ** I could sort the array in descending order, thus giving precedence?
            ' ** Yah, go for it!
18870       DoMultiLots_Sort L_COST  ' ** Procedure: Below.
18880       Do While varRemaining <> 0&
18890         For lngX = (lngLots - 1&) To 0& Step -1&
                ' ** If Number Is  Sgn() Returns
                ' ** ============  ===================
                ' **     -1        Less than zero
                ' **      0        Equal to zero
                ' **      1        Greater than zero
18900           Select Case Sgn(varRemaining)
                Case -1
18910             arr_varLot(L_CADD, lngX) = arr_varLot(L_CADD, lngX) - 0.01
18920             varRemaining = varRemaining + 0.01
18930           Case Else
18940             arr_varLot(L_CADD, lngX) = arr_varLot(L_CADD, lngX) + 0.01
18950             varRemaining = varRemaining - 0.01
18960           End Select
18970           varRemaining = Round(varRemaining, 2)
                ' ** Lessen precision to remove E-Notation.
18980           varRemaining = Rem_ENotation(varRemaining)  ' ** Module Function: modStringFuncs.
18990           varRemaining = Val(CStr(varRemaining))
19000           If varRemaining = 0# Then
19010             Exit For
19020           End If
19030         Next
19040       Loop
19050     End If

19060     varRemaining = CDbl((Val(CStr(dblICashTot * 100#)) - Val(CStr(dblICashNew * 100#))) / 100)
19070     If varRemaining <> 0# Then
19080       DoMultiLots_Sort L_ICASH  ' ** Procedure: Below.
19090       Do While varRemaining <> 0&
19100         For lngX = (lngLots - 1&) To 0& Step -1&
19110           Select Case Sgn(varRemaining)
                Case -1
19120             arr_varLot(L_IADD, lngX) = arr_varLot(L_IADD, lngX) - 0.01
19130             varRemaining = varRemaining + 0.01
19140           Case Else
19150             arr_varLot(L_IADD, lngX) = arr_varLot(L_IADD, lngX) + 0.01
19160             varRemaining = varRemaining - 0.01
19170           End Select
19180           varRemaining = Round(varRemaining, 2)
                ' ** Lessen precision to remove E-Notation.
19190           varRemaining = Rem_ENotation(varRemaining)  ' ** Module Function: modStringFuncs.
19200           varRemaining = Val(CStr(varRemaining))
19210           If varRemaining = 0# Then
19220             Exit For
19230           End If
19240         Next
19250       Loop
19260     End If

19270     varRemaining = CDbl((Val(CStr(dblPCashTot * 100#)) - Val(CStr(dblPCashNew * 100#))) / 100)
19280     If varRemaining <> 0# Then
19290       DoMultiLots_Sort L_PCASH  ' ** Procedure: Below.
19300       Do While varRemaining <> 0#
19310         For lngX = (lngLots - 1&) To 0& Step -1&
19320           Select Case Sgn(varRemaining)
                Case -1
19330             arr_varLot(L_PADD, lngX) = arr_varLot(L_PADD, lngX) - 0.01
19340             varRemaining = varRemaining + 0.01
19350           Case Else
19360             arr_varLot(L_PADD, lngX) = arr_varLot(L_PADD, lngX) + 0.01
19370             varRemaining = varRemaining - 0.01
19380           End Select
19390           varRemaining = Round(varRemaining, 2)
                ' ** Lessen precision to remove E-Notation.
19400           varRemaining = Rem_ENotation(varRemaining)  ' ** Module Function: modStringFuncs.
19410           varRemaining = Val(CStr(varRemaining))
19420           If varRemaining = 0# Then
19430             Exit For
19440           End If
19450         Next
19460       Loop
19470     End If

          ' ** Put the array back into standard, lot number sort.
19480     DoMultiLots_Sort L_LOTNO  ' ** Procedure: Below.

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

19490     For lngX = 0& To (lngLots - 1&)
19500       If arr_varLot(L_SNEW, lngX) = 0# Then
19510         If lngX < (lngLots - 1&) Then
19520           arr_varLot(L_SNEW, lngX) = arr_varLot(L_SHARE, lngX)
19530         Else
19540           arr_varLot(L_SNEW, lngX) = arr_varLot(L_SHARE, lngX)  ' ** 07/06/2009: This is the line that was missing.
                ' ** I have no idea why it didn't show up before. Harumph!
                ' ** This added another ten-thousandth to the total ShareFace, making
                ' ** it 0.0001 off from the entered amount (836.1193 vs 836.1194).
                ' ** I guess subsequent transactions will just have to deal with that.
                ' ** Special handling for the final record in the multi-lot sale.
19550         End If
19560       End If
19570     Next

          ' ** OK. Now all NEW's have a rounded, unadjusted amount (except L_SNEW).
          ' ** All ADD's have the pennies to be added to NEW, as fully decimalized Doubles.
          ' ** So now those Journal records all have to be updated with these NEW + ADD's.

19580   Else

          ' ** If rounding isn't necessary, just copy the original values.
19590     For lngX = 0& To (lngLots - 1&)
19600       If arr_varLot(L_SNEW, lngX) = 0# Then
19610         arr_varLot(L_SNEW, lngX) = arr_varLot(L_SHARE, lngX)
19620       End If
19630       If arr_varLot(L_INEW, lngX) = 0@ Then
19640         arr_varLot(L_INEW, lngX) = arr_varLot(L_ICASH, lngX)
19650       End If
19660       If arr_varLot(L_PNEW, lngX) = 0@ Then
19670         arr_varLot(L_PNEW, lngX) = arr_varLot(L_PCASH, lngX)
19680       End If
19690       If arr_varLot(L_CNEW, lngX) = 0@ Then
19700         arr_varLot(L_CNEW, lngX) = arr_varLot(L_COST, lngX)
19710         If arr_varLot(L_CURID, lngX) = 150& Then  ' ** USD.
19720           arr_varLot(L_UNEW, lngX) = arr_varLot(L_CNEW, lngX)
19730         Else
19740           arr_varLot(L_UNEW, lngX) = Round((arr_varLot(L_CNEW, lngX) * dblCurrRate2), 2)
19750         End If
19760       End If
19770     Next

19780   End If

EXITP:
19790   DoMultiLots_Round = blnRetVal
19800   Exit Function

ERRH:
19810   blnRetVal = False
19820   Select Case ERR.Number
        Case Else
19830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19840   End Select
19850   Resume EXITP

End Function

Private Function CCenti(varInput As Variant, intPrec As Integer) As Double
' ** I wanted to use CMill(), but that's really only
' ** tenths of a penny, and I'm multiplying by 100.
' ** Darn!

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "CCenti"

        Dim strTmp01 As String, dblTmp02 As Double
        Dim intPos01 As Integer
        Dim dblRetVal As Double

19910   dblRetVal = 0#

19920   If IsNull(varInput) = False Then
19930     If IsNumeric(varInput) = True Then
19940       If InStr(CStr(varInput), "E") = 0 Then  ' ** Anything producing scientific notation is to small to worry about.
19950         dblTmp02 = varInput
19960         Select Case intPrec
              Case 2
19970           dblTmp02 = dblTmp02 * 100#
19980         Case 4
19990           dblTmp02 = dblTmp02 * 10000#
20000         End Select
20010         strTmp01 = CStr(dblTmp02)
20020         intPos01 = InStr(strTmp01, ".")
20030         If intPos01 > 0 Then
20040           strTmp01 = Mid(strTmp01, (intPos01 + 1))  ' ** Everything from the 3rd/5th decimal place and to the right.
20050           Select Case Len(strTmp01)
                Case 1
                  ' ** Limited to tenths of a penny.
20060             strTmp01 = strTmp01 & "0"
20070             dblRetVal = Val(strTmp01)
20080           Case 2
                  ' ** Hundredths of a penny.
20090             dblRetVal = Val(strTmp01)
20100           Case Else
                  ' ** Thousandths of a penny, etc.
20110             strTmp01 = Left(strTmp01, 2) & "." & Mid(strTmp01, 3)
20120             dblRetVal = Val(strTmp01)
                  ' ** 1. 72304.734288566  Original result
                  ' ** 2. 7230473.4288566  x 100
                  ' ** 3. 4288566
                  ' ** 4. 42.88566
20130           End Select
20140         Else
                ' ** Precision already no more than 2 decimal places.
20150         End If
20160       Else
              ' ** Too small. Throw it out!
20170       End If
20180     End If
20190   End If

EXITP:
20200   CCenti = dblRetVal
20210   Exit Function

ERRH:
20220   dblRetVal = 0#
20230   Select Case ERR.Number
        Case Else
20240     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20250   End Select
20260   Resume EXITP

End Function

Private Sub DoMultiLots_Sort(intField As Integer)

20300 On Error GoTo ERRH

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
20310   For lngX = UBound(arr_varLot, 2) To 1 Step -1
20320     For lngY = 0 To (lngX - 1)
20330       Select Case intField
            Case L_LOTNO
20340         dblSortA = arr_varLot(L_LOTNO, lngY)
20350         dblSortB = arr_varLot(L_LOTNO, (lngY + 1))
20360       Case L_SHARE
20370         dblSortA = arr_varLot(L_SHARE, lngY)
20380         dblSortB = arr_varLot(L_SHARE, (lngY + 1))
20390       Case L_ICASH
20400         dblSortA = arr_varLot(L_ICASH, lngY)
20410         dblSortB = arr_varLot(L_ICASH, (lngY + 1))
20420       Case L_PCASH
20430         dblSortA = arr_varLot(L_PCASH, lngY)
20440         dblSortB = arr_varLot(L_PCASH, (lngY + 1))
20450       Case L_COST
20460         dblSortA = arr_varLot(L_COST, lngY)
20470         dblSortB = arr_varLot(L_COST, (lngY + 1))
20480       End Select
20490       If dblSortA > dblSortB Then
20500         For lngZ = 0& To L_ELEMS
20510           varTmp00 = arr_varLot(lngZ, lngY)
20520           arr_varLot(lngZ, lngY) = arr_varLot(lngZ, (lngY + 1))
20530           arr_varLot(lngZ, (lngY + 1)) = varTmp00
20540         Next
20550       End If
20560     Next
20570   Next

EXITP:
20580   Exit Sub

ERRH:
20590   Select Case ERR.Number
        Case Else
20600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20610   End Select
20620   Resume EXITP

End Sub

Private Sub JrnlLocUpdate(blnIsMulti As Boolean, lngThisLocID As Long)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlLocUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
        Dim lngLocID As Long

20710   DoCmd.Hourglass True
20720   DoEvents

20730   Set dbs = CurrentDb
20740   With dbs
          ' ** qryLotInformation_29 (tmpEdit04, grouped by accountno, assetno, Location_ID),
          ' ** grouped by accountno, assetno, with cnt, Location_ID_min, Location_ID_max.
20750     Set qdf = .QueryDefs("qryLotInformation_30")
20760     Set rst = qdf.OpenRecordset
20770     With rst
20780       If .BOF = True And .EOF = True Then
              ' ** Shouldn't happen.
20790       Else
20800         .MoveFirst
20810         Select Case ![cnt]
              Case 1  ' ** All Tax Lots have the same Location_ID.
20820           Select Case blnIsMulti
                Case True
20830             lngLocID = ![Location_ID_max]  ' ** They'll both be the same.
20840           Case False
20850             If ![Location_ID_max] <> lngThisLocID And lngThisLocID > 0& Then
20860               lngLocID = lngThisLocID  ' ** Perhaps they changed it here?
20870             Else
20880               lngLocID = ![Location_ID_max]  ' ** They'll both be the same.
20890             End If
20900           End Select
20910         Case 2  ' ** A total of 2 different Location_ID's among all Tax Lots.
20920           Select Case blnIsMulti
                Case True
20930             If ![Location_ID_min] = 1& Then
20940               lngLocID = ![Location_ID_max]
20950             Else
20960               lngLocID = 1&  ' ** {Unassigned}, {no entry}.
20970             End If
20980           Case False
20990             If lngThisLocID = 1& Then
21000               If ![Location_ID_min] = 1& Then
21010                 lngLocID = ![Location_ID_max]
21020               Else
21030                 lngLocID = 1&  ' ** {Unassigned}, {no entry}.
21040               End If
21050             ElseIf lngThisLocID > 1& Then
21060               lngLocID = lngThisLocID
21070             Else
21080               lngLocID = 1&  ' ** {Unassigned}, {no entry}.
21090             End If
21100           End Select
21110         Case Else  ' ** Several different Location_ID's.
21120           Select Case blnIsMulti
                Case True
21130             lngLocID = 1&  ' ** {Unassigned}, {no entry}.
21140           Case False
21150             If lngThisLocID > 1& Then
21160               lngLocID = lngThisLocID
21170             Else
21180               lngLocID = 1&  ' ** {Unassigned}, {no entry}.
21190             End If
21200           End Select
21210         End Select
21220       End If
21230       .Close
21240     End With
21250     .Close
21260   End With

21270   Set frm = Forms("frmJournal").frmJournal_Sub4_Sold.Form
21280   With frm
21290     If .[Location_ID] <> lngLocID Then
21300       .[Location_ID] = lngLocID
21310     End If
21320   End With

EXITP:
21330   Set frm = Nothing
21340   Set rst = Nothing
21350   Set qdf = Nothing
21360   Set dbs = Nothing
21370   Exit Sub

ERRH:
21380   DoCmd.Hourglass False
21390   Select Case ERR.Number
        Case Else
21400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21410   End Select
21420   Resume EXITP

End Sub

Public Sub RecArray_Load()

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "RecArray_Load"

        Dim rst As DAO.Recordset
        Dim lngTmp01 As Long
        Dim lngX As Long, lngE As Long

21510   With Me

21520     lngRecs = 0&
21530     ReDim arr_varRec(R_ELEMS, 0)

          'THIS SHOULD ALREADY BE ADJUSTED FOR OTHER JOURNAL ENTRIES!
21540     Set rst = .frmTaxLot_Sub.Form.RecordsetClone
21550     With rst
21560       .MoveLast
21570       lngTmp01 = .RecordCount
21580       .MoveFirst
21590       strCurrSym = vbNullString
21600       For lngX = 1& To lngTmp01
21610         lngRecs = lngRecs + 1&
21620         lngE = lngRecs - 1&
21630         ReDim Preserve arr_varRec(R_ELEMS, lngE)
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
              ' **     7       6     cost_usd        R_USD
              ' **     8       7     priceperunit    R_PPU
              ' **     9       8     curr_id         R_CURID
              ' **
              ' *************************************************
21640         arr_varRec(R_TID, lngE) = ![tmpedit04_id]
21650         arr_varRec(R_ASTNO, lngE) = ![assetno]
21660         arr_varRec(R_SHRS, lngE) = ![shareface]
21670         arr_varRec(R_ADAT, lngE) = ![assetdate]
21680         arr_varRec(R_ZERO, lngE) = ![IsZero]
21690         arr_varRec(R_COST, lngE) = ![Cost]
21700         arr_varRec(R_USD, lngE) = ![cost_usd]
21710         Select Case ![IsAverage]
              Case True
21720           arr_varRec(R_PPU, lngE) = ![averagepriceperunit]
21730         Case False
21740           arr_varRec(R_PPU, lngE) = ![priceperunit]
21750         End Select
21760         arr_varRec(R_CURID, lngE) = ![curr_id]
21770         If ![curr_id] <> 150& Then  ' ** 150 = USD.
21780           blnHasForEx = True
21790         End If
21800         If strCurrSym = vbNullString Then
21810           strCurrSym = ![currsym_symbol]
21820           dblCurrRate2 = ![curr_rate2]
21830         End If
21840         If lngX < lngTmp01 Then .MoveNext
21850       Next
21860       .Close
21870     End With
21880     Set rst = Nothing
21890     strFormat = strCurrSym & "#,##0.00;(" & strCurrSym & "#,##0.00);" & strCurrSym & "0.00;" & strCurrSym & "0.00"
21900   End With

EXITP:
21910   Exit Sub

ERRH:
21920   Select Case ERR.Number
        Case Else
21930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21940   End Select
21950   Resume EXITP

End Sub

Public Function RecArray_Get() As Variant

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "RecArray_Get"

        Dim arr_varRetVal As Variant

22010   If lngRecs = 0& Then
22020     RecArray_Load  ' ** Procedure: Above.
22030   End If

22040   arr_varRetVal = arr_varRec

EXITP:
22050   RecArray_Get = arr_varRetVal
22060   Exit Function

ERRH:
22070   arr_varRetVal = Empty
22080   Select Case ERR.Number
        Case Else
22090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22100   End Select
22110   Resume EXITP

End Function

Private Sub IncludeCurrency(blnShow As Boolean)

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrency"

        Dim lngTmp01 As Long, lngTmp02 As Long

22210   With Me

22220     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
22230       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
22240     End If

22250     lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
22260     lngMonitorNum = 1&: lngTmp02 = 0&
22270     EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
22280     If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

22290     Select Case blnShow
          Case True
22300       DoCmd.SelectObject acForm, THIS_NAME, False
22310       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
22320       DoCmd.MoveSize lngFrm_Left, lngTmp02, lngFrm_Width, lngFrm_Height  'lngFrm_Top
22330       If lngMonitorNum > 1& Then
22340         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
22350       End If
22360       .Width = lngForm_Width
22370       .cmdCancel.Left = lngCancel_Left
22380       .cmdOK.Left = lngOK_Left
22390       .FocusHolder.Left = lngFocusHolder_Left
22400       .RecordCount_lbl.Left = lngRecCnt_Left
22410       .RecordCount_lbl_box.Left = lngRecCnt_Left
22420       .frmTaxLot_Sub.Width = lngSub_Width
22430       .frmTaxLot_Sub_box.Width = (.frmTaxLot_Sub.Width + (2& * lngTpp))
22440       .Header_vline01.Left = lngForm_Width
22450       .Header_vline02.Left = lngForm_Width
22460       .Detail_vline03.Left = lngForm_Width
22470       .Detail_vline04.Left = lngForm_Width
22480       .Footer_vline01.Left = lngForm_Width
22490       .Footer_vline02.Left = lngForm_Width
22500       .Header_hline01.Width = lngForm_Width
22510       .Header_hline02.Width = lngForm_Width
22520       .Detail_hline01.Width = lngDetailLine_Width
22530       .Detail_hline02.Width = lngDetailLine_Width
22540       .Detail_hline03.Width = lngForm_Width
22550       .Detail_hline04.Width = lngForm_Width
22560       .Footer_hline01.Width = lngForm_Width
22570       .Footer_hline02.Width = lngForm_Width
22580     Case False
22590       lngTmp01 = (lngCurrID_Width + lngCostUSD_Width)
22600       .RecordCount_lbl_box.Left = (lngRecCnt_Left - (lngTmp01 / 2))
22610       .RecordCount_lbl.Left = .RecordCount_lbl_box.Left
22620       .FocusHolder.Left = (lngFocusHolder_Left - lngTmp01)
22630       .cmdOK.Left = (lngOK_Left - lngTmp01)
22640       .cmdCancel.Left = (lngCancel_Left - lngTmp01)
22650       .frmTaxLot_Sub.Width = (lngSub_Width - lngTmp01)
22660       .frmTaxLot_Sub_box.Width = (.frmTaxLot_Sub.Width + (2& * lngTpp))
22670       .Header_vline01.Left = (lngForm_Width - lngTmp01)
22680       .Header_vline02.Left = .Header_vline01.Left
22690       .Detail_vline03.Left = (lngForm_Width - lngTmp01)
22700       .Detail_vline04.Left = .Header_vline01.Left
22710       .Footer_vline01.Left = .Header_vline01.Left
22720       .Footer_vline02.Left = .Header_vline01.Left
22730       .Header_hline01.Width = .Header_vline01.Left
22740       .Header_hline02.Width = .Header_vline01.Left
22750       .Detail_hline01.Width = (lngDetailLine_Width - lngTmp01)
22760       .Detail_hline02.Width = (lngDetailLine_Width - lngTmp01)
22770       .Detail_hline03.Width = .Header_vline01.Left
22780       .Detail_hline04.Width = .Header_vline01.Left
22790       .Footer_hline01.Width = .Header_vline01.Left
22800       .Footer_hline02.Width = .Header_vline01.Left
22810       .Width = (lngForm_Width - lngTmp01)
22820       DoCmd.SelectObject acForm, THIS_NAME, False
22830       If lngMonitorNum = 1& Then lngTmp02 = lngFrm_Top
22840       DoCmd.MoveSize (lngFrm_Left + (lngTmp01 / 2)), lngTmp02, (lngFrm_Width - lngTmp01), lngFrm_Height  'lngFrm_Top
22850       If lngMonitorNum > 1& Then
22860         LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
22870       End If
22880     End Select

22890   End With

EXITP:
22900   Exit Sub

ERRH:
22910   Select Case ERR.Number
        Case Else
22920     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22930   End Select
22940   Resume EXITP

End Sub
