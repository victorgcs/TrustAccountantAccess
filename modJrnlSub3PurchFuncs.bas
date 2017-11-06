Attribute VB_Name = "modJrnlSub3PurchFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modJrnlSub3PurchFuncs"

'VGC 09/07/2017: CHANGES!

' ** Combo box column constants: purchaseCurr_ID.
Private Const CBX_C_CURRID As Integer = 0  'curr_id
'Private Const CBX_C_CODE   As Integer = 1  'curr_code
'Private Const CBX_C_NAME   As Integer = 2  'curr_name
Private Const CBX_C_SYM    As Integer = 3  'currsym_symbol
Private Const CBX_C_DEC    As Integer = 4  'curr_decimal
'Private Const CBX_C_RATE1  As Integer = 5  'curr_rate1
Private Const CBX_C_RATE2  As Integer = 6  'curr_rate2
Private Const CBX_C_DATE   As Integer = 7  'curr_date
' **

Public Sub Calendar_Handler_Sub3(strProc As String, blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean, blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean, clsMonthClass As clsMonthCal, frm As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_Sub3"

        Dim strEvent As String, strCtlName As String
        Dim datStartDate As Date, datEndDate As Date
        Dim intNum As Integer
        Dim blnRetVal As Boolean

110     With frm

120       strEvent = Mid(strProc, (CharPos(strProc, 1, "_") + 1))  ' ** Module Function: modStringFuncs.
130       strCtlName = Left(strProc, (CharPos(strProc, 1, "_") - 1))  ' ** Module Function: modStringFuncs.
140       intNum = Val(Right(strCtlName, 1))

150       Select Case strEvent
          Case "Click"
160         Select Case intNum
            Case 1
170           datStartDate = Date
180           datEndDate = 0
190           blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
200           If blnRetVal = True Then
                ' ** Allow posting up to 1 month into the future.
210             If datStartDate > DateAdd("m", 1, Date) Then
220               MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
230               .purchaseTransDate = CDate(Format(Date, "mm/dd/yyyy"))
240             Else
250               .purchaseTransDate = datStartDate
260             End If
270           Else
280             .purchaseTransDate = CDate(Format(Date, "mm/dd/yyyy"))
290           End If
300           .purchaseAssetNo.SetFocus
310         Case 2
320           datStartDate = Date
330           datEndDate = 0
340           blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
350           If blnRetVal = True Then
360             If Compare_DateA_DateB(datStartDate, ">", Date) = True Then  ' ** Module Function: modStringFuncs.
370               MsgBox "Future trade dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
380               .assetdate = Now()
390               .purchaseAssetDate = CDate(Format(.assetdate, "mm/dd/yyyy"))
400             Else
410               .assetdate = datStartDate + time
420               .purchaseAssetDate = CDate(Format(.assetdate, "mm/dd/yyyy"))
430             End If
440           Else
450             .assetdate = Now()
460             .purchaseAssetDate = CDate(Format(.assetdate, "mm/dd/yyyy"))
470           End If
480           .cmbLocations.SetFocus
490         End Select
500       Case "GotFocus"
510         Select Case intNum
            Case 1
520           blnCalendar1_Focus = True
530           .cmdCalendar1_raised_semifocus_dots_img.Visible = True
540           .cmdCalendar1_raised_img.Visible = False
550           .cmdCalendar1_raised_focus_img.Visible = False
560           .cmdCalendar1_raised_focus_dots_img.Visible = False
570           .cmdCalendar1_sunken_focus_dots_img.Visible = False
580           .cmdCalendar1_raised_img_dis.Visible = False
590         Case 2
600           blnCalendar2_Focus = True
610           .cmdCalendar2_raised_semifocus_dots_img.Visible = True
620           .cmdCalendar2_raised_img.Visible = False
630           .cmdCalendar2_raised_focus_img.Visible = False
640           .cmdCalendar2_raised_focus_dots_img.Visible = False
650           .cmdCalendar2_sunken_focus_dots_img.Visible = False
660           .cmdCalendar2_raised_img_dis.Visible = False
670         End Select
680       Case "MouseDown"
690         Select Case intNum
            Case 1
700           blnCalendar1_MouseDown = True
710           .cmdCalendar1_sunken_focus_dots_img.Visible = True
720           .cmdCalendar1_raised_img.Visible = False
730           .cmdCalendar1_raised_semifocus_dots_img.Visible = False
740           .cmdCalendar1_raised_focus_img.Visible = False
750           .cmdCalendar1_raised_focus_dots_img.Visible = False
760           .cmdCalendar1_raised_img_dis.Visible = False
770         Case 2
780           blnCalendar2_MouseDown = True
790           .cmdCalendar2_sunken_focus_dots_img.Visible = True
800           .cmdCalendar2_raised_img.Visible = False
810           .cmdCalendar2_raised_semifocus_dots_img.Visible = False
820           .cmdCalendar2_raised_focus_img.Visible = False
830           .cmdCalendar2_raised_focus_dots_img.Visible = False
840           .cmdCalendar2_raised_img_dis.Visible = False
850         End Select
860       Case "MouseMove"
870         Select Case intNum
            Case 1
880           If blnCalendar1_MouseDown = False Then
890             Select Case blnCalendar1_Focus
                Case True
900               .cmdCalendar1_raised_focus_dots_img.Visible = True
910               .cmdCalendar1_raised_focus_img.Visible = False
920             Case False
930               .cmdCalendar1_raised_focus_img.Visible = True
940               .cmdCalendar1_raised_focus_dots_img.Visible = False
950             End Select
960             .cmdCalendar1_raised_img.Visible = False
970             .cmdCalendar1_raised_semifocus_dots_img.Visible = False
980             .cmdCalendar1_sunken_focus_dots_img.Visible = False
990             .cmdCalendar1_raised_img_dis.Visible = False
1000          End If
1010        Case 2
1020          If blnCalendar2_MouseDown = False Then
1030            Select Case blnCalendar2_Focus
                Case True
1040              .cmdCalendar2_raised_focus_dots_img.Visible = True
1050              .cmdCalendar2_raised_focus_img.Visible = False
1060            Case False
1070              .cmdCalendar2_raised_focus_img.Visible = True
1080              .cmdCalendar2_raised_focus_dots_img.Visible = False
1090            End Select
1100            .cmdCalendar2_raised_img.Visible = False
1110            .cmdCalendar2_raised_semifocus_dots_img.Visible = False
1120            .cmdCalendar2_sunken_focus_dots_img.Visible = False
1130            .cmdCalendar2_raised_img_dis.Visible = False
1140          End If
1150        End Select
1160      Case "MouseUp"
1170        Select Case intNum
            Case 1
1180          .cmdCalendar1_raised_focus_dots_img.Visible = True
1190          .cmdCalendar1_raised_img.Visible = False
1200          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
1210          .cmdCalendar1_raised_focus_img.Visible = False
1220          .cmdCalendar1_sunken_focus_dots_img.Visible = False
1230          .cmdCalendar1_raised_img_dis.Visible = False
1240          blnCalendar1_MouseDown = False
1250        Case 2
1260          .cmdCalendar2_raised_focus_dots_img.Visible = True
1270          .cmdCalendar2_raised_img.Visible = False
1280          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
1290          .cmdCalendar2_raised_focus_img.Visible = False
1300          .cmdCalendar2_sunken_focus_dots_img.Visible = False
1310          .cmdCalendar2_raised_img_dis.Visible = False
1320          blnCalendar2_MouseDown = False
1330        End Select
1340      Case "LostFocus"
1350        Select Case intNum
            Case 1
1360          .cmdCalendar1_raised_img.Visible = True
1370          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
1380          .cmdCalendar1_raised_focus_img.Visible = False
1390          .cmdCalendar1_raised_focus_dots_img.Visible = False
1400          .cmdCalendar1_sunken_focus_dots_img.Visible = False
1410          .cmdCalendar1_raised_img_dis.Visible = False
1420          blnCalendar1_Focus = False
1430          If IsNull(.purchaseType) = False Then
1440            If .purchaseType.Column(0) = "Liability" Or .purchaseType.Column(0) = "Liability (+)" Then
1450              If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03b" Then
1460                .purchaseAssetNo.RowSource = "qryJournal_Purchase_03b"  '#curr_id
1470              End If
1480            Else
1490              If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03c" Then
1500                .purchaseAssetNo.RowSource = "qryJournal_Purchase_03c"  '#curr_id
1510              End If
1520            End If
1530          End If
1540        Case 2
1550          .cmdCalendar2_raised_img.Visible = True
1560          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
1570          .cmdCalendar2_raised_focus_img.Visible = False
1580          .cmdCalendar2_raised_focus_dots_img.Visible = False
1590          .cmdCalendar2_sunken_focus_dots_img.Visible = False
1600          .cmdCalendar2_raised_img_dis.Visible = False
1610          blnCalendar2_Focus = False
1620        End Select
1630      End Select

1640    End With

EXITP:
1650    Exit Sub

ERRH:
1660    Select Case ERR.Number
        Case Else
1670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1680    End Select
1690    Resume EXITP

End Sub

Public Sub JournalType_After_Sub3(blnTypeIsNull As Boolean, blnGoToSaleForm As Boolean, blnDefTypeAssigned As Boolean, frm As Access.Form)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_After_Sub3"

1710    With frm

1720      If blnTypeIsNull = True Or .purchaseType <> .purchaseType.OldValue Or blnDefTypeAssigned = True Then

1730        blnDefTypeAssigned = False

1740        If .purchaseICash <> 0 Then
1750          .purchaseICash = 0
1760        End If
1770        If .purchasePCash <> 0 Then
1780          .purchasePCash = 0
1790        End If
1800        If .purchaseCost <> 0 Then
1810          .purchaseCost = 0
1820        End If
1830        Sub3Purch_Changed True, frm  ' ** Module Procedure: modPurchaseSold.

            'If .purchaseType = "Liability" Then
            '  ' ** qryJournal_Purchase_03b.
            '  strSQL = "SELECT qryJournal_Purchase_03a.assetno, qryJournal_Purchase_03a.totdesc, qryJournal_Purchase_03a.cusip, " & _
            '    "qryJournal_Purchase_03a.assettype, assettype.taxcode " & _
            '    "FROM qryJournal_Purchase_03a INNER JOIN assettype ON qryJournal_Purchase_03a.assettype = assettype.assettype " & _
            '    "WHERE (((qryJournal_Purchase_03a.assettype)='90') AND ((Left([qryJournal_Purchase_03a].[totdesc],3))<>'HA-')) " & _
            '    "ORDER BY qryJournal_Purchase_03a.assettype, qryJournal_Purchase_03a.totdesc;"
            'Else
            '  ' ** qryJournal_Purchase_03c.
            '  strSQL = "SELECT qryJournal_Purchase_03a.assetno, qryJournal_Purchase_03a.totdesc, qryJournal_Purchase_03a.cusip, " & _
            '    "qryJournal_Purchase_03a.assettype, assettype.taxcode " & _
            '    "FROM qryJournal_Purchase_03a INNER JOIN assettype ON qryJournal_Purchase_03a.assettype = assettype.assettype " & _
            '    "WHERE (((qryJournal_Purchase_03a.assettype)<>'90') AND ((Left([qryJournal_Purchase_03a].[totdesc],3))<>'HA-')) " & _
            '    "ORDER BY qryJournal_Purchase_03a.assettype, qryJournal_Purchase_03a.totdesc;"
            'End If
            '.purchaseAssetNo.RowSource = strSQL  ' ** See Sub3Purch_TaxCode Me(), below.

1840        If IsNull(.purchaseType) = False Then
1850          If .purchaseType.Column(0) = "Liability" Or .purchaseType.Column(0) = "Liability (+)" Then
1860            If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03b" Then
1870              .purchaseAssetNo.RowSource = "qryJournal_Purchase_03b"  '#curr_id
1880            End If
1890          Else
1900            If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03c" Then
1910              .purchaseAssetNo.RowSource = "qryJournal_Purchase_03c"  '#curr_id
1920            End If
1930          End If
1940        End If

1950        .purchaseAssetNo.Requery

1960        If .purchaseType = "Deposit" Then
1970          .purchaseICash.Enabled = False
1980          .purchaseICash.BorderColor = WIN_CLR_DISR
1990          .purchaseICash.BackStyle = acBackStyleTransparent
2000          .purchaseICash_lbl.BackStyle = acBackStyleTransparent
2010          .purchaseICash_lbl_box.Visible = True
2020          .purchasePCash.Enabled = False
2030          .purchasePCash.BorderColor = WIN_CLR_DISR
2040          .purchasePCash.BackStyle = acBackStyleTransparent
2050          .purchasePCash_lbl.BackStyle = acBackStyleTransparent
2060          .purchasePCash_lbl_box.Visible = True
2070        Else
2080          If .purchaseType = "Liability" Then
2090            .purchaseICash.Enabled = False
2100            .purchaseICash.BorderColor = WIN_CLR_DISR
2110            .purchaseICash.BackStyle = acBackStyleTransparent
2120            .purchaseICash_lbl.BackStyle = acBackStyleTransparent
2130            .purchaseICash_lbl_box.Visible = True
2140            .purchasePCash.Enabled = True
2150            .purchasePCash.BorderColor = CLR_LTBLU2
2160            .purchasePCash.BackStyle = acBackStyleNormal
2170            .purchasePCash_lbl.BackStyle = acBackStyleNormal
2180            .purchasePCash_lbl_box.Visible = False
2190          Else
2200            .purchaseICash.Enabled = True
2210            .purchaseICash.BorderColor = CLR_LTBLU2
2220            .purchaseICash.BackStyle = acBackStyleNormal
2230            .purchaseICash_lbl.BackStyle = acBackStyleNormal
2240            .purchaseICash_lbl_box.Visible = False
2250            If .journalSubtype = "Reinvest" Then
2260              .purchasePCash.Enabled = False
2270              .purchasePCash.BorderColor = WIN_CLR_DISR
2280              .purchasePCash.BackStyle = acBackStyleTransparent
2290              .purchasePCash_lbl.BackStyle = acBackStyleTransparent
2300              .purchasePCash_lbl_box.Visible = True
2310              .purchaseAssetDate.Enabled = True
2320            Else
2330              .purchasePCash.Enabled = True
2340              .purchasePCash.BorderColor = CLR_LTBLU2
2350              .purchasePCash.BackStyle = acBackStyleNormal
2360              .purchasePCash_lbl.BackStyle = acBackStyleNormal
2370              .purchasePCash_lbl_box.Visible = False
2380              .purchaseAssetDate.Enabled = True
2390            End If
2400          End If
2410        End If

2420        blnGoToSaleForm = False
2430        If .purchaseType = "Purchase" Then
2440          .tglPurchaseSale.Enabled = True
2450          .tglPurchaseSale_false_raised_img.Visible = True
2460          .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
2470          .tglPurchaseSale_false_raised_focus_img.Visible = False
2480          .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
2490          .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
2500          .tglPurchaseSale_false_raised_img_dis.Visible = False
2510          .tglPurchaseSale_true_raised_img.Visible = False
2520          .tglPurchaseSale_true_raised_focus_img.Visible = False
2530          .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
2540          .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
2550          .tglPurchaseSale_true_raised_img_dis.Visible = False
2560        End If

2570  On Error Resume Next
2580        DoCmd.RunCommand acCmdSaveRecord
2590  On Error GoTo ERRH

2600      End If

2610    End With

EXITP:
2620    Exit Sub

ERRH:
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Public Sub FormCurrent_Sub3(blnCmdButton As Boolean, blnGoToSaleForm As Boolean, blnDefTypeAssigned As Boolean, blnStartTrans As Boolean, blnTypeIsNull As Boolean, blnAccountNoErr As Boolean, blnAssetDateChecked As Boolean, datAssetDate_OldValue As Date, lngAssetDate_OldValue As Long, blnSpecialCap As Boolean, intSpecialCapOpt As Integer, lngNoChars As Long, arr_varNoChar As Variant, lngCurrID As Long, lngErrCnt As Long, frm As Access.Form)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormCurrent_Sub3"

        Dim strTmp01 As String, dblTmp02 As Double

2710    With frm

2720      DoCmd.SelectObject acForm, .Parent.Name, False

2730      blnDefTypeAssigned = False     ' ** Default.
2740      gblnIsLiability = False        ' ** Default.
2750      gstrPurchaseType = vbNullString
2760      blnStartTrans = False
2770      lngErrCnt = 0&

2780      .cmbAccountHelper = Null

2790      If .purchaseCurr_ID.Visible = True Then
2800        If .purchaseCurr_Date.Visible = True Then .purchaseCurr_Date.Visible = False
2810        .purchaseICash_usd = Null
2820        .purchaseICash_usd.Visible = False
2830        .purchaseICash.Format = "Currency"
2840        .purchaseICash.DecimalPlaces = 2
2850        .purchaseICash.BackColor = CLR_WHT
2860        .purchasePCash_usd = Null
2870        .purchasePCash_usd.Visible = False
2880        .purchasePCash.Format = "Currency"
2890        .purchasePCash.DecimalPlaces = 2
2900        .purchasePCash.BackColor = CLR_WHT
2910        .purchaseCost_usd = Null
2920        .purchaseCost_usd.Visible = False
2930        .purchaseCost.Format = "Currency"
2940        .purchaseCost.DecimalPlaces = 2
2950        .purchaseCost.BackColor = CLR_WHT
2960        If lngNoChars = 0& Or IsEmpty(arr_varNoChar) Then
2970          arr_varNoChar = .Parent.NoChar_Get  ' ** Form Function: frmJournal.
2980          lngNoChars = UBound(arr_varNoChar, 2) + 1&
2990        End If
3000      End If

3010      gstrFormQuerySpec = .Parent.Name  ' ** Make sure this is set for the assetno combo box.

3020      Select Case .NewRecord
          Case True
            ' ** I can't figure out what triggers creation of that new record after OK!
            ' ** I think it's the Nulling out after goto newrec!
3030        .cmbRevenueCodes.Undo
3040        .purchaseAccountNo.Undo
3050        .purchaseAccountNo_Data.Undo
3060        .purchaseAssetDate = Null  ' ** Display only.
3070        datAssetDate_OldValue = 0#
3080        lngAssetDate_OldValue = 0&
3090        gblnPurchaseChanged = False
3100        gblnPurchaseValidated = False
3110        blnAccountNoErr = False
3120      Case False
3130        blnAssetDateChecked = False
3140        Select Case IsNull(.assetdate)
            Case True
3150          datAssetDate_OldValue = 0
3160          lngAssetDate_OldValue = 0&
3170        Case False
3180          Select Case IsDate(.assetdate)
              Case True
3190            datAssetDate_OldValue = .assetdate
3200            dblTmp02 = CDbl(datAssetDate_OldValue)
3210            strTmp01 = CStr(dblTmp02)
3220            If InStr(strTmp01, ".") > 0 Then
3230              strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
3240            End If
3250            lngAssetDate_OldValue = CLng(strTmp01)
3260            .purchaseAssetDate = CDate(lngAssetDate_OldValue)  ' ** Display without timestamp.
3270          Case False
3280            datAssetDate_OldValue = 0
3290            lngAssetDate_OldValue = 0&
3300            .assetdate = Null
3310            .purchaseAssetDate = Null
3320          End Select
3330        End Select
3340      End Select

3350      Select Case IsNull(.purchaseAccountNo_Data)
          Case True
3360        .purchaseAccountNo = vbNullString
3370        .purchaseAccountNo.Enabled = True
3380        .purchaseAccountNo.BorderColor = CLR_LTBLU2
3390        .purchaseAccountNo.BackStyle = acBackStyleNormal
3400        .purchaseAccountNo_lbl.BackStyle = acBackStyleNormal
3410        .purchaseAccountNo_lbl_box.Visible = False
3420        .cmdLock.Enabled = False
3430        .cmdLock_open_raised_img_dis.Visible = True ' & strSysClrSfx).Visible = True
3440        .cmdLock_open_raised_img.Visible = False ' & strSysClrSfx).Visible = False
3450        .cmdLock_closed_raised_img.Visible = False ' & strSysClrSfx).Visible = False
3460        .cmbAccountHelper.Enabled = True
3470        .cmbAccountHelper.BorderColor = CLR_LTBLU2
3480        .cmbAccountHelper.BackStyle = acBackStyleNormal
3490        .purchaseAccountNo.SetFocus
3500        .cmdPurchaseMap.Enabled = True
3510      Case False
3520        .purchaseAccountNo = .purchaseAccountNo_Data
3530        gstrPurchaseAccountNumber = .purchaseAccountNo
3540        .purchaseTransDate.SetFocus  ' ** Make sure it's not on AccountNo before disabling.
3550        .purchaseAccountNo.Enabled = False
3560        .purchaseAccountNo.BorderColor = WIN_CLR_DISR
3570        .purchaseAccountNo.BackStyle = acBackStyleTransparent
3580        .purchaseAccountNo_lbl.BackStyle = acBackStyleTransparent
3590        .purchaseAccountNo_lbl_box.Visible = True
3600        .cmdLock.Enabled = True
3610        .cmdLock_open_raised_img_dis.Visible = False ' & strSysClrSfx).Visible = False
3620        .cmdLock_open_raised_img.Visible = False ' & strSysClrSfx).Visible = False
3630        .cmdLock_closed_raised_img.Visible = True ' & strSysClrSfx).Visible = True
3640        .cmbAccountHelper.Enabled = False
3650        .cmbAccountHelper.BorderColor = WIN_CLR_DISR
3660        .cmbAccountHelper.BackStyle = acBackStyleTransparent
3670        .cmdPurchaseMap.Enabled = False
3680      End Select

          ' ** Requery the locations to get the latest set.
3690      .cmbLocations.Requery

          ' ** Make sure validation is reset.
3700      gblnPurchaseValidated = False

3710      Select Case IsNull(.purchaseType)
          Case True
3720        blnTypeIsNull = True
3730      Case False
3740        blnTypeIsNull = False
3750      End Select

3760      .purchaseAssetNo.Requery

3770      If .purchaseType = "Deposit" Then
3780        .purchaseICash.Enabled = False
3790        .purchaseICash.BorderColor = WIN_CLR_DISR
3800        .purchaseICash.BackStyle = acBackStyleTransparent
3810        .purchaseICash_lbl.BackStyle = acBackStyleTransparent
3820        .purchaseICash_lbl_box.Visible = True
3830        .purchasePCash.Enabled = False
3840        .purchasePCash.BorderColor = WIN_CLR_DISR
3850        .purchasePCash.BackStyle = acBackStyleTransparent
3860        .purchasePCash_lbl.BackStyle = acBackStyleTransparent
3870        .purchasePCash_lbl_box.Visible = True
3880        .purchaseCost.Enabled = True
3890      Else
3900        If .purchaseType = "Liability" Then
3910          If .journalSubtype = "Reinvest" Then
3920            If .purchaseCost > 0 Then
3930              .purchasePCash = .purchaseCost
3940              .purchaseCost = .purchaseCost * -1
3950            Else
3960              .purchasePCash = .purchaseCost * -1
3970            End If
3980          End If
3990          .purchaseICash.Enabled = False
4000          .purchaseICash.BorderColor = WIN_CLR_DISR
4010          .purchaseICash.BackStyle = acBackStyleTransparent
4020          .purchaseICash_lbl.BackStyle = acBackStyleTransparent
4030          .purchaseICash_lbl_box.Visible = True
4040          .purchasePCash.Enabled = True
4050          .purchasePCash.BorderColor = CLR_LTBLU2
4060          .purchasePCash.BackStyle = acBackStyleNormal
4070          .purchasePCash_lbl.BackStyle = acBackStyleNormal
4080          .purchasePCash_lbl_box.Visible = False
4090          .purchaseCost.Enabled = True
4100          .purchaseICash = 0
4110        Else
4120          .purchaseICash.Enabled = True
4130          .purchaseICash.BorderColor = CLR_LTBLU2
4140          .purchaseICash.BackStyle = acBackStyleNormal
4150          .purchaseICash_lbl.BackStyle = acBackStyleNormal
4160          .purchaseICash_lbl_box.Visible = False
4170          .purchaseCost.Enabled = True
4180          If .journalSubtype = "Reinvest" Then
4190            .purchasePCash.Enabled = False
4200            .purchasePCash.BorderColor = WIN_CLR_DISR
4210            .purchasePCash.BackStyle = acBackStyleTransparent
4220            .purchasePCash_lbl.BackStyle = acBackStyleTransparent
4230            .purchasePCash_lbl_box.Visible = True
4240            .purchaseAssetDate.Enabled = True
4250          Else
4260            .purchasePCash.Enabled = True
4270            .purchasePCash.BorderColor = CLR_LTBLU2
4280            .purchasePCash.BackStyle = acBackStyleNormal
4290            .purchasePCash_lbl.BackStyle = acBackStyleNormal
4300            .purchasePCash_lbl_box.Visible = False
4310            .purchaseAssetDate.Enabled = True
4320          End If
4330        End If
4340      End If

4350      lngCurrID = .purchaseCurr_ID.Column(CBX_C_CURRID)

4360      gstrPurchaseType = IIf(IsNull(.purchaseType), vbNullString, .purchaseType)
4370      gstrPurchaseAsset = IIf(IsNull(.purchaseAssetNo), vbNullString, .purchaseAssetNo)
4380      gstrPurchaseShareFace = IIf(IsNull(.purchaseShareFace), vbNullString, .purchaseShareFace)
4390      gstrPurchaseAccountNumber = IIf(IsNull(.purchaseAccountNo), vbNullString, .purchaseAccountNo)
4400      gstrPurchaseICash = IIf(IsNull(.purchaseICash), vbNullString, .purchaseICash)
4410      gstrPurchaseICash = Rem_Dollar(gstrPurchaseICash, lngCurrID)  ' ** Module Function: modStringFuncs.
4420      gstrPurchasePCash = IIf(IsNull(.purchasePCash), vbNullString, .purchasePCash)
4430      gstrPurchasePCash = Rem_Dollar(gstrPurchasePCash, lngCurrID)  ' ** Module Function: modStringFuncs.
4440      gstrPurchaseCost = IIf(IsNull(.purchaseCost), vbNullString, .purchaseCost)
4450      gstrPurchaseCost = Rem_Dollar(gstrPurchaseCost, lngCurrID)  ' ** Module Function: modStringFuncs.

4460      Select Case .NewRecord
          Case True
4470        .cmdPurchaseCancel.Enabled = False
4480        .cmbLocations = Null
4490      Case False
4500        .cmdPurchaseCancel.Enabled = True
4510        .cmbLocations = .[Location_ID]
4520      End Select

4530      Select Case IsNull(.purchaseShareFace)
          Case True
4540        .purchaseShareFace.Format = "#,###"
4550      Case False
4560        If CLng(.purchaseShareFace) = CDbl(.purchaseShareFace) Then
4570          .purchaseShareFace.Format = "#,###"
4580        Else
4590          .purchaseShareFace.Format = "#,###.00"
4600        End If
4610      End Select

4620      If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
4630        .cmbTaxCodes.Enabled = True
4640        .cmbTaxCodes.BorderColor = CLR_LTBLU2
4650        .cmbTaxCodes.BackStyle = acBackStyleNormal
4660        .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
4670        .cmbTaxCodes_lbl_box.Visible = False
4680        .cmbRevenueCodes.Enabled = True
4690        .cmbRevenueCodes.BorderColor = CLR_LTBLU2
4700        .cmbRevenueCodes.BackStyle = acBackStyleNormal
4710        .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
4720        .cmbRevenueCodes_lbl_box.Visible = False
4730      ElseIf .purchaseICash <> 0 Then
4740        .cmbTaxCodes.Enabled = True
4750        .cmbTaxCodes.BorderColor = CLR_LTBLU2
4760        .cmbTaxCodes.BackStyle = acBackStyleNormal
4770        .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
4780        .cmbTaxCodes_lbl_box.Visible = False
4790        .cmbRevenueCodes.Enabled = True
4800        .cmbRevenueCodes.BorderColor = CLR_LTBLU2
4810        .cmbRevenueCodes.BackStyle = acBackStyleNormal
4820        .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
4830        .cmbRevenueCodes_lbl_box.Visible = False
4840      Else
4850        .cmbTaxCodes.Enabled = False
4860        .cmbTaxCodes.BorderColor = WIN_CLR_DISR
4870        .cmbTaxCodes.BackStyle = acBackStyleTransparent
4880        .cmbTaxCodes_lbl.BackStyle = acBackStyleTransparent
4890        .cmbTaxCodes_lbl_box.Visible = True
4900        .cmbRevenueCodes.Enabled = False
4910        .cmbRevenueCodes.BorderColor = WIN_CLR_DISR
4920        .cmbRevenueCodes.BackStyle = acBackStyleTransparent
4930        .cmbRevenueCodes_lbl.BackStyle = acBackStyleTransparent
4940        .cmbRevenueCodes_lbl_box.Visible = True
4950      End If

4960      DoEvents

4970      If IsNull(.purchaseType) = False And .NewRecord = False Then
4980        If .purchaseType.Column(0) = "Liability" Or .purchaseType.Column(0) = "Liability (+)" Then
4990          If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03b" Then
5000            .purchaseAssetNo.RowSource = "qryJournal_Purchase_03b"  '#curr_id
5010          End If
5020        Else
5030          If .purchaseAssetNo.RowSource <> "qryJournal_Purchase_03c" Then
5040            .purchaseAssetNo.RowSource = "qryJournal_Purchase_03c"  '#curr_id
5050          End If
5060        End If
5070      End If

5080      If blnCmdButton = False Then  ' ** That is, only run this if Form_Current() wasn't called from another button, like cmdPurchaseOK_Click().
5090        blnGoToSaleForm = False
5100        .tglPurchaseSale_false_raised_img_dis.Visible = True
5110        .tglPurchaseSale_false_raised_img.Visible = False
5120        .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
5130        .tglPurchaseSale_false_raised_focus_img.Visible = False
5140        .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
5150        .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
5160        .tglPurchaseSale_true_raised_img.Visible = False
5170        .tglPurchaseSale_true_raised_focus_img.Visible = False
5180        .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
5190        .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
5200        .tglPurchaseSale_true_raised_img_dis.Visible = False
5210        .tglPurchaseSale.Enabled = False
5220        If blnTypeIsNull = False Then
5230          If .purchaseType = "Paid" Then
5240            Select Case .posted
                Case True
5250              .tglPurchaseSale.Enabled = False
5260              .tglPurchaseSale_true_raised_img_dis.Visible = True
5270              .tglPurchaseSale_false_raised_img_dis.Visible = False
5280            Case False
5290              .tglPurchaseSale.Enabled = True
5300              .tglPurchaseSale_false_raised_img.Visible = True
5310              .tglPurchaseSale_false_raised_img_dis.Visible = False
5320            End Select
5330          End If
5340        End If

5350      End If  ' ** blnCmdButton.

5360      .Repaint

5370      DoEvents

          ' ** Set the currency symbol.
5380      .purchaseCurr_ID_AfterUpdate  ' ** Procedure: Below.

5390      If .purchaseCurr_ID.Visible = True Then
5400        .purchaseICash_AfterUpdate  ' ** Procedure: Below.
5410        .purchasePCash_AfterUpdate  ' ** Procedure: Below.
5420        .purchaseCost_AfterUpdate  ' ** Procedure: Below.
5430      End If

5440      Sub3Purch_TaxCode frm  ' ** Module Procedure: modPurchaseSold.
5450      Sub3Purch_RevCode frm  ' ** Module Procedure: modPurchaseSold.

5460    End With

EXITP:
5470    Exit Sub

ERRH:
5480    Select Case ERR.Number
        Case Else
5490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5500    End Select
5510    Resume EXITP

End Sub

Public Sub DetailMouse_Sub3(blnCalendar1_Focus As Boolean, blnCalendar2_Focus As Boolean, blnPurchaseSale_Focus As Boolean, frm As Access.Form)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "DetailMouse_Sub3"

5610    With frm
5620      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
5630        Select Case blnCalendar1_Focus
            Case True
5640          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
5650          .cmdCalendar1_raised_img.Visible = False
5660        Case False
5670          .cmdCalendar1_raised_img.Visible = True
5680          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5690        End Select
5700        .cmdCalendar1_raised_focus_img.Visible = False
5710        .cmdCalendar1_raised_focus_dots_img.Visible = False
5720        .cmdCalendar1_sunken_focus_dots_img.Visible = False
5730        .cmdCalendar1_raised_img_dis.Visible = False
5740      End If
5750      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
5760        Select Case blnCalendar2_Focus
            Case True
5770          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
5780          .cmdCalendar2_raised_img.Visible = False
5790        Case False
5800          .cmdCalendar2_raised_img.Visible = True
5810          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
5820        End Select
5830        .cmdCalendar2_raised_focus_img.Visible = False
5840        .cmdCalendar2_raised_focus_dots_img.Visible = False
5850        .cmdCalendar2_sunken_focus_dots_img.Visible = False
5860        .cmdCalendar2_raised_img_dis.Visible = False
5870      End If
5880      If .tglPurchaseSale_true_raised_focus_img.Visible = True Or .tglPurchaseSale_true_raised_focus_dots_img.Visible = True Or _
              .tglPurchaseSale_false_raised_focus_img.Visible = True Or .tglPurchaseSale_false_raised_focus_dots_img.Visible = True Then
5890        Select Case .posted
            Case True
5900          Select Case blnPurchaseSale_Focus
              Case True
5910            .tglPurchaseSale_true_raised_focus_dots_img.Visible = True  ' ** Same for ON.
5920            .tglPurchaseSale_true_raised_img.Visible = False
5930          Case False
5940            .tglPurchaseSale_true_raised_img.Visible = True
5950            .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
5960          End Select
5970          .tglPurchaseSale_false_raised_img.Visible = False
5980          .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
5990        Case False
6000          Select Case blnPurchaseSale_Focus
              Case True
6010            .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = True
6020            .tglPurchaseSale_false_raised_img.Visible = False
6030          Case False
6040            .tglPurchaseSale_false_raised_img.Visible = True
6050            .tglPurchaseSale_false_raised_semifocus_dots_img.Visible = False
6060          End Select
6070          .tglPurchaseSale_true_raised_img.Visible = False
6080          .tglPurchaseSale_true_raised_focus_dots_img.Visible = False
6090        End Select
6100        .tglPurchaseSale_false_raised_focus_img.Visible = False
6110        .tglPurchaseSale_false_raised_focus_dots_img.Visible = False
6120        .tglPurchaseSale_false_sunken_focus_dots_img.Visible = False
6130        .tglPurchaseSale_false_raised_img_dis.Visible = False
6140        .tglPurchaseSale_true_raised_focus_img.Visible = False
6150        .tglPurchaseSale_true_sunken_focus_dots_img.Visible = False
6160        .tglPurchaseSale_true_raised_img_dis.Visible = False
6170      End If
6180    End With

EXITP:
6190    Exit Sub

ERRH:
6200    Select Case ERR.Number
        Case Else
6210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6220    End Select
6230    Resume EXITP

End Sub
