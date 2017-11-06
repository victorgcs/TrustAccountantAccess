Attribute VB_Name = "modJrnlSub4SoldFuncs"
Option Compare Database
Option Explicit

'VGC 09/04/2017: CHANGES!

Private Const THIS_NAME As String = "modJrnlSub4SoldFuncs"
' **

Public Sub Calendar_Handler_Sub4(strProc As String, blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean, blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean, clsMonthClass As clsMonthCal, frm As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_Sub4"

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
230               .saleTransDate = CDate(Format(Date, "mm/dd/yyyy"))
240             Else
250               .saleTransDate = datStartDate
260             End If
270           Else
280             .saleTransDate = CDate(Format(Date, "mm/dd/yyyy"))
290           End If
300           .saleAssetno.SetFocus
310         Case 2
320           datStartDate = Date
330           datEndDate = 0
340           blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
350           If blnRetVal = True Then
360             If Compare_DateA_DateB(datStartDate, ">", Date) = True Then  ' ** Module Function: modStringFuncs.
370               MsgBox "Future trade dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
380               .saleAssetDate = Now()  ' ** The control is formatted mm/dd/yyyy.
390             Else
400               .saleAssetDate = datStartDate + time  ' ** The control is formatted mm/dd/yyyy.
410             End If
420           Else
430             .saleAssetDate = Now()  ' ** The control is formatted mm/dd/yyyy.
440           End If
450           .saleICash.SetFocus
460         End Select
470       Case "GotFocus"
480         Select Case intNum
            Case 1
490           blnCalendar1_Focus = True
500           .cmdCalendar1_raised_semifocus_dots_img.Visible = True
510           .cmdCalendar1_raised_img.Visible = False
520           .cmdCalendar1_raised_focus_img.Visible = False
530           .cmdCalendar1_raised_focus_dots_img.Visible = False
540           .cmdCalendar1_sunken_focus_dots_img.Visible = False
550           .cmdCalendar1_raised_img_dis.Visible = False
560         Case 2
570           blnCalendar2_Focus = True
580           .cmdCalendar2_raised_semifocus_dots_img.Visible = True
590           .cmdCalendar2_raised_img.Visible = False
600           .cmdCalendar2_raised_focus_img.Visible = False
610           .cmdCalendar2_raised_focus_dots_img.Visible = False
620           .cmdCalendar2_sunken_focus_dots_img.Visible = False
630           .cmdCalendar2_raised_img_dis.Visible = False
640         End Select
650       Case "MouseDown"
660         Select Case intNum
            Case 1
670           blnCalendar1_MouseDown = True
680           .cmdCalendar1_sunken_focus_dots_img.Visible = True
690           .cmdCalendar1_raised_img.Visible = False
700           .cmdCalendar1_raised_semifocus_dots_img.Visible = False
710           .cmdCalendar1_raised_focus_img.Visible = False
720           .cmdCalendar1_raised_focus_dots_img.Visible = False
730           .cmdCalendar1_raised_img_dis.Visible = False
740         Case 2
750           blnCalendar2_MouseDown = True
760           .cmdCalendar2_sunken_focus_dots_img.Visible = True
770           .cmdCalendar2_raised_img.Visible = False
780           .cmdCalendar2_raised_semifocus_dots_img.Visible = False
790           .cmdCalendar2_raised_focus_img.Visible = False
800           .cmdCalendar2_raised_focus_dots_img.Visible = False
810           .cmdCalendar2_raised_img_dis.Visible = False
820         End Select
830       Case "MouseMove"
840         Select Case intNum
            Case 1
850           If blnCalendar1_MouseDown = False Then
860             Select Case blnCalendar1_Focus
                Case True
870               .cmdCalendar1_raised_focus_dots_img.Visible = True
880               .cmdCalendar1_raised_focus_img.Visible = False
890             Case False
900               .cmdCalendar1_raised_focus_img.Visible = True
910               .cmdCalendar1_raised_focus_dots_img.Visible = False
920             End Select
930             .cmdCalendar1_raised_img.Visible = False
940             .cmdCalendar1_raised_semifocus_dots_img.Visible = False
950             .cmdCalendar1_sunken_focus_dots_img.Visible = False
960             .cmdCalendar1_raised_img_dis.Visible = False
970           End If
980         Case 2
990           If blnCalendar2_MouseDown = False Then
1000            Select Case blnCalendar2_Focus
                Case True
1010              .cmdCalendar2_raised_focus_dots_img.Visible = True
1020              .cmdCalendar2_raised_focus_img.Visible = False
1030            Case False
1040              .cmdCalendar2_raised_focus_img.Visible = True
1050              .cmdCalendar2_raised_focus_dots_img.Visible = False
1060            End Select
1070            .cmdCalendar2_raised_img.Visible = False
1080            .cmdCalendar2_raised_semifocus_dots_img.Visible = False
1090            .cmdCalendar2_sunken_focus_dots_img.Visible = False
1100            .cmdCalendar2_raised_img_dis.Visible = False
1110          End If
1120        End Select
1130      Case "MouseUp"
1140        Select Case intNum
            Case 1
1150          .cmdCalendar1_raised_focus_dots_img.Visible = True
1160          .cmdCalendar1_raised_img.Visible = False
1170          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
1180          .cmdCalendar1_raised_focus_img.Visible = False
1190          .cmdCalendar1_sunken_focus_dots_img.Visible = False
1200          .cmdCalendar1_raised_img_dis.Visible = False
1210          blnCalendar1_MouseDown = False
1220        Case 2
1230          .cmdCalendar2_raised_focus_dots_img.Visible = True
1240          .cmdCalendar2_raised_img.Visible = False
1250          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
1260          .cmdCalendar2_raised_focus_img.Visible = False
1270          .cmdCalendar2_sunken_focus_dots_img.Visible = False
1280          .cmdCalendar2_raised_img_dis.Visible = False
1290          blnCalendar2_MouseDown = False
1300        End Select
1310      Case "LostFocus"
1320        Select Case intNum
            Case 1
1330          .cmdCalendar1_raised_img.Visible = True
1340          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
1350          .cmdCalendar1_raised_focus_img.Visible = False
1360          .cmdCalendar1_raised_focus_dots_img.Visible = False
1370          .cmdCalendar1_sunken_focus_dots_img.Visible = False
1380          .cmdCalendar1_raised_img_dis.Visible = False
1390          blnCalendar1_Focus = False
1400        Case 2
1410          .cmdCalendar2_raised_img.Visible = True
1420          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
1430          .cmdCalendar2_raised_focus_img.Visible = False
1440          .cmdCalendar2_raised_focus_dots_img.Visible = False
1450          .cmdCalendar2_sunken_focus_dots_img.Visible = False
1460          .cmdCalendar2_raised_img_dis.Visible = False
1470          blnCalendar2_Focus = False
1480        End Select
1490      End Select

1500    End With

EXITP:
1510    Exit Sub

ERRH:
1520    Select Case ERR.Number
        Case Else
1530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1540    End Select
1550    Resume EXITP

End Sub

Public Sub PaidTotal_Sub4(blnFromElsewhere As Boolean, frm As Access.Form)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "PaidTotal_Sub4"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String
        Dim strJrnlUser As String
        Dim curPaidTot As Currency
        Dim varTmp00 As Variant
        Dim blnContinue As Boolean

1610    blnContinue = True
1620    curPaidTot = 0@

1630    With frm

1640      strJrnlUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

1650      If IsNull(.saleAccountNo_Data) = False And IsNull(.saleAccountNo) = False Then
1660        strAccountNo = .saleAccountNo_Data
1670      Else
1680        blnContinue = False
1690        MsgBox "Please enter a valid account number to continue.", vbInformation + vbOKOnly, "Entry Required"
1700      End If

1710      If blnContinue = True Then

1720        Set dbs = CurrentDb
1730        With dbs
1740          If strJrnlUser = "TAAdmin" Or strJrnlUser = "Superuser" Then
                ' ** Journal, grouped and summed, for 'Paid', all, by specified [actno].
1750            Set qdf = .QueryDefs("qryJournal_Sale_06b")
1760            With qdf.Parameters
1770              ![actno] = strAccountNo
1780            End With
1790          Else
                ' ** Journal, grouped and summed, for 'Paid', by specified [actno], [jusr].
1800            Set qdf = .QueryDefs("qryJournal_Sale_06a")
1810            With qdf.Parameters
1820              ![actno] = strAccountNo
1830              ![jusr] = strJrnlUser
1840            End With
1850          End If
1860          Set rst = qdf.OpenRecordset
1870          With rst
1880            If .BOF = True And .EOF = True Then
1890              curPaidTot = 0@
1900            Else
1910              .MoveFirst
1920              curPaidTot = CCur(Abs(Nz(![ICash], 0) + Nz(![PCash], 0)))
1930            End If
1940            .Close
1950          End With
1960          .Close
1970        End With

1980        If curPaidTot = 0@ Then
1990          blnContinue = False
2000          Select Case strJrnlUser
              Case "TAAdmin", "Superuser"
2010            MsgBox "There are no available 'Paid' transactions in the Journal for this account.", _
                  vbInformation + vbOKOnly, "Nothing To Do"
2020          Case Else
2030            MsgBox "There are no available 'Paid' transactions for User " & strJrnlUser & " in the Journal for this account.", _
                  vbInformation + vbOKOnly, "Nothing To Do"
2040          End Select
2050          .saleShareFace = 0#
2060          .saleICash = 0@
2070          .saleICash_usd = Null
2080          blnFromElsewhere = True
2090          .saleICash_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
2100          .salePCash = 0@
2110          .salePCash_usd = Null
2120          blnFromElsewhere = True
2130          .salePCash_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
2140          .saleCost = 0@
2150          .saleCost_usd = Null
2160          blnFromElsewhere = True
2170          .saleCost_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
2180        End If  ' ** curPaidTot.

2190      End If  ' ** blnContinue.

2200      If blnContinue = True Then
2210        .saleShareFace = CDbl(curPaidTot)
2220        .salePCash = curPaidTot
2230        .salePCash_usd = Null
2240        blnFromElsewhere = True
2250        .salePCash_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
2260        If IsNull(.saleCost) = False Then
2270          If .saleCost <> 0@ Then
2280            .saleCost = 0@
2290            .saleCost_usd = Null
2300            blnFromElsewhere = True
2310            .saleCost_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
2320          End If
2330        End If
2340        varTmp00 = DLookup("[taxlot]", "account", "[accountno] = '" & strAccountNo & "'")
2350        If IsNull(varTmp00) = False Then
2360          If Val(varTmp00) > 0 Then
2370            .Parent.DefAssetNo = CLng(Val(varTmp00))
2380            .saleAssetno = CLng(Val(varTmp00))
2390          Else
2400            .Parent.DefAssetNo = 0&
2410          End If
2420        Else
2430          .Parent.DefAssetNo = 0&
2440        End If
2450        .Parent.DefPaidTot = curPaidTot
2460        .saleShareFace.SetFocus
2470      End If  ' ** blnContinue.

2480    End With

EXITP:
2490    Set rst = Nothing
2500    Set qdf = Nothing
2510    Set dbs = Nothing
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Public Sub JournalType_After_Sub4(blnFromElsewhere As Boolean, blnDefTypeAssigned As Boolean, frm As Access.Form)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalType_After_Sub4"

        Dim blnContinue As Boolean

2610    blnContinue = True

2620    With frm

2630      Select Case IsNull(.saleType.OldValue)
          Case True
            ' ** OK.
2640      Case False
2650        Select Case blnDefTypeAssigned
            Case True
              ' ** OK.
2660        Case False
2670          If .saleType <> .saleType.OldValue Then
                ' ** OK, as previously set, but now may be prevented by SoldCostedSet().
2680          Else
2690            blnContinue = False
2700          End If
2710        End Select
2720      End Select

2730      If blnContinue = True Then

2740        blnDefTypeAssigned = False

2750        .SaleChanged True  ' ** Form Procedure: frmJournal_Sub4_Sold.

2760        .subSetAssetCombo  ' ** Form Procedure: frmJournal_Sub4_Sold.

2770        .saleCost.Locked = True
2780        If .saleType = "Withdrawn" Then
2790          .saleICash.Enabled = False
2800          .saleICash.BorderColor = WIN_CLR_DISR
2810          .saleICash.BackStyle = acBackStyleTransparent
2820          .saleICash_lbl.BackStyle = acBackStyleTransparent
2830          .saleICash_lbl_box.Visible = True
2840          .salePCash.Enabled = False
2850          .salePCash.BorderColor = WIN_CLR_DISR
2860          .salePCash.BackStyle = acBackStyleTransparent
2870          .salePCash_lbl.BackStyle = acBackStyleTransparent
2880          .salePCash_lbl_box.Visible = True
2890          .saleShareFace.Enabled = True
2900          .saleShareFace.BorderColor = CLR_LTBLU2
2910          .saleShareFace.BackStyle = acBackStyleNormal
2920          .saleShareFace_lbl.BackStyle = acBackStyleNormal
2930          .saleShareFace_lbl_box.Visible = False
2940        ElseIf .saleType = "Liability" Then
2950          .saleICash.Enabled = True  ' ** Per Rich, 07/15/08.
2960          .saleICash.BorderColor = CLR_LTBLU2
2970          .saleICash.BackStyle = acBackStyleNormal
2980          .saleICash_lbl.BackStyle = acBackStyleNormal
2990          .saleICash_lbl_box.Visible = False
3000          .salePCash.Enabled = True
3010          .salePCash.BorderColor = CLR_LTBLU2
3020          .salePCash.BackStyle = acBackStyleNormal
3030          .salePCash_lbl.BackStyle = acBackStyleNormal
3040          .salePCash_lbl_box.Visible = False
3050          .saleShareFace.Enabled = True
3060          .saleShareFace.BorderColor = CLR_LTBLU2
3070          .saleShareFace.BackStyle = acBackStyleNormal
3080          .saleShareFace_lbl.BackStyle = acBackStyleNormal
3090          .saleShareFace_lbl_box.Visible = False
3100        ElseIf .saleType = "Cost Adj." Then
3110          .saleShareFace.Enabled = False
3120          .saleShareFace.BorderColor = WIN_CLR_DISR
3130          .saleShareFace.BackStyle = acBackStyleTransparent
3140          .saleShareFace_lbl.BackStyle = acBackStyleTransparent
3150          .saleShareFace_lbl_box.Visible = True
3160          .saleICash.Enabled = False
3170          .saleICash.Locked = False
3180          .saleICash.BorderColor = WIN_CLR_DISR
3190          .saleICash.BackStyle = acBackStyleTransparent
3200          .saleICash_lbl.BackStyle = acBackStyleTransparent
3210          .saleICash_lbl_box.Visible = True
3220          .salePCash.Enabled = False
3230          .salePCash.Locked = False
3240          .salePCash.BorderColor = WIN_CLR_DISR
3250          .salePCash.BackStyle = acBackStyleTransparent
3260          .salePCash_lbl.BackStyle = acBackStyleTransparent
3270          .salePCash_lbl_box.Visible = True
3280          .saleAssetDate.Enabled = True
3290          .saleShareFace = 0
3300          .saleShareFace.Format = "#,###"
3310          .saleCost.Locked = False
3320        Else
3330          .saleICash.Enabled = True
3340          .saleICash.BorderColor = CLR_LTBLU2
3350          .saleICash.BackStyle = acBackStyleNormal
3360          .saleICash_lbl.BackStyle = acBackStyleNormal
3370          .saleICash_lbl_box.Visible = False
3380          .salePCash.Enabled = True
3390          .salePCash.BorderColor = CLR_LTBLU2
3400          .salePCash.BackStyle = acBackStyleNormal
3410          .salePCash_lbl.BackStyle = acBackStyleNormal
3420          .salePCash_lbl_box.Visible = False
3430          .saleAssetDate.Enabled = True
3440          .saleShareFace.Enabled = True
3450          .saleShareFace.BorderColor = CLR_LTBLU2
3460          .saleShareFace.BackStyle = acBackStyleNormal
3470          .saleShareFace_lbl.BackStyle = acBackStyleNormal
3480          .saleShareFace_lbl_box.Visible = False
3490        End If

3500        If .saleType = "Sold" Then
3510          If .posted = False And IsNull(.CheckNum) = True Then  ' ** To be sure.
3520            .tglSaleReinvest.Enabled = True
3530            .tglSaleReinvest_false_raised_img.Visible = True
3540            .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
3550            .tglSaleReinvest_false_raised_focus_img.Visible = False
3560            .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
3570            .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
3580            .tglSaleReinvest_false_raised_img_dis.Visible = False
3590            .tglSaleReinvest_true_raised_img.Visible = False
3600            .tglSaleReinvest_true_raised_focus_img.Visible = False
3610            .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
3620            .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
3630            .tglSaleReinvest_true_raised_img_dis.Visible = False
3640          End If
3650        End If

3660        If .saleICash <> 0 Then
3670          .saleICash = 0
3680          .saleICash_usd = Null
3690          blnFromElsewhere = True
3700          .saleICash_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
3710        End If
3720        If .salePCash <> 0 Then
3730          .salePCash = 0
3740          .salePCash_usd = Null
3750          blnFromElsewhere = True
3760          .salePCash_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
3770        End If
3780        If .saleCost <> 0 Then
3790          .saleCost = 0
3800          .saleCost_usd = Null
3810          blnFromElsewhere = True
3820          .saleCost_AfterUpdate  ' ** Form Procedure: frmJournal_Sub4_Sold.
3830        End If

3840        If Not IsNull(.saleType) Then
3850          DoCmd.SelectObject acForm, "frmJournal", False
3860          .Parent.frmJournal_Sub4_Sold.SetFocus
3870          DoCmd.RunCommand acCmdSaveRecord
3880        End If

3890      End If  ' ** blnContinue.

3900    End With

EXITP:
3910    Exit Sub

ERRH:
3920    Select Case ERR.Number
        Case Else
3930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3940    End Select
3950    Resume EXITP

End Sub

Public Sub FormCurrent_Sub4(intMode As Integer, blnAccountNoErr As Boolean, blnBeenToLotInfo As Boolean, blnClickedLotInfo As Boolean, blnFromSaleAssetnoEnter As Boolean, blnGoToSaleReinvest As Boolean, blnSpecialCap As Boolean, intSpecialCapOpt As Integer, frm As Access.Form)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormCurrent_Sub4"

4010    With frm
4020      Select Case intMode
          Case 1
4030        .saleICash_usd = Null
4040        .saleICash_usd.Visible = False
4050        .saleICash.Format = "Currency"
4060        .saleICash.DecimalPlaces = 2
4070        .saleICash.BackColor = CLR_WHT
4080        .salePCash_usd = Null
4090        .salePCash_usd.Visible = False
4100        .salePCash.Format = "Currency"
4110        .salePCash.DecimalPlaces = 2
4120        .salePCash.BackColor = CLR_WHT
4130        .saleCost_usd = Null
4140        .saleCost_usd.Visible = False
4150        .saleCost.Format = "Currency"
4160        .saleCost.DecimalPlaces = 2
4170        .saleCost.BackColor = CLR_WHT
4180      Case 2
4190        Select Case IsNull(.saleAccountNo_Data)
            Case True
4200          .saleAccountNo = vbNullString
4210          .saleAccountNo.Enabled = True
4220          .saleAccountNo.BorderColor = CLR_LTBLU2
4230          .saleAccountNo.BackStyle = acBackStyleNormal
4240          .saleAccountNo_lbl.BackStyle = acBackStyleNormal
4250          .saleAccountNo_lbl_box.Visible = False
4260          .cmdLock.Enabled = False
4270          .cmdLock_open_raised_img_dis.Visible = True
4280          .cmdLock_open_raised_img.Visible = False
4290          .cmdLock_closed_raised_img.Visible = False
4300          .cmbAccountHelper.Enabled = True
4310          .cmbAccountHelper.BorderColor = CLR_LTBLU2
4320          .cmbAccountHelper.BackStyle = acBackStyleNormal
4330          .saleAccountNo.SetFocus
4340          .saleShareFace.Locked = False
4350          .saleCost.Locked = False
4360          blnAccountNoErr = False
4370        Case False
4380          If blnBeenToLotInfo = False And blnClickedLotInfo = False Then
                ' ** This is where it comes when returning to a previously entered Sold.
4390            .saleAccountNo = .saleAccountNo_Data
4400            gstrSaleAccountNumber = .saleAccountNo
4410            .saleTransDate.SetFocus  ' ** Make sure it's not on AccountNo before disabling.
4420            .saleAccountNo.Enabled = False
4430            .saleAccountNo.BorderColor = WIN_CLR_DISR
4440            .saleAccountNo.BackStyle = acBackStyleTransparent
4450            .saleAccountNo_lbl.BackStyle = acBackStyleTransparent
4460            .saleAccountNo_lbl_box.Visible = True
4470            .cmdLock.Enabled = True
4480            .cmdLock_open_raised_img_dis.Visible = False
4490            .cmdLock_open_raised_img.Visible = False
4500            .cmdLock_closed_raised_img.Visible = True
4510            .cmbAccountHelper.Enabled = False
4520            .cmbAccountHelper.BorderColor = WIN_CLR_DISR
4530            .cmbAccountHelper.BackStyle = acBackStyleTransparent
4540            If blnFromSaleAssetnoEnter = False And .saleType <> "Cost Adj." Then  ' ** Because shareface is disabled, don't want it locked.
4550              .saleShareFace.Locked = True
4560              .saleCost.Locked = True
4570            End If
4580          End If
4590        End Select
4600      Case 3
4610        If .saleType = "Withdrawn" Then
4620          .saleICash.Enabled = False
4630          .saleICash.BorderColor = WIN_CLR_DISR
4640          .saleICash.BackStyle = acBackStyleTransparent
4650          .saleICash_lbl.BackStyle = acBackStyleTransparent
4660          .saleICash_lbl_box.Visible = True
4670          .salePCash.Enabled = False
4680          .salePCash.BorderColor = WIN_CLR_DISR
4690          .salePCash.BackStyle = acBackStyleTransparent
4700          .salePCash_lbl.BackStyle = acBackStyleTransparent
4710          .salePCash_lbl_box.Visible = True
4720          .saleShareFace.Enabled = True
4730          .saleShareFace.BorderColor = CLR_LTBLU2
4740          .saleShareFace.BackStyle = acBackStyleNormal
4750          .saleShareFace_lbl.BackStyle = acBackStyleNormal
4760          .saleShareFace_lbl_box.Visible = False
4770        ElseIf .saleType = "Liability" Then
4780          .saleICash.Enabled = True  ' ** Per Rich, 07/15/08.
4790          .saleICash.BorderColor = CLR_LTBLU2
4800          .saleICash.BackStyle = acBackStyleNormal
4810          .saleICash_lbl.BackStyle = acBackStyleNormal
4820          .saleICash_lbl_box.Visible = False
4830          .salePCash.Enabled = True
4840          .salePCash.BorderColor = CLR_LTBLU2
4850          .salePCash.BackStyle = acBackStyleNormal
4860          .salePCash_lbl.BackStyle = acBackStyleNormal
4870          .salePCash_lbl_box.Visible = False
4880          .saleShareFace.Enabled = True
4890          .saleShareFace.BorderColor = CLR_LTBLU2
4900          .saleShareFace.BackStyle = acBackStyleNormal
4910          .saleShareFace_lbl.BackStyle = acBackStyleNormal
4920          .saleShareFace_lbl_box.Visible = False
4930        ElseIf .saleType = "Cost Adj." Then
4940          .saleShareFace.Enabled = False
4950          .saleShareFace.BorderColor = WIN_CLR_DISR
4960          .saleShareFace.BackStyle = acBackStyleTransparent
4970          .saleShareFace_lbl.BackStyle = acBackStyleTransparent
4980          .saleShareFace_lbl_box.Visible = True
4990          .saleICash.Enabled = False
5000          .saleICash.BorderColor = WIN_CLR_DISR
5010          .saleICash.BackStyle = acBackStyleTransparent
5020          .saleICash_lbl.BackStyle = acBackStyleTransparent
5030          .saleICash_lbl_box.Visible = True
5040          .salePCash.Enabled = False
5050          .salePCash.BorderColor = WIN_CLR_DISR
5060          .salePCash.BackStyle = acBackStyleTransparent
5070          .salePCash_lbl.BackStyle = acBackStyleTransparent
5080          .salePCash_lbl_box.Visible = True
5090          .saleAssetDate.Enabled = True
5100          .saleShareFace = 0
5110          .saleShareFace.Format = "#,###"
5120          .saleCost.Locked = False
5130        Else
5140          .saleICash.Enabled = True
5150          .saleICash.BorderColor = CLR_LTBLU2
5160          .saleICash.BackStyle = acBackStyleNormal
5170          .saleICash_lbl.BackStyle = acBackStyleNormal
5180          .saleICash_lbl_box.Visible = False
5190          .salePCash.Enabled = True
5200          .salePCash.BorderColor = CLR_LTBLU2
5210          .salePCash.BackStyle = acBackStyleNormal
5220          .salePCash_lbl.BackStyle = acBackStyleNormal
5230          .salePCash_lbl_box.Visible = False
5240          .saleAssetDate.Enabled = True
5250          .saleShareFace.Enabled = True
5260          .saleShareFace.BorderColor = CLR_LTBLU2
5270          .saleShareFace.BackStyle = acBackStyleNormal
5280          .saleShareFace_lbl.BackStyle = acBackStyleNormal
5290          .saleShareFace_lbl_box.Visible = False
5300        End If
5310        If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
5320          .cmbTaxCodes.Enabled = True
5330          .cmbTaxCodes.BorderColor = CLR_LTBLU2
5340          .cmbTaxCodes.BackStyle = acBackStyleNormal
5350          .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
5360          .cmbTaxCodes_lbl_box.Visible = False
5370          .cmbRevenueCodes.Enabled = True
5380          .cmbRevenueCodes.BorderColor = CLR_LTBLU2
5390          .cmbRevenueCodes.BackStyle = acBackStyleNormal
5400          .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
5410          .cmbRevenueCodes_lbl_box.Visible = False
5420        ElseIf .saleICash <> 0 Then
5430          .cmbTaxCodes.Enabled = True
5440          .cmbTaxCodes.BorderColor = CLR_LTBLU2
5450          .cmbTaxCodes.BackStyle = acBackStyleNormal
5460          .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
5470          .cmbTaxCodes_lbl_box.Visible = False
5480          .cmbRevenueCodes.Enabled = True
5490          .cmbRevenueCodes.BorderColor = CLR_LTBLU2
5500          .cmbRevenueCodes.BackStyle = acBackStyleNormal
5510          .cmbRevenueCodes_lbl.BackStyle = acBackStyleNormal
5520          .cmbRevenueCodes_lbl_box.Visible = False
5530        Else
5540          .cmbRevenueCodes.Enabled = False
5550          .cmbRevenueCodes.BorderColor = WIN_CLR_DISR
5560          .cmbRevenueCodes.BackStyle = acBackStyleTransparent
5570          .cmbRevenueCodes_lbl.BackStyle = acBackStyleTransparent
5580          .cmbRevenueCodes_lbl_box.Visible = True
5590          If .saleType = "Withdrawn" Then
5600            .cmbTaxCodes.Enabled = True
5610            .cmbTaxCodes.BorderColor = CLR_LTBLU2
5620            .cmbTaxCodes.BackStyle = acBackStyleNormal
5630            .cmbTaxCodes_lbl.BackStyle = acBackStyleNormal
5640            .cmbTaxCodes_lbl_box.Visible = False
5650          Else
5660            .cmbTaxCodes.Enabled = False
5670            .cmbTaxCodes.BorderColor = WIN_CLR_DISR
5680            .cmbTaxCodes.BackStyle = acBackStyleTransparent
5690            .cmbTaxCodes_lbl.BackStyle = acBackStyleTransparent
5700            .cmbTaxCodes_lbl_box.Visible = True
5710          End If
5720        End If
5730      Case 4
5740        blnGoToSaleReinvest = False
5750        .tglSaleReinvest_false_raised_img_dis.Visible = True
5760        .tglSaleReinvest_false_raised_img.Visible = False
5770        .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
5780        .tglSaleReinvest_false_raised_focus_img.Visible = False
5790        .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
5800        .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
5810        .tglSaleReinvest_true_raised_img.Visible = False
5820        .tglSaleReinvest_true_raised_focus_img.Visible = False
5830        .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
5840        .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
5850        .tglSaleReinvest_true_raised_img_dis.Visible = False
5860        .tglSaleReinvest.Enabled = False
5870        Select Case .posted
            Case True
5880          .tglSaleReinvest.Enabled = False
5890          .tglSaleReinvest_true_raised_img_dis.Visible = True
5900          .tglSaleReinvest_false_raised_img_dis.Visible = False
5910        Case False
              ' ** Let's wait before enabling this.
              ' ** A Sold via the Sale button on Purchase shouldn't offer to reinvest.
5920          If .NewRecord = False And IsNull(.CheckNum) = True Then
5930            .tglSaleReinvest.Enabled = True
5940            .tglSaleReinvest_false_raised_img.Visible = True
5950            .tglSaleReinvest_false_raised_img_dis.Visible = False
5960          End If
5970        End Select
5980      End Select
5990    End With

EXITP:
6000    Exit Sub

ERRH:
6010    Select Case ERR.Number
        Case Else
6020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6030    End Select
6040    Resume EXITP

End Sub

Public Sub DetailMouse_Sub4(blnCalendar1_Focus As Boolean, blnCalendar2_Focus As Boolean, blnSaleReinvest_Focus As Boolean, frm As Access.Form)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "DetailMouse_Sub4"

6110    With frm
6120      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
6130        Select Case blnCalendar1_Focus
            Case True
6140          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
6150          .cmdCalendar1_raised_img.Visible = False
6160        Case False
6170          .cmdCalendar1_raised_img.Visible = True
6180          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6190        End Select
6200        .cmdCalendar1_raised_focus_img.Visible = False
6210        .cmdCalendar1_raised_focus_dots_img.Visible = False
6220        .cmdCalendar1_sunken_focus_dots_img.Visible = False
6230        .cmdCalendar1_raised_img_dis.Visible = False
6240      End If
6250      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
6260        Select Case blnCalendar2_Focus
            Case True
6270          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
6280          .cmdCalendar2_raised_img.Visible = False
6290        Case False
6300          .cmdCalendar2_raised_img.Visible = True
6310          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
6320        End Select
6330        .cmdCalendar2_raised_focus_img.Visible = False
6340        .cmdCalendar2_raised_focus_dots_img.Visible = False
6350        .cmdCalendar2_sunken_focus_dots_img.Visible = False
6360        .cmdCalendar2_raised_img_dis.Visible = False
6370      End If
6380      If .tglSaleReinvest_true_raised_focus_img.Visible = True Or .tglSaleReinvest_true_raised_focus_dots_img.Visible = True Or _
              .tglSaleReinvest_false_raised_focus_img.Visible = True Or .tglSaleReinvest_false_raised_focus_dots_img.Visible = True Then
6390        Select Case .posted
            Case True
6400          Select Case blnSaleReinvest_Focus
              Case True
6410            .tglSaleReinvest_true_raised_focus_dots_img.Visible = True  ' ** Same for ON.
6420            .tglSaleReinvest_true_raised_img.Visible = False
6430          Case False
6440            .tglSaleReinvest_true_raised_img.Visible = True
6450            .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
6460          End Select
6470          .tglSaleReinvest_false_raised_img.Visible = False
6480          .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
6490        Case False
6500          Select Case blnSaleReinvest_Focus
              Case True
6510            .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = True
6520            .tglSaleReinvest_false_raised_img.Visible = False
6530          Case False
6540            .tglSaleReinvest_false_raised_img.Visible = True
6550            .tglSaleReinvest_false_raised_semifocus_dots_img.Visible = False
6560          End Select
6570          .tglSaleReinvest_true_raised_img.Visible = False
6580          .tglSaleReinvest_true_raised_focus_dots_img.Visible = False
6590        End Select
6600        .tglSaleReinvest_false_raised_focus_img.Visible = False
6610        .tglSaleReinvest_false_raised_focus_dots_img.Visible = False
6620        .tglSaleReinvest_false_sunken_focus_dots_img.Visible = False
6630        .tglSaleReinvest_false_raised_img_dis.Visible = False
6640        .tglSaleReinvest_true_raised_focus_img.Visible = False
6650        .tglSaleReinvest_true_sunken_focus_dots_img.Visible = False
6660        .tglSaleReinvest_true_raised_img_dis.Visible = False
6670      End If
6680    End With

EXITP:
6690    Exit Sub

ERRH:
6700    Select Case ERR.Number
        Case Else
6710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6720    End Select
6730    Resume EXITP

End Sub
