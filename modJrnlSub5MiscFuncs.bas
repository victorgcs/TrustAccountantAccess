Attribute VB_Name = "modJrnlSub5MiscFuncs"
Option Compare Database
Option Explicit

'VGC 09/07/2017: CHANGES!

Private Const THIS_NAME As String = "modJrnlSub5MiscFuncs"
' **

Public Sub DetailMouse_Sub5(blnCalendar1_Focus As Boolean, blnMiscMap_LTCL_Focus As Boolean, blnMiscMap_STCGL_Focus As Boolean, blnMiscSale_Focus As Boolean, blnMiscReinvest_Focus As Boolean, blnSpecialCap As Boolean, intSpecialCapOpt As Integer, frm As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "DetailMouse_Sub5"

110     With frm
120       If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
130         Select Case blnCalendar1_Focus
            Case True
140           .cmdCalendar1_raised_semifocus_dots_img.Visible = True
150           .cmdCalendar1_raised_img.Visible = False
160         Case False
170           .cmdCalendar1_raised_img.Visible = True
180           .cmdCalendar1_raised_semifocus_dots_img.Visible = False
190         End Select
200         .cmdCalendar1_raised_focus_img.Visible = False
210         .cmdCalendar1_raised_focus_dots_img.Visible = False
220         .cmdCalendar1_sunken_focus_dots_img.Visible = False
230         .cmdCalendar1_raised_img_dis.Visible = False
240       End If
250       If gblnGoToReport = False Then
260         If blnSpecialCap = True And ((gblnAdmin = True) Or (gblnAdmin = False And intSpecialCapOpt <> 2)) Then
270           If .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = True Or .cmdMiscMap_LTCL_raised_focus_img.Visible = True Then
280             Select Case blnMiscMap_LTCL_Focus
                Case True
290               .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = True
300               .cmdMiscMap_LTCL_raised_img.Visible = False
310             Case False
320               .cmdMiscMap_LTCL_raised_img.Visible = True
330               .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
340             End Select
350             .cmdMiscMap_LTCL_raised_focus_img.Visible = False
360             .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
370             .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
380             .cmdMiscMap_LTCL_raised_img_dis.Visible = False
390           End If
400           If .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = True Or .cmdMiscMap_STCGL_raised_focus_img.Visible = True Then
410             Select Case blnMiscMap_STCGL_Focus
                Case True
420               .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = True
430               .cmdMiscMap_STCGL_raised_img.Visible = False
440             Case False
450               .cmdMiscMap_STCGL_raised_img.Visible = True
460               .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
470             End Select
480             .cmdMiscMap_STCGL_raised_focus_img.Visible = False
490             .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
500             .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
510             .cmdMiscMap_STCGL_raised_img_dis.Visible = False
520           End If
530         End If  ' ** blnSpecialCap.
540       End If
550       If .tglMiscSale_true_raised_focus_img.Visible = True Or .tglMiscSale_true_raised_focus_dots_img.Visible = True Or _
              .tglMiscSale_false_raised_focus_img.Visible = True Or .tglMiscSale_false_raised_focus_dots_img.Visible = True Then
560         Select Case .posted
            Case True
570           Select Case blnMiscSale_Focus
              Case True
580             .tglMiscSale_true_raised_focus_dots_img.Visible = True  ' ** Same for ON.
590             .tglMiscSale_true_raised_img.Visible = False
600           Case False
610             .tglMiscSale_true_raised_img.Visible = True
620             .tglMiscSale_true_raised_focus_dots_img.Visible = False
630           End Select
640           .tglMiscSale_false_raised_img.Visible = False
650           .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
660         Case False
670           Select Case blnMiscSale_Focus
              Case True
680             .tglMiscSale_false_raised_semifocus_dots_img.Visible = True
690             .tglMiscSale_false_raised_img.Visible = False
700           Case False
710             .tglMiscSale_false_raised_img.Visible = True
720             .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
730           End Select
740           .tglMiscSale_true_raised_img.Visible = False
750           .tglMiscSale_true_raised_focus_dots_img.Visible = False
760         End Select
770         .tglMiscSale_false_raised_focus_img.Visible = False
780         .tglMiscSale_false_raised_focus_dots_img.Visible = False
790         .tglMiscSale_false_sunken_focus_dots_img.Visible = False
800         .tglMiscSale_false_raised_img_dis.Visible = False
810         .tglMiscSale_true_raised_focus_img.Visible = False
820         .tglMiscSale_true_sunken_focus_dots_img.Visible = False
830         .tglMiscSale_true_raised_img_dis.Visible = False
840       End If
850       If .tglMiscReinvest_true_raised_focus_img.Visible = True Or .tglMiscReinvest_true_raised_focus_dots_img.Visible = True Or _
              .tglMiscReinvest_false_raised_focus_img.Visible = True Or .tglMiscReinvest_false_raised_focus_dots_img.Visible = True Then
860         Select Case .posted
            Case True
870           Select Case blnMiscReinvest_Focus
              Case True
880             .tglMiscReinvest_true_raised_focus_dots_img.Visible = True  ' ** Same for ON.
890             .tglMiscReinvest_true_raised_img.Visible = False
900           Case False
910             .tglMiscReinvest_true_raised_img.Visible = True
920             .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
930           End Select
940           .tglMiscReinvest_false_raised_img.Visible = False
950           .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
960         Case False
970           Select Case blnMiscReinvest_Focus
              Case True
980             .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = True
990             .tglMiscReinvest_false_raised_img.Visible = False
1000          Case False
1010            .tglMiscReinvest_false_raised_img.Visible = True
1020            .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
1030          End Select
1040          .tglMiscReinvest_true_raised_img.Visible = False
1050          .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
1060        End Select
1070        .tglMiscReinvest_false_raised_focus_img.Visible = False
1080        .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
1090        .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
1100        .tglMiscReinvest_false_raised_img_dis.Visible = False
1110        .tglMiscReinvest_true_raised_focus_img.Visible = False
1120        .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
1130        .tglMiscReinvest_true_raised_img_dis.Visible = False
1140      End If
1150    End With

EXITP:
1160    Exit Sub

ERRH:
1170    Select Case ERR.Number
        Case Else
1180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1190    End Select
1200    Resume EXITP

End Sub

Public Sub Calendar_Handler_Sub5(strProc As String, blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean, clsMonthClass As clsMonthCal, frm As Access.Form)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_Sub5"

        Dim strEvent As String, strCtlName As String
        Dim datStartDate As Date, datEndDate As Date
        Dim intNum As Integer
        Dim blnRetVal As Boolean

1310    With frm

1320      strEvent = Mid(strProc, (CharPos(strProc, 1, "_") + 1))  ' ** Module Function: modStringFuncs.
1330      strCtlName = Left(strProc, (CharPos(strProc, 1, "_") - 1))  ' ** Module Function: modStringFuncs.
1340      intNum = Val(Right(strCtlName, 1))

1350      Select Case strEvent
          Case "Click"
1360        Select Case intNum
            Case 1
1370          datStartDate = Date
1380          datEndDate = 0
1390          blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
1400          If blnRetVal = True Then
                ' ** Allow posting up to 1 month into the future.
1410            If datStartDate > DateAdd("m", 1, Date) Then
1420              MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
1430              .miscTransDate = CDate(Format(Date, "mm/dd/yyyy"))
1440            Else
1450              .miscTransDate = datStartDate
1460            End If
1470          Else
1480            .miscTransDate = CDate(Format(Date, "mm/dd/yyyy"))
1490          End If
1500          If .cmbRecurringItems.Visible = True Then
1510            If .cmbRecurringItems.Enabled = True Then
1520              .cmbRecurringItems.SetFocus
1530            Else
1540              .miscICash.SetFocus
1550            End If
1560          ElseIf .miscAssetNo.Visible = True Then
1570            .miscICash.SetFocus
1580          Else
1590            .miscICash.SetFocus
1600          End If
1610        End Select
1620      Case "GotFocus"
1630        Select Case intNum
            Case 1
1640          blnCalendar1_Focus = True
1650          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
1660          .cmdCalendar1_raised_img.Visible = False
1670          .cmdCalendar1_raised_focus_img.Visible = False
1680          .cmdCalendar1_raised_focus_dots_img.Visible = False
1690          .cmdCalendar1_sunken_focus_dots_img.Visible = False
1700          .cmdCalendar1_raised_img_dis.Visible = False
1710        End Select
1720      Case "MouseDown"
1730        Select Case intNum
            Case 1
1740          blnCalendar1_MouseDown = True
1750          .cmdCalendar1_sunken_focus_dots_img.Visible = True
1760          .cmdCalendar1_raised_img.Visible = False
1770          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
1780          .cmdCalendar1_raised_focus_img.Visible = False
1790          .cmdCalendar1_raised_focus_dots_img.Visible = False
1800          .cmdCalendar1_raised_img_dis.Visible = False
1810        End Select
1820      Case "MouseMove"
1830        Select Case intNum
            Case 1
1840          If blnCalendar1_MouseDown = False Then
1850            Select Case blnCalendar1_Focus
                Case True
1860              .cmdCalendar1_raised_focus_dots_img.Visible = True
1870              .cmdCalendar1_raised_focus_img.Visible = False
1880            Case False
1890              .cmdCalendar1_raised_focus_img.Visible = True
1900              .cmdCalendar1_raised_focus_dots_img.Visible = False
1910            End Select
1920            .cmdCalendar1_raised_img.Visible = False
1930            .cmdCalendar1_raised_semifocus_dots_img.Visible = False
1940            .cmdCalendar1_sunken_focus_dots_img.Visible = False
1950            .cmdCalendar1_raised_img_dis.Visible = False
1960          End If
1970        End Select
1980      Case "MouseUp"
1990        Select Case intNum
            Case 1
2000          .cmdCalendar1_raised_focus_dots_img.Visible = True
2010          .cmdCalendar1_raised_img.Visible = False
2020          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
2030          .cmdCalendar1_raised_focus_img.Visible = False
2040          .cmdCalendar1_sunken_focus_dots_img.Visible = False
2050          .cmdCalendar1_raised_img_dis.Visible = False
2060          blnCalendar1_MouseDown = False
2070        End Select
2080      Case "LostFocus"
2090        Select Case intNum
            Case 1
2100          .cmdCalendar1_raised_img.Visible = True
2110          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
2120          .cmdCalendar1_raised_focus_img.Visible = False
2130          .cmdCalendar1_raised_focus_dots_img.Visible = False
2140          .cmdCalendar1_sunken_focus_dots_img.Visible = False
2150          .cmdCalendar1_raised_img_dis.Visible = False
2160          blnCalendar1_Focus = False
2170        End Select
2180      End Select

2190    End With

EXITP:
2200    Exit Sub

ERRH:
2210    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
2220    Case Else
2230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2240    End Select
2250    Resume EXITP

End Sub

Public Sub Map_Handler_Sub5(strProc As String, blnMiscMap_LTCL_Focus As Boolean, blnMiscMap_LTCL_MouseDown As Boolean, blnMiscMap_STCGL_Focus As Boolean, blnMiscMap_STCGL_MouseDown As Boolean, frm As Access.Form)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Map_Handler_Sub5"

        Dim strEvent As String, strCtlName As String
        Dim strDocName As String
        Dim intPos01 As Integer, lngCnt As Long
        Dim varTmp00 As Variant

2310    With frm

2320      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
2330      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
2340      strEvent = Mid(strProc, (intPos01 + 1))
2350      strCtlName = Left(strProc, (intPos01 - 1))

2360      Select Case strEvent
          Case "Click"
2370        Select Case strCtlName
            Case "cmdMiscMap_LTCL"
2380          DoCmd.Hourglass True
2390          DoEvents
2400          gblnSetFocus = True
2410          .Parent.LoadMapDropdowns  ' ** Form Procedure: frmJournal.
2420          DoEvents
2430          strDocName = "frmMap_Misc_LTCL"
2440          DoCmd.OpenForm strDocName, , , , , , "frmJournal"
2450          If gblnGoToReport = True Then
2460            varTmp00 = DMax("[ID]", "journal")  ' ** Save this so we can delete any fake Journal records.
2470            Select Case IsNull(varTmp00)
                Case True
2480              glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
2490            Case False
2500              glngTaxCode_Distribution = varTmp00
2510            End Select
2520            Forms(strDocName).TimerInterval = 100&
2530            .GoToReport_arw_mapltcl_img.Visible = False
2540            .cmdMiscMap.Visible = True
2550            .cmdMiscMap.Enabled = True
2560            DoCmd.Hourglass True  ' ** Make sure it's still running.
2570            DoEvents
2580          End If
2590        Case "cmdMiscMap_STCGL"
2600          DoCmd.Hourglass True
2610          DoEvents
2620          gblnSetFocus = True
2630          .Parent.LoadMapDropdowns  ' ** Form Procedure: frmJournal.
2640          DoEvents
2650          strDocName = "frmMap_Misc_STCGL"
2660          DoCmd.OpenForm strDocName, , , , , , "frmJournal"
2670          If gblnGoToReport = True Then
2680            varTmp00 = DMax("[ID]", "journal")  ' ** Save this so we can delete any fake Journal records.
2690            Select Case IsNull(varTmp00)
                Case True
2700              glngTaxCode_Distribution = 0&  ' ** Borrowing this variable from the Court Reports.
2710            Case False
2720              glngTaxCode_Distribution = varTmp00
2730            End Select
2740            Forms(strDocName).TimerInterval = 100&
2750            .GoToReport_arw_mapstcgl_img.Visible = False
2760            .cmdMiscMap_LTCL.Enabled = True
2770            .cmdMiscMap_LTCL_raised_img.Visible = True
2780            .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
2790            .cmdMiscMap_LTCL_raised_focus_img.Visible = False
2800            .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
2810            .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
2820            .cmdMiscMap_LTCL_raised_img_dis.Visible = False
2830            DoCmd.Hourglass True  ' ** Make sure it's still running.
2840            DoEvents
2850          End If
2860        End Select
2870      Case "GotFocus"
2880        Select Case strCtlName
            Case "cmdMiscMap_LTCL"
2890          blnMiscMap_LTCL_Focus = True
2900          .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = True
2910          .cmdMiscMap_LTCL_raised_img.Visible = False
2920          .cmdMiscMap_LTCL_raised_focus_img.Visible = False
2930          .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
2940          .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
2950          .cmdMiscMap_LTCL_raised_img_dis.Visible = False
2960        Case "cmdMiscMap_STCGL"
2970          blnMiscMap_STCGL_Focus = True
2980          .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = True
2990          .cmdMiscMap_STCGL_raised_img.Visible = False
3000          .cmdMiscMap_STCGL_raised_focus_img.Visible = False
3010          .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
3020          .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
3030          .cmdMiscMap_STCGL_raised_img_dis.Visible = False
3040        End Select
3050      Case "MouseDown"
3060        Select Case strCtlName
            Case "cmdMiscMap_LTCL"
3070          blnMiscMap_LTCL_MouseDown = True
3080          .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = True
3090          .cmdMiscMap_LTCL_raised_img.Visible = False
3100          .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
3110          .cmdMiscMap_LTCL_raised_focus_img.Visible = False
3120          .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
3130          .cmdMiscMap_LTCL_raised_img_dis.Visible = False
3140        Case "cmdMiscMap_STCGL"
3150          blnMiscMap_STCGL_MouseDown = True
3160          .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = True
3170          .cmdMiscMap_STCGL_raised_img.Visible = False
3180          .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
3190          .cmdMiscMap_STCGL_raised_focus_img.Visible = False
3200          .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
3210          .cmdMiscMap_STCGL_raised_img_dis.Visible = False
3220        End Select
3230      Case "MouseMove"
3240        Select Case strCtlName
            Case "cmdMiscMap_LTCL"
3250          If blnMiscMap_LTCL_MouseDown = False Then
3260            Select Case blnMiscMap_LTCL_Focus
                Case True
3270              .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = True
3280              .cmdMiscMap_LTCL_raised_focus_img.Visible = False
3290            Case False
3300              .cmdMiscMap_LTCL_raised_focus_img.Visible = True
3310              .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
3320            End Select
3330            .cmdMiscMap_LTCL_raised_img.Visible = False
3340            .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
3350            .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
3360            .cmdMiscMap_LTCL_raised_img_dis.Visible = False
3370          End If
3380          If .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = True Or .cmdMiscMap_STCGL_raised_focus_img.Visible = True Then
3390            Select Case blnMiscMap_STCGL_Focus
                Case True
3400              .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = True
3410              .cmdMiscMap_STCGL_raised_img.Visible = False
3420            Case False
3430              .cmdMiscMap_STCGL_raised_img.Visible = True
3440              .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
3450            End Select
3460            .cmdMiscMap_STCGL_raised_focus_img.Visible = False
3470            .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
3480            .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
3490            .cmdMiscMap_STCGL_raised_img_dis.Visible = False
3500          End If
3510        Case "cmdMiscMap_STCGL"
3520          If blnMiscMap_STCGL_MouseDown = False Then
3530            Select Case blnMiscMap_STCGL_Focus
                Case True
3540              .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = True
3550              .cmdMiscMap_STCGL_raised_focus_img.Visible = False
3560            Case False
3570              .cmdMiscMap_STCGL_raised_focus_img.Visible = True
3580              .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
3590            End Select
3600            .cmdMiscMap_STCGL_raised_img.Visible = False
3610            .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
3620            .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
3630            .cmdMiscMap_STCGL_raised_img_dis.Visible = False
3640          End If
3650          If .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = True Or .cmdMiscMap_LTCL_raised_focus_img.Visible = True Then
3660            Select Case blnMiscMap_LTCL_Focus
                Case True
3670              .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = True
3680              .cmdMiscMap_LTCL_raised_img.Visible = False
3690            Case False
3700              .cmdMiscMap_LTCL_raised_img.Visible = True
3710              .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
3720            End Select
3730            .cmdMiscMap_LTCL_raised_focus_img.Visible = False
3740            .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
3750            .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
3760            .cmdMiscMap_LTCL_raised_img_dis.Visible = False
3770          End If
3780        End Select
3790      Case "MouseUp"
3800        Select Case strCtlName
            Case "cmdMiscMap_LTCL"
3810          .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = True
3820          .cmdMiscMap_LTCL_raised_img.Visible = False
3830          .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
3840          .cmdMiscMap_LTCL_raised_focus_img.Visible = False
3850          .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
3860          .cmdMiscMap_LTCL_raised_img_dis.Visible = False
3870          blnMiscMap_LTCL_MouseDown = False
3880        Case "cmdMiscMap_STCGL"
3890          .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = True
3900          .cmdMiscMap_STCGL_raised_img.Visible = False
3910          .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
3920          .cmdMiscMap_STCGL_raised_focus_img.Visible = False
3930          .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
3940          .cmdMiscMap_STCGL_raised_img_dis.Visible = False
3950          blnMiscMap_STCGL_MouseDown = False
3960        End Select
3970      Case "LostFocus"
3980        Select Case strCtlName
            Case "cmdMiscMap_LTCL"
3990          .cmdMiscMap_LTCL_raised_img.Visible = True
4000          .cmdMiscMap_LTCL_raised_semifocus_dots_img.Visible = False
4010          .cmdMiscMap_LTCL_raised_focus_img.Visible = False
4020          .cmdMiscMap_LTCL_raised_focus_dots_img.Visible = False
4030          .cmdMiscMap_LTCL_sunken_focus_dots_img.Visible = False
4040          .cmdMiscMap_LTCL_raised_img_dis.Visible = False
4050          blnMiscMap_LTCL_Focus = False
4060        Case "cmdMiscMap_STCGL"
4070          .cmdMiscMap_STCGL_raised_img.Visible = True
4080          .cmdMiscMap_STCGL_raised_semifocus_dots_img.Visible = False
4090          .cmdMiscMap_STCGL_raised_focus_img.Visible = False
4100          .cmdMiscMap_STCGL_raised_focus_dots_img.Visible = False
4110          .cmdMiscMap_STCGL_sunken_focus_dots_img.Visible = False
4120          .cmdMiscMap_STCGL_raised_img_dis.Visible = False
4130          blnMiscMap_STCGL_Focus = False
4140        End Select
4150      End Select

4160    End With

EXITP:
4170    Exit Sub

ERRH:
4180    Select Case ERR.Number
        Case Else
4190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4200    End Select
4210    Resume EXITP

End Sub

Public Sub Tgl_Handler_Sub5(strProc As String, blnMiscSale_Focus As Boolean, blnMiscSale_MouseDown As Boolean, blnMiscReinvest_Focus As Boolean, blnMiscReinvest_MouseDown As Boolean, blnGoToSaleForm As Boolean, blnGoToRecReinvest As Boolean, frm As Access.Form)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "Tgl_Handler_Sub5"

        Dim strEvent As String, strCtlName As String
        Dim intPos01 As Integer, lngCnt As Long

4310    With frm

4320      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
4330      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
4340      strEvent = Mid(strProc, (intPos01 + 1))
4350      strCtlName = Left(strProc, (intPos01 - 1))

4360      Select Case strEvent
          Case "Click"
4370        Select Case strCtlName
            Case "tglMiscSale"
4380          .tglMiscReinvest_false_raised_img_dis.Visible = True  ' ** Should already be this way.
4390          .tglMiscReinvest_false_raised_img.Visible = False
4400          .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
4410          .tglMiscReinvest_false_raised_focus_img.Visible = False
4420          .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
4430          .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
4440          .tglMiscReinvest_true_raised_img.Visible = False
4450          .tglMiscReinvest_true_raised_focus_img.Visible = False
4460          .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
4470          .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
4480          .tglMiscReinvest_true_raised_img_dis.Visible = False
4490          .tglMiscReinvest.Enabled = False
              ' ** Do I need to do anything with the images, or will that be handled by the other events?
4500          Select Case .posted
              Case True  ' ** If it's True, flip to False, and vice versa.
4510            blnGoToSaleForm = False
4520            .posted = False
4530            DoCmd.RunCommand acCmdSaveRecord
4540          Case False
4550            blnGoToSaleForm = True
4560            .posted = True
4570            DoCmd.RunCommand acCmdSaveRecord
4580          End Select
4590          ChangedMisc_Sub5 True, frm  ' ** Module Procedure: modJrnlSub5MiscFuncs.
4600          DoEvents
4610        Case "tglMiscReinvest"
4620          .tglMiscSale_false_raised_img_dis.Visible = True  ' ** Should already be this way.
4630          .tglMiscSale_false_raised_img.Visible = False
4640          .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
4650          .tglMiscSale_false_raised_focus_img.Visible = False
4660          .tglMiscSale_false_raised_focus_dots_img.Visible = False
4670          .tglMiscSale_false_sunken_focus_dots_img.Visible = False
4680          .tglMiscSale_true_raised_img.Visible = False
4690          .tglMiscSale_true_raised_focus_img.Visible = False
4700          .tglMiscSale_true_raised_focus_dots_img.Visible = False
4710          .tglMiscSale_true_sunken_focus_dots_img.Visible = False
4720          .tglMiscSale_true_raised_img_dis.Visible = False
4730          .tglMiscSale.Enabled = False
              ' ** Do I need to do anything with the images, or will that be handled by the other events?
4740          Select Case .posted
              Case True  ' ** If it's True, flip to False, and vice versa.
4750            blnGoToRecReinvest = False
4760            .posted = False
4770            DoCmd.RunCommand acCmdSaveRecord
4780          Case False
4790            blnGoToRecReinvest = True
4800            .posted = True
4810            DoCmd.RunCommand acCmdSaveRecord
4820          End Select
4830          ChangedMisc_Sub5 True, frm  ' ** Module Procedure: modJrnlSub5MiscFuncs.
4840          DoEvents
4850        End Select
4860      Case "GotFocus"
4870        Select Case strCtlName
            Case "tglMiscSale"
4880          blnMiscSale_Focus = True
4890          Select Case .posted
              Case True
4900            .tglMiscSale_true_raised_focus_dots_img.Visible = True
4910            .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
4920          Case False
4930            .tglMiscSale_false_raised_semifocus_dots_img.Visible = True
4940            .tglMiscSale_true_raised_focus_dots_img.Visible = False
4950          End Select
4960          .tglMiscSale_false_raised_img.Visible = False
4970          .tglMiscSale_false_raised_focus_img.Visible = False
4980          .tglMiscSale_false_raised_focus_dots_img.Visible = False
4990          .tglMiscSale_false_sunken_focus_dots_img.Visible = False
5000          .tglMiscSale_false_raised_img_dis.Visible = False
5010          .tglMiscSale_true_raised_img.Visible = False
5020          .tglMiscSale_true_raised_focus_img.Visible = False
5030          .tglMiscSale_true_sunken_focus_dots_img.Visible = False
5040          .tglMiscSale_true_raised_img_dis.Visible = False
5050        Case "tglMiscReinvest"
5060          blnMiscReinvest_Focus = True
5070          Select Case .posted
              Case True
5080            .tglMiscReinvest_true_raised_focus_dots_img.Visible = True
5090            .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
5100          Case False
5110            .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = True
5120            .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
5130          End Select
5140          .tglMiscReinvest_false_raised_img.Visible = False
5150          .tglMiscReinvest_false_raised_focus_img.Visible = False
5160          .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
5170          .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
5180          .tglMiscReinvest_false_raised_img_dis.Visible = False
5190          .tglMiscReinvest_true_raised_img.Visible = False
5200          .tglMiscReinvest_true_raised_focus_img.Visible = False
5210          .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
5220          .tglMiscReinvest_true_raised_img_dis.Visible = False
5230        End Select
5240      Case "MouseDown"
5250        Select Case strCtlName
            Case "tglMiscSale"
5260          blnMiscSale_MouseDown = True
5270          Select Case .posted
              Case True
5280            .tglMiscSale_true_sunken_focus_dots_img.Visible = True
5290            .tglMiscSale_false_sunken_focus_dots_img.Visible = False
5300          Case False
5310            .tglMiscSale_false_sunken_focus_dots_img.Visible = True
5320            .tglMiscSale_true_sunken_focus_dots_img.Visible = False
5330          End Select
5340          .tglMiscSale_false_raised_img.Visible = False
5350          .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
5360          .tglMiscSale_false_raised_focus_img.Visible = False
5370          .tglMiscSale_false_raised_focus_dots_img.Visible = False
5380          .tglMiscSale_false_raised_img_dis.Visible = False
5390          .tglMiscSale_true_raised_img.Visible = False
5400          .tglMiscSale_true_raised_focus_img.Visible = False
5410          .tglMiscSale_true_raised_focus_dots_img.Visible = False
5420          .tglMiscSale_true_raised_img_dis.Visible = False
5430        Case "tglMiscReinvest"
5440          blnMiscReinvest_MouseDown = True
5450          Select Case .posted
              Case True
5460            .tglMiscReinvest_true_sunken_focus_dots_img.Visible = True
5470            .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
5480          Case False
5490            .tglMiscReinvest_false_sunken_focus_dots_img.Visible = True
5500            .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
5510          End Select
5520          .tglMiscReinvest_false_raised_img.Visible = False
5530          .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
5540          .tglMiscReinvest_false_raised_focus_img.Visible = False
5550          .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
5560          .tglMiscReinvest_false_raised_img_dis.Visible = False
5570          .tglMiscReinvest_true_raised_img.Visible = False
5580          .tglMiscReinvest_true_raised_focus_img.Visible = False
5590          .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
5600          .tglMiscReinvest_true_raised_img_dis.Visible = False
5610        End Select
5620      Case "MouseMove"
5630        Select Case strCtlName
            Case "tglMiscSale"
5640          If blnMiscSale_MouseDown = False Then
5650            Select Case .posted
                Case True
5660              Select Case blnMiscSale_Focus
                  Case True
5670                .tglMiscSale_true_raised_focus_dots_img.Visible = True
5680                .tglMiscSale_true_raised_focus_img.Visible = False
5690              Case False
5700                .tglMiscSale_true_raised_focus_img.Visible = True
5710                .tglMiscSale_true_raised_focus_dots_img.Visible = False
5720              End Select
5730              .tglMiscSale_false_raised_focus_img.Visible = False
5740              .tglMiscSale_false_raised_focus_dots_img.Visible = False
5750            Case False
5760              Select Case blnMiscSale_Focus
                  Case True
5770                .tglMiscSale_false_raised_focus_dots_img.Visible = True
5780                .tglMiscSale_false_raised_focus_img.Visible = False
5790              Case False
5800                .tglMiscSale_false_raised_focus_img.Visible = True
5810                .tglMiscSale_false_raised_focus_dots_img.Visible = False
5820              End Select
5830              .tglMiscSale_true_raised_focus_img.Visible = False
5840              .tglMiscSale_true_raised_focus_dots_img.Visible = False
5850            End Select
5860            .tglMiscSale_false_raised_img.Visible = False
5870            .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
5880            .tglMiscSale_false_raised_img_dis.Visible = False
5890            .tglMiscSale_false_sunken_focus_dots_img.Visible = False
5900            .tglMiscSale_true_raised_img.Visible = False
5910            .tglMiscSale_true_sunken_focus_dots_img.Visible = False
5920            .tglMiscSale_true_raised_img_dis.Visible = False
5930          End If
5940        Case "tglMiscReinvest"
5950          If blnMiscReinvest_MouseDown = False Then
5960            Select Case .posted
                Case True
5970              Select Case blnMiscReinvest_Focus
                  Case True
5980                .tglMiscReinvest_true_raised_focus_dots_img.Visible = True
5990                .tglMiscReinvest_true_raised_focus_img.Visible = False
6000              Case False
6010                .tglMiscReinvest_true_raised_focus_img.Visible = True
6020                .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
6030              End Select
6040              .tglMiscReinvest_false_raised_focus_img.Visible = False
6050              .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
6060            Case False
6070              Select Case blnMiscReinvest_Focus
                  Case True
6080                .tglMiscReinvest_false_raised_focus_dots_img.Visible = True
6090                .tglMiscReinvest_false_raised_focus_img.Visible = False
6100              Case False
6110                .tglMiscReinvest_false_raised_focus_img.Visible = True
6120                .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
6130              End Select
6140              .tglMiscReinvest_true_raised_focus_img.Visible = False
6150              .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
6160            End Select
6170            .tglMiscReinvest_false_raised_img.Visible = False
6180            .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
6190            .tglMiscReinvest_false_raised_img_dis.Visible = False
6200            .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
6210            .tglMiscReinvest_true_raised_img.Visible = False
6220            .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
6230            .tglMiscReinvest_true_raised_img_dis.Visible = False
6240          End If
6250        End Select
6260      Case "MouseUp"
6270        Select Case strCtlName
            Case "tglMiscSale"
6280          Select Case .posted
              Case True
6290            .tglMiscSale_true_raised_focus_dots_img.Visible = True
6300            .tglMiscSale_false_raised_focus_dots_img.Visible = False
6310          Case False
6320            .tglMiscSale_false_raised_focus_dots_img.Visible = True
6330            .tglMiscSale_true_raised_focus_dots_img.Visible = False
6340          End Select
6350          .tglMiscSale_false_raised_img.Visible = False
6360          .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
6370          .tglMiscSale_false_raised_focus_img.Visible = False
6380          .tglMiscSale_false_raised_img_dis.Visible = False
6390          .tglMiscSale_false_sunken_focus_dots_img.Visible = False
6400          .tglMiscSale_true_raised_img.Visible = False
6410          .tglMiscSale_true_raised_focus_img.Visible = False
6420          .tglMiscSale_true_sunken_focus_dots_img.Visible = False
6430          .tglMiscSale_true_raised_img_dis.Visible = False
6440          blnMiscSale_MouseDown = False
6450        Case "tglMiscReinvest"
6460          Select Case .posted
              Case True
6470            .tglMiscReinvest_true_raised_focus_dots_img.Visible = True
6480            .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
6490          Case False
6500            .tglMiscReinvest_false_raised_focus_dots_img.Visible = True
6510            .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
6520          End Select
6530          .tglMiscReinvest_false_raised_img.Visible = False
6540          .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
6550          .tglMiscReinvest_false_raised_focus_img.Visible = False
6560          .tglMiscReinvest_false_raised_img_dis.Visible = False
6570          .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
6580          .tglMiscReinvest_true_raised_img.Visible = False
6590          .tglMiscReinvest_true_raised_focus_img.Visible = False
6600          .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
6610          .tglMiscReinvest_true_raised_img_dis.Visible = False
6620          blnMiscReinvest_MouseDown = False
6630        End Select
6640      Case "LostFocus"
6650        Select Case strCtlName
            Case "tglMiscSale"
6660          Select Case .posted
              Case True
6670            .tglMiscSale_true_raised_img.Visible = True
6680            .tglMiscSale_false_raised_img.Visible = False
6690          Case False
6700            .tglMiscSale_false_raised_img.Visible = True
6710            .tglMiscSale_true_raised_img.Visible = False
6720          End Select
6730          .tglMiscSale_false_raised_semifocus_dots_img.Visible = False
6740          .tglMiscSale_false_raised_focus_img.Visible = False
6750          .tglMiscSale_false_raised_focus_dots_img.Visible = False
6760          .tglMiscSale_false_sunken_focus_dots_img.Visible = False
6770          .tglMiscSale_false_raised_img_dis.Visible = False
6780          .tglMiscSale_true_raised_focus_img.Visible = False
6790          .tglMiscSale_true_raised_focus_dots_img.Visible = False
6800          .tglMiscSale_true_sunken_focus_dots_img.Visible = False
6810          .tglMiscSale_true_raised_img_dis.Visible = False
6820          blnMiscSale_Focus = False
6830        Case "tglMiscReinvest"
6840          Select Case .posted
              Case True
6850            .tglMiscReinvest_true_raised_img.Visible = True
6860            .tglMiscReinvest_false_raised_img.Visible = False
6870          Case False
6880            .tglMiscReinvest_false_raised_img.Visible = True
6890            .tglMiscReinvest_true_raised_img.Visible = False
6900          End Select
6910          .tglMiscReinvest_false_raised_semifocus_dots_img.Visible = False
6920          .tglMiscReinvest_false_raised_focus_img.Visible = False
6930          .tglMiscReinvest_false_raised_focus_dots_img.Visible = False
6940          .tglMiscReinvest_false_sunken_focus_dots_img.Visible = False
6950          .tglMiscReinvest_false_raised_img_dis.Visible = False
6960          .tglMiscReinvest_true_raised_focus_img.Visible = False
6970          .tglMiscReinvest_true_raised_focus_dots_img.Visible = False
6980          .tglMiscReinvest_true_sunken_focus_dots_img.Visible = False
6990          .tglMiscReinvest_true_raised_img_dis.Visible = False
7000          blnMiscReinvest_Focus = False
7010        End Select
7020      End Select
7030    End With

EXITP:
7040    Exit Sub

ERRH:
7050    Select Case ERR.Number
        Case Else
7060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7070    End Select
7080    Resume EXITP

End Sub

Public Sub RevCodeMisc_Sub5(frm As Access.Form)
' ** cmbRevenueCodes
' **   RowSource is 0-Based:
' **     Col 0: revcode_ID
' **     Col 1: revcode_DESC
' **     Col 2: revcode_TYPE
' **     Col 3: revcode_TYPE_Code (I/E)
' **     Col 4: taxcode_type
' **     Col 5: taxcode_type_Code (I/D)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "RevCodeMisc_Sub5"

        Dim strRevCode As String, lngTaxcode As Long

7110    With frm
7120  On Error Resume Next
7130      strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
7140  On Error GoTo ERRH
7150  On Error Resume Next
7160      lngTaxcode = Nz(.cmbTaxCodes, 0&)
7170  On Error GoTo ERRH
7180      If IsNull(.miscType) = False Then
7190        Select Case .miscType
            Case "Paid"
              ' ** EXPENSE.
7200          If .cmbRevenueCodes.RowSource <> "qryRevCodeComboE" Then
7210            .cmbRevenueCodes.RowSource = "qryRevCodeComboE"
7220            .cmbRevenueCodes.Requery
7230          End If
7240          If strRevCode = "I" Then
7250            .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
7260          End If
7270        Case "Received"
              ' ** INCOME.
7280          If .cmbRevenueCodes.RowSource <> "qryRevCodeComboI" Then
7290            .cmbRevenueCodes.RowSource = "qryRevCodeComboI"
7300            .cmbRevenueCodes.Requery
7310          End If
7320          If strRevCode = "E" Then
7330            .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
7340          End If
7350        Case "Misc."
              ' ** ALL.
7360          If .cmbRevenueCodes.RowSource <> "qryRevCodeComboIE" Then
7370            .cmbRevenueCodes.RowSource = "qryRevCodeComboIE"
7380            .cmbRevenueCodes.Requery
7390          End If
7400          If IsNull(.cmbRevenueCodes) = True Then
7410            If gblnLinkRevTaxCodes = True Then
7420              If lngTaxcode = 0& Then
7430                .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
7440              Else
7450                If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
7460                  .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
7470                ElseIf .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
7480                  .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
7490                End If
7500              End If
7510            Else
7520              .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
7530            End If
7540          End If
7550        End Select
7560      End If
7570    End With

EXITP:
7580    Exit Sub

ERRH:
7590    Select Case ERR.Number
        Case Else
7600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7610    End Select
7620    Resume EXITP

End Sub

Public Sub TaxCodeMisc_Sub5(frm As Access.Form)
' ** cmbTaxCodes:
' **   RowSource is 0-Based:
' **     Col 0: taxcode
' **     Col 1: taxcode_description
' **     Col 2: taxcode_type
' **     Col 3: taxcode_type_Code (I/D)
' **     Col 4: revcode_TYPE
' **     Col 5: revcode_TYPE_Code (I/E)
' **   BoundColumn is 1-Based:
' **     Col 0: ListIndex

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "TaxCodeMisc_Sub5"

        Dim strRevCode As String, lngTaxcode As Long

7710    With frm
7720  On Error Resume Next
7730      strRevCode = Trim(Nz(.cmbRevenueCodes.Column(3), vbNullString))
7740  On Error GoTo ERRH
7750  On Error Resume Next
7760      lngTaxcode = Nz(.cmbTaxCodes, 0&)
7770  On Error GoTo ERRH
7780      If IsNull(.miscType) = False Then
7790        Select Case .miscType
            Case "Paid"
              ' ** EXPENSE.
7800          If .cmbTaxCodes.RowSource <> "qryTaxCode_03" Then
7810            .cmbTaxCodes.RowSource = "qryTaxCode_03"
7820            .cmbTaxCodes.Requery
7830          End If
7840          If IsNull(.cmbTaxCodes) = True Then
7850            .cmbTaxCodes = 0&
7860          Else
7870            If lngTaxcode > 0& Then
7880              If .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
7890                .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
7900              End If
7910            End If
7920          End If
7930        Case "Received"
              ' ** INCOME.
7940          If .cmbTaxCodes.RowSource <> "qryTaxCode_02" Then
7950            .cmbTaxCodes.RowSource = "qryTaxCode_02"
7960            .cmbTaxCodes.Requery
7970          End If
7980          If IsNull(.cmbTaxCodes) = True Then
7990            .cmbTaxCodes = 0&
8000          Else
8010            If lngTaxcode > 0& Then
8020              If .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
8030                .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
8040              End If
8050            End If
8060          End If
8070        Case "Misc."
8080          If .cmbTaxCodes.RowSource <> "qryTaxCode_05" Then
8090            .cmbTaxCodes.RowSource = "qryTaxCode_05"
8100            .cmbTaxCodes.Requery
8110          End If
8120          If IsNull(.cmbTaxCodes) = True Then
8130            .cmbTaxCodes = 0&
8140          Else
8150            If gblnLinkRevTaxCodes = True Then
8160              If .cmbTaxCodes = 0& Then
8170                If strRevCode = "I" Then
                      ' ** INCOME.
8180                  .cmbTaxCodes = TAXID_INC  ' ** Unspecified Income.
8190                ElseIf strRevCode = "E" Then
                      ' ** EXPENSE.
8200                  .cmbTaxCodes = TAXID_DED  ' ** Unspecified Deduction.
8210                End If
8220              Else
8230                If .cmbTaxCodes.Column(2) = 2 Then  ' ** taxcode_type, Deduction.
8240                  If strRevCode = "I" Then
                        ' ** EXPENSE.
8250                    .cmbRevenueCodes = REVID_EXP  ' ** Unspecified Expense.
8260                  End If
8270                ElseIf .cmbTaxCodes.Column(2) = 1 Then  ' ** taxcode_type, Income.
8280                  If strRevCode = "E" Then
                        ' ** INCOME.
8290                    .cmbRevenueCodes = REVID_INC  ' ** Unspecified Income.
8300                  End If
8310                End If
8320              End If
8330            End If
8340          End If
8350        End Select
8360      End If
8370    End With

EXITP:
8380    Exit Sub

ERRH:
8390    Select Case ERR.Number
        Case Else
8400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8410    End Select
8420    Resume EXITP

End Sub

Public Sub IncludeCurrencyMisc_Sub5(blnShow As Boolean, lngTpp As Long, lngICash_Left As Long, lngPCash_Left As Long, lngCurrID_Left As Long, frm As Access.Form)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "IncludeCurrencyMisc_Sub5"

8510    With frm
8520      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
8530        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
8540      End If
8550      Select Case blnShow
          Case True
8560        .miscCurr_ID.Left = lngCurrID_Left
8570        .miscCurr_ID_lbl.Left = .miscCurr_ID.Left
8580        .miscICash.Left = lngICash_Left
8590        .miscICash_lbl.Left = .miscICash.Left
8600        .miscICash_usd.Left = .miscICash.Left
8610        .miscPCash.Left = lngPCash_Left
8620        .miscPCash_lbl.Left = .miscPCash.Left
8630        .miscPCash_usd.Left = .miscPCash.Left
8640        .MiscXFer_PtoI_lbl.Left = ((.miscICash.Left + .miscICash.Width) + (4& * lngTpp))
8650        .MiscXFer_ItoP_lbl.Left = .MiscXFer_PtoI_lbl.Left
8660      Case False
8670        .miscCurr_ID.Left = .miscCurr_ID_alt_box.Left
8680        .miscCurr_ID_lbl.Left = .miscCurr_ID.Left
8690        .miscICash.Left = .miscICash_alt_box.Left
8700        .miscICash_lbl.Left = .miscICash.Left
8710        .miscICash_usd.Left = .miscICash.Left
8720        .miscPCash.Left = .miscPCash_alt_box.Left
8730        .miscPCash_lbl.Left = .miscPCash.Left
8740        .miscPCash_usd.Left = .miscPCash.Left
8750        .MiscXFer_PtoI_lbl.Left = .miscXFer_PtoI_lbl_alt_box.Left
8760        .MiscXFer_ItoP_lbl.Left = .MiscXFer_PtoI_lbl.Left
8770      End Select
8780    End With

EXITP:
8790    Exit Sub

ERRH:
8800    Select Case ERR.Number
        Case Else
8810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8820    End Select
8830    Resume EXITP

End Sub

Public Sub MemoUpdate_Sub5(lngJrnlID As Long, strAccountNo As String, datTransDate As Date, strMemo_New As String, frm As Access.Form)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "MemoUpdate_Sub5"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strJrnlType As String
        Dim lngRecs As Long
        Dim blnDelete As Boolean

8910    Set dbs = CurrentDb
8920    With dbs
8930      Set rst = .OpenRecordset("tblJournal_Memo", dbOpenDynaset)
8940      With rst
8950        If .BOF = True And .EOF = True Then
8960          lngRecs = 0&
8970        Else
8980          .MoveLast
8990          lngRecs = .RecordCount
9000          .MoveFirst
9010        End If
9020      End With
9030    End With

9040    With frm
9050      strJrnlType = "Paid"
9060      blnDelete = False
9070      If strMemo_New <> vbNullString Then  ' ** This sub only called if strMemo_New <> strMemo_Orig.
9080        With rst
9090          .FindFirst "[Journal_ID] = " & CStr(lngJrnlID)
9100          If .NoMatch = False Then
9110            .Edit
9120            If ![transdate] <> datTransDate Then
9130              ![transdate] = datTransDate
9140            End If
9150            ![JrnlMemo_Memo] = strMemo_New
9160            ![JrnlMemo_DateModified] = Now()
9170            .Update
9180          Else
9190            .AddNew
9200            ![Journal_ID] = lngJrnlID
9210            ![journaltype] = strJrnlType
9220            ![accountno] = strAccountNo
9230            ![transdate] = datTransDate
9240            ![JrnlMemo_Memo] = strMemo_New
9250            ![JrnlMemo_DateModified] = Now()
9260            .Update
9270          End If
9280        End With
9290      Else
9300        blnDelete = True
9310      End If
9320      If blnDelete = True Then
9330        If lngRecs > 0& Then
9340          With rst
9350            .FindFirst "[Journal_ID] = " & CStr(lngJrnlID)
9360            If .NoMatch = False Then
9370              If ![accountno] = strAccountNo Then
                    ' ** I'm checking this as well because the Journal table is so temporary,
                    ' ** and if it's compacted when empty, the sequence will restart.
9380                .Delete
9390              End If
9400            End If
9410          End With
9420        End If
9430      End If
9440    End With

9450    rst.Close
9460    dbs.Close

EXITP:
9470    Set rst = Nothing
9480    Set dbs = Nothing
9490    Exit Sub

ERRH:
9500    Select Case ERR.Number
        Case Else
9510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9520    End Select
9530    Resume EXITP

End Sub

Public Sub ChangedMisc_Sub5(blnChanged As Boolean, frm As Access.Form)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "ChangedMisc_Sub5"

9610    With frm
9620      Select Case blnChanged
          Case True
9630        gblnMiscChanged = True
9640        .NavigationButtons = False
9650        DoCmd.SelectObject acForm, .Parent.Name, False
9660        With .Parent
9670  On Error Resume Next
              ' ** I think the problem is that the parent's focus is on the subform,
              ' ** and it can't move off the subform while we're working in it.
              ' ** When it's the other way around (changing a subform's focus while
              ' ** working on the parent), there's no problem.
              ' ** So, I don't think this'll ever work, shouldn't work,
              ' ** and the command shouldn't be here!
9680          .FocusHolder.SetFocus
9690  On Error GoTo ERRH
9700          .opgJournal.Enabled = False
9710          .opgJournal_optMisc_lbl_box.Visible = True
9720          .cmdSwitch.Enabled = False
9730          .cmdSwitch_raised_img_dis.Visible = True
9740          .cmdSwitch_raised_img.Visible = False
9750          .cmdSwitch_raised_semifocus_dots_img.Visible = False
9760          .cmdSwitch_raised_focus_img.Visible = False
9770          .cmdSwitch_raised_focus_dots_img.Visible = False
9780          .cmdSwitch_sunken_focus_dots_img.Visible = False
9790          .frmJournal_Sub5_Misc.SetFocus
9800        End With
9810        .cmdMiscClose.Enabled = False
9820        .cmdMiscOK.Enabled = True
9830        .cmdMiscCancel.Enabled = True
9840        .Parent.NavVis False  ' ** Form Procedure: frmJournal.
9850      Case False
9860        gblnMiscChanged = False
9870        .NavigationButtons = True
9880        DoCmd.SelectObject acForm, .Parent.Name, False
9890        With .Parent
9900          .opgJournal.Enabled = True
9910          .opgJournal_optMisc_lbl_box.Visible = False
9920          .cmdSwitch.Enabled = True
9930          .cmdSwitch_raised_img.Visible = True
9940          .cmdSwitch_raised_img_dis.Visible = False
9950          .cmdSwitch_raised_semifocus_dots_img.Visible = False
9960          .cmdSwitch_raised_focus_img.Visible = False
9970          .cmdSwitch_raised_focus_dots_img.Visible = False
9980          .cmdSwitch_sunken_focus_dots_img.Visible = False
9990        End With
10000       .cmdMiscClose.Enabled = True
10010       .cmdMiscOK.Enabled = False
10020       .cmdMiscCancel.Enabled = False
10030       .Parent.NavVis True  ' ** Form Procedure: frmJournal.
10040     End Select
10050   End With

EXITP:
10060   Exit Sub

ERRH:
10070   Select Case ERR.Number
        Case Else
10080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10090   End Select
10100   Resume EXITP

End Sub
